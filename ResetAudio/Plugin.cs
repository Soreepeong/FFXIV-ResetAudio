using Dalamud.Game;
using Dalamud.Game.ClientState;
using Dalamud.Game.Command;
using Dalamud.Game.Gui;
using Dalamud.Interface;
using Dalamud.IoC;
using Dalamud.Logging;
using Dalamud.Plugin;
using ImGuiNET;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using static ResetAudio.CantBelieveItsNotCpp;

namespace ResetAudio {
    public unsafe sealed class Plugin : IDalamudPlugin {
        // When this signature breaks, look for references to the following GUIDs in the game binary.
        // IID_IMMDeviceEnumerator: A95664D2-9614-4F35-A746-DE8DB63617E6
        // CLSID_MMDeviceEnumerator: bcde0395-e52f-467c-8e3d-c4579291692e
        private static readonly string XIV_AUDIO_ENUMERATOR_INITIALIZE_SIGNATURE = string.Join(' ', new string[]{
            // Function preamble
            /* 0x00 */ "48 89 5c 24 ??"             , // MOV [rsp+0x08], rbx
            /* 0x05 */ "48 89 74 24 ??"             , // MOV [rsp+0x10], rsi
            /* 0x0A */ "57"                         , // PUSH rdi
            /* 0x0B */ "48 83 ec ??"                , // SUB rsp, 0x30

            // Save parameter 1 (rcx) to rbx
            // parameter 1: class containing audio enumeration stuff
            /* 0x0F */ "48 8b d9"                   , // MOV rbx, rcx

            // Call a function that does nothing more than to do the following:
            // 1. MOV rax, rcx
            // 2. RET
            /* 0x12 */ "e8 ?? ?? ?? ??"             , // CALL ffxiv_dx11.exe+0x0155CEF0

            // Anything marked as <x> can be placed anywhere below.
            /* 0x17 */ "48 8d 05 ?? ?? ?? ??"       , // LEA rax, [ffxiv_dx11.exe+0x01ACCBA0]     // <1>
            /* 0x1E */ "48 c7 43 ?? 00 00 00 00"    , // MOV [rbx+0x10], 0x00000000               // <x> Initialize pIMMNotificationClient to nullptr
            /* 0x26 */ "48 89 03"                   , // MOV [rbx], rax                           // <1> Assign probably-destructor
            /* 0x29 */ "48 8d 4b ??"                , // LEA rcx, [rbx+0x10]                      // <x> 
            /* 0x2D */ "48 8d 05 ?? ?? ?? ??"       , // LEA rax, [ffxiv_dx11.exe+0x01ACCBA8]     // <2> This contains the value we want.
            /* 0x34 */ "48 89 43 ??"                , // MOV [rbx+0x08], rax                      // <2> Assign IMMNotificationClientVtbl
        });

        public string Name => "Reset Audio";
        private readonly string SlashCommand = "/resetaudio";
        private readonly string SlashCommandHelpMessage = "Manually trigger game audio reset.\n* /resetaudio (r|reset): Reset audio right now.\n* /resetaudio c|configure: Open ResetAudio configuration window.\n* /resetaudio h|help: Print help message.";

        private readonly DalamudPluginInterface _pluginInterface;
        private readonly CommandManager _commandManager;
        private readonly ChatGui _chatGui;
        private readonly Configuration _config;

        private readonly List<IDisposable> _disposableList = new();
        private readonly CancellationTokenSource _disposeToken;

        private readonly IMMDeviceEnumerator _pDeviceEnumerator;

        private readonly IMMNotificationClientVtbl* _pNotificationClientVtbl;
        private readonly IMMNotificationClientVtbl _notificationClientVtblOriginalValue;

        private readonly bool* _resetAudio;

        private readonly OnDeviceStateChangedDelegate _originalOnDeviceStateChanged;
        private readonly OnDeviceAddedDelegate _originalOnDeviceAdded;
        private readonly OnDeviceRemovedDelegate _originalOnDeviceRemoved;
        private readonly OnDefaultDeviceChangedDelegate _originalOnDefaultDeviceChanged;
        private readonly OnPropertyValueChangedDelegate _originalOnPropertyValueChanged;

        private readonly OnDeviceStateChangedDelegate _newOnDeviceStateChanged;
        private readonly OnDeviceAddedDelegate _newOnDeviceAdded;
        private readonly OnDeviceRemovedDelegate _newOnDeviceRemoved;
        private readonly OnDefaultDeviceChangedDelegate _newOnDefaultDeviceChanged;
        private readonly OnPropertyValueChangedDelegate _newOnPropertyValueChanged;

        private readonly Dictionary<PropertyKey, int> _propertyUpdateCount = new();
        private readonly Dictionary<PropertyKey, DateTime> _propertyUpdateLatest = new();

        private Task? _resetAudioSoonTask;

        public Plugin(
            [RequiredVersion("1.0")] DalamudPluginInterface pluginInterface,
            [RequiredVersion("1.0")] CommandManager commandManager,
            [RequiredVersion("1.0")] ClientState clientState,
            [RequiredVersion("1.0")] ChatGui chatGui,
            [RequiredVersion("1.0")] SigScanner sigScanner) {
            try {
                _pluginInterface = pluginInterface;
                _commandManager = commandManager;
                _chatGui = chatGui;

                _disposableList.Add(_disposeToken = new());

                _config = _pluginInterface.GetPluginConfig() as Configuration ?? new Configuration();
                _config.Initialize(_pluginInterface);

                _pluginInterface.UiBuilder.Draw += DrawUI;
                _pluginInterface.UiBuilder.OpenConfigUi += () => { _config.ConfigVisible = !_config.ConfigVisible; };

                var pInitXivAudioEnumerator = sigScanner.ScanText(XIV_AUDIO_ENUMERATOR_INITIALIZE_SIGNATURE);
                var pOpBase = pInitXivAudioEnumerator + 0x34;
                var pAddrDelta = Marshal.ReadInt32(pInitXivAudioEnumerator + 0x2D + 0x03);
                var pVtbl = pOpBase + pAddrDelta;
                _pNotificationClientVtbl = (IMMNotificationClientVtbl*)pVtbl;
                _notificationClientVtblOriginalValue = *_pNotificationClientVtbl;

                PluginLog.Verbose($"IMMNotificationClientVtbl: {MainModuleRva(pVtbl)}");
                PluginLog.Verbose($"OnDeviceStateChanged: {MainModuleRva(_pNotificationClientVtbl->OnDeviceStateChanged)}");
                PluginLog.Verbose($"OnDeviceAdded: {MainModuleRva(_pNotificationClientVtbl->OnDeviceAdded)}");
                PluginLog.Verbose($"OnDeviceRemoved: {MainModuleRva(_pNotificationClientVtbl->OnDeviceRemoved)}");
                PluginLog.Verbose($"OnDefaultDeviceChanged: {MainModuleRva(_pNotificationClientVtbl->OnDefaultDeviceChanged)}");
                PluginLog.Verbose($"OnPropertyValueChanged: {MainModuleRva(_pNotificationClientVtbl->OnPropertyValueChanged)}");

                _originalOnDeviceStateChanged = Marshal.GetDelegateForFunctionPointer<OnDeviceStateChangedDelegate>(_notificationClientVtblOriginalValue.OnDeviceStateChanged);
                _originalOnDeviceAdded = Marshal.GetDelegateForFunctionPointer<OnDeviceAddedDelegate>(_notificationClientVtblOriginalValue.OnDeviceAdded);
                _originalOnDeviceRemoved = Marshal.GetDelegateForFunctionPointer<OnDeviceRemovedDelegate>(_notificationClientVtblOriginalValue.OnDeviceRemoved);
                _originalOnDefaultDeviceChanged = Marshal.GetDelegateForFunctionPointer<OnDefaultDeviceChangedDelegate>(_notificationClientVtblOriginalValue.OnDefaultDeviceChanged);
                _originalOnPropertyValueChanged = Marshal.GetDelegateForFunctionPointer<OnPropertyValueChangedDelegate>(_notificationClientVtblOriginalValue.OnPropertyValueChanged);

                _pDeviceEnumerator = (IMMDeviceEnumerator)new MMDeviceEnumerator();
                _resetAudio = (bool*)(_notificationClientVtblOriginalValue.OnDefaultDeviceChanged + 0x07 + Marshal.ReadInt32(_notificationClientVtblOriginalValue.OnDefaultDeviceChanged + 0x03));

                PluginLog.Log($"ResetAudio Byte: {MainModuleRva((IntPtr)_resetAudio)}");

                VirtualProtect(pVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), MemoryProtection.PAGE_READWRITE, out var prevMemoryProtection);
                try {
                    _pNotificationClientVtbl->OnDeviceStateChanged = Marshal.GetFunctionPointerForDelegate(_newOnDeviceStateChanged = new(OnDeviceStateChanged));
                    _pNotificationClientVtbl->OnDeviceAdded = Marshal.GetFunctionPointerForDelegate(_newOnDeviceAdded = new(OnDeviceAdded));
                    _pNotificationClientVtbl->OnDeviceRemoved = Marshal.GetFunctionPointerForDelegate(_newOnDeviceRemoved = new(OnDeviceRemoved));
                    _pNotificationClientVtbl->OnDefaultDeviceChanged = Marshal.GetFunctionPointerForDelegate(_newOnDefaultDeviceChanged = new(OnDefaultDeviceChanged));
                    _pNotificationClientVtbl->OnPropertyValueChanged = Marshal.GetFunctionPointerForDelegate(_newOnPropertyValueChanged = new(OnPropertyValueChanged));
                } finally {
                    VirtualProtect(pVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), prevMemoryProtection, out _);
                }

                commandManager.AddHandler(SlashCommand, new(OnCommand) {
                    HelpMessage = SlashCommandHelpMessage,
                });
            } catch {
                Dispose();
                throw;
            }
        }

        private void Save() {
            if (_config == null)
                return;

            _config.Save();
        }

        public void Dispose() {
            _disposeToken.Cancel();

            _commandManager.RemoveHandler(SlashCommand);

            if (_pDeviceEnumerator != null)
                Marshal.ReleaseComObject(_pDeviceEnumerator);

            if (_pNotificationClientVtbl != null) {
                VirtualProtect((IntPtr)_pNotificationClientVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), MemoryProtection.PAGE_READWRITE, out var prevMemoryProtection);
                try {
                    *_pNotificationClientVtbl = _notificationClientVtblOriginalValue;
                } finally {
                    VirtualProtect((IntPtr)_pNotificationClientVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), prevMemoryProtection, out _);
                }
            }

            Save();

            foreach (var item in _disposableList.AsEnumerable().Reverse()) {
                try {
                    item.Dispose();
                } catch (Exception e) {
                    PluginLog.Warning(e, "{0}: Dispose failure", item);
                }
            }
            _disposableList.Clear();
        }

        private void DrawUI() {
            if (!_config.ConfigVisible)
                return;

            var scale = ImGui.GetIO().FontGlobalScale;
            var windowFlags = ImGuiWindowFlags.NoCollapse;
            Vector2 windowSize;
            if (_config.AdvanceConfigExpanded) {
                windowSize = new Vector2(720, 320) * scale;
                ImGui.PushStyleVar(ImGuiStyleVar.WindowMinSize, windowSize);
                ImGui.SetNextWindowSize(windowSize, ImGuiCond.Once);

            } else {
                windowSize = new Vector2(450, 135) * scale;
                ImGui.PushStyleVar(ImGuiStyleVar.WindowMinSize, windowSize);
                ImGui.SetNextWindowSize(windowSize, ImGuiCond.Always);
                windowFlags |= ImGuiWindowFlags.NoResize;
            }

            if (ImGui.Begin("ResetAudio Configuration###MainWindow", ref _config.ConfigVisible, windowFlags)) {
                try {
                    if (ImGui.Button("Reset Audio Now"))
                        ResetAudioNow(DeviceResetTriggerReason.UserRequest);

                    ImGui.SameLine();
                    ImGui.TextUnformatted("Use /resetaudio to invoke it as a text command.");

                    ImGuiHelpers.ScaledDummy(10);

                    var changed = false;
                    changed |= ImGui.Checkbox("Log audio reset notice message to default log channel", ref _config.PrintAudioResetToChat);
                    changed |= ImGui.Checkbox("Show advanced configuration", ref _config.AdvanceConfigExpanded);
                    if (_config.AdvanceConfigExpanded) {
                        ImGuiHelpers.ScaledDummy(10);

                        changed |= ImGui.Checkbox("Do not relay audio device change notification to game", ref _config.SuppressMultimediaDeviceChangeNotificationToGame);
                        changed |= ImGui.SliderInt("Merge audio reset requests (ms)", ref _config.CoalesceAudioResetRequestDurationMs, 50, 1000);

                        ImGuiHelpers.ScaledDummy(10);

                        ImGui.TextUnformatted("Do not reset audio when following property changes:");
                        ImGui.SameLine(ImGui.GetWindowContentRegionWidth() - 120 * scale);
                        if (ImGui.Button($"Reset Counter###table#Reset Counter", new(120 * scale, 0))) {
                            _propertyUpdateLatest.Clear();
                            _propertyUpdateCount.Clear();
                        }

                        if (ImGui.BeginTable("table", 6)) {
                            ImGui.TableSetupColumn("", ImGuiTableColumnFlags.WidthFixed, 20 * scale);  // Enable/Disable checkbox
                            ImGui.TableSetupColumn("Property Key", ImGuiTableColumnFlags.WidthFixed, 300 * scale);
                            ImGui.TableSetupColumn("Comment", ImGuiTableColumnFlags.WidthStretch);
                            ImGui.TableSetupColumn("Last Seen", ImGuiTableColumnFlags.WidthFixed);
                            ImGui.TableSetupColumn("Count", ImGuiTableColumnFlags.WidthFixed);
                            ImGui.TableSetupColumn("", ImGuiTableColumnFlags.WidthFixed, 60 * scale);  // Add/Delete button
                            ImGui.TableHeadersRow();

                            for (var i = 0; i < _config.IgnorePropertyUpdateKeys.Count; i++) {
                                var prop = _config.IgnorePropertyUpdateKeys[i];

                                ImGui.TableNextRow();

                                ImGui.TableSetColumnIndex(0);
                                changed |= ImGui.Checkbox($"###table#{prop.PropKey}#Enable", ref prop.Enable);

                                ImGui.TableSetColumnIndex(1);
                                ImGui.TextUnformatted(prop.PropKey.ToString());

                                ImGui.TableSetColumnIndex(2);
                                ImGui.PushItemWidth(ImGui.GetColumnWidth());
                                changed |= ImGui.InputText($"###table#{prop.PropKey}#Comment", ref prop.Comment, 1024);
                                ImGui.PopItemWidth();

                                ImGui.TableSetColumnIndex(3);
                                if (_propertyUpdateLatest.TryGetValue(prop.PropKey, out var value))
                                    ImGui.TextUnformatted($"{FormatTimeAgo(value)}");
                                else
                                    ImGui.TextUnformatted("Never");

                                ImGui.TableSetColumnIndex(4);
                                ImGui.TextUnformatted($"{_propertyUpdateCount.GetValueOrDefault(prop.PropKey, 0)}");

                                ImGui.TableSetColumnIndex(5);
                                if (ImGui.Button($"Delete###table#{prop.PropKey}#Delete", new(ImGui.GetColumnWidth(), 0))) {
                                    _config.IgnorePropertyUpdateKeys.RemoveAt(i);
                                    changed = true;
                                    i--;
                                }
                            }

                            foreach (var key in _propertyUpdateLatest.Keys) {
                                if (_config.IgnorePropertyUpdateKeys.Any(x => x.PropKey == key))
                                    continue;

                                ImGui.TableNextRow();

                                ImGui.TableSetColumnIndex(1);
                                ImGui.TextUnformatted(key.ToString());

                                ImGui.TableSetColumnIndex(3);
                                ImGui.TextUnformatted($"{FormatTimeAgo(_propertyUpdateLatest[key])}");

                                ImGui.TableSetColumnIndex(4);
                                ImGui.TextUnformatted($"{_propertyUpdateCount[key]}");

                                ImGui.TableSetColumnIndex(5);
                                if (ImGui.Button($"Add###table#{key}#Add", new(ImGui.GetColumnWidth(), 0))) {
                                    _config.IgnorePropertyUpdateKeys.Add(new() {
                                        PropKey = key,
                                        Comment = "",
                                        Enable = true,
                                    });
                                    changed = true;
                                }
                            }

                            ImGui.EndTable();
                        }
                    }

                    if (changed)
                        Save();
                } finally { ImGui.End(); }
            } else {
                Save();
            }

            ImGui.PopStyleVar();
        }

        private void OnCommand(string command, string arguments) {
            arguments = arguments.Trim().ToLowerInvariant();
            if (arguments == "") {
                ResetAudioNow(DeviceResetTriggerReason.UserRequest);

            } else if (arguments.Length > 0 && arguments.Length <= 9 && "configure"[..arguments.Length] == arguments) {
                _config.ConfigVisible = true;
                Save();

            } else if (arguments.Length > 0 && arguments.Length <= 5 && "reset"[..arguments.Length] == arguments) {
                ResetAudioNow(DeviceResetTriggerReason.UserRequest);

            } else if (arguments.Length > 0 && arguments.Length <= 4 && "help"[..arguments.Length] == arguments) {
                _chatGui.Print(SlashCommandHelpMessage);

            } else {
                _chatGui.PrintError(string.Format("Invalid argument supplied: \"{0}\". Type \"/resetaudio help\" for help.", arguments));
            }
        }

        private void ResetAudioNow(DeviceResetTriggerReason reason) {
            PluginLog.Information("Resetting audio NOW! (Reason: {0})", reason);

            if (_config.PrintAudioResetToChat)
                _chatGui.Print(string.Format("[{0}] Resetting audio. (Reason: {1})", Name, reason));

            *_resetAudio = true;
        }

        private void ResetAudioSoon(DeviceResetTriggerReason reason) {
            if (_resetAudioSoonTask != null)
                return;

            PluginLog.Information("Resetting audio SOON! (Reason: {0})", reason);

            _resetAudioSoonTask = Task.Run(() => {
                try {
                    Task.Delay(_config.CoalesceAudioResetRequestDurationMs, _disposeToken.Token).Wait();
                    if (_disposeToken.IsCancellationRequested)
                        return;

                    ResetAudioNow(reason);
                } catch (Exception) {
                    // Don't know, don't care
                } finally {
                    _resetAudioSoonTask = null;
                }
            });
        }

        private int OnDeviceStateChanged(IMMNotificationClientVtbl* pNotificationClient, string deviceId, uint dwNewState) {
            PluginLog.Information("OnDeviceStateChanged({0}, {1})", GetDeviceFriendlyName(deviceId), dwNewState);
            return _config.SuppressMultimediaDeviceChangeNotificationToGame ? 0 : _originalOnDeviceStateChanged(pNotificationClient, deviceId, dwNewState);
        }

        private int OnDeviceAdded(IMMNotificationClientVtbl* pNotificationClient, string deviceId) {
            PluginLog.Information("OnDeviceAdded({0})", GetDeviceFriendlyName(deviceId));
            return _config.SuppressMultimediaDeviceChangeNotificationToGame ? 0 : _originalOnDeviceAdded(pNotificationClient, deviceId);
        }

        private int OnDeviceRemoved(IMMNotificationClientVtbl* pNotificationClient, string deviceId) {
            PluginLog.Information("OnDeviceRemoved({0})", GetDeviceFriendlyName(deviceId));
            return _config.SuppressMultimediaDeviceChangeNotificationToGame ? 0 : _originalOnDeviceRemoved(pNotificationClient, deviceId);
        }

        private int OnDefaultDeviceChanged(IMMNotificationClientVtbl* pNotificationClient, EDataFlow flow, ERole role, string? defaultDeviceId) {
            PluginLog.Information("OnDefaultDeviceChanged({0}, {1}, {2})", flow, role, defaultDeviceId == null ? "<null>" : GetDeviceFriendlyName(defaultDeviceId));

            if (defaultDeviceId != null)
                ResetAudioSoon(DeviceResetTriggerReason.DefaultDeviceChange);

            return _config.SuppressMultimediaDeviceChangeNotificationToGame ? 0 : _originalOnDefaultDeviceChanged(pNotificationClient, flow, role, defaultDeviceId);
        }

        private int OnPropertyValueChanged(IMMNotificationClientVtbl* pNotificationClient, string deviceId, PropertyKey propertyKey) {
            PluginLog.Information("OnPropertyValueChanged({0}, {1})", GetDeviceFriendlyName(deviceId), PropertyKeyToName(propertyKey) ?? propertyKey.ToString());

            if (IsDefaultRenderDevice(deviceId)) {
                _propertyUpdateCount[propertyKey] = _propertyUpdateCount.GetValueOrDefault(propertyKey, 0) + 1;
                _propertyUpdateLatest[propertyKey] = DateTime.Now;

                if (!_config.IgnorePropertyUpdateKeys.Any(x => x.PropKey.Equals(propertyKey)))
                    ResetAudioSoon(DeviceResetTriggerReason.DefaultDevicePropertyChange);
            }

            return _config.SuppressMultimediaDeviceChangeNotificationToGame ? 0 : _originalOnPropertyValueChanged(pNotificationClient, deviceId, propertyKey);
        }

        private bool IsDefaultRenderDevice(string deviceId) {
            IMMDevice? dev = null;
            try {
                _pDeviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole, out dev);
                dev.GetId(out var defaultDeviceId);
                return defaultDeviceId == deviceId;

            } catch (COMException ex) {
                if (ex.HResult == unchecked((int)0x80070490)) {
                    // Element not found.
                    // Happens when there does not exist a default audio endpoint.
                    return false;

                } else {
                    PluginLog.Error(ex, "IsDefaultRenderDevice({0}): COM Error occurred", deviceId);
                    return false;
                }

            } catch (Exception ex) {
                PluginLog.Error(ex, "IsDefaultRenderDevice({0}): Error occurred", deviceId);
                return false;

            } finally {
                if (dev != null)
                    Marshal.ReleaseComObject(dev);
            }
        }

        private string GetDeviceFriendlyName(string deviceId) {
            try {
                return GetDeviceProperty<string>(deviceId, PKEY_Device_FriendlyName) ?? $"Unknown({deviceId})";
            } catch (Exception ex) {
                PluginLog.Error(ex, "GetDeviceFriendlyName({0}) failure", deviceId);
                return $"Unknown({deviceId})";
            }
        }

        private T? GetDeviceProperty<T>(string deviceId, PropertyKey propertyKey) {
            if (0 > _pDeviceEnumerator.GetDevice(deviceId, out var device))
                return default(T);
            try {
                if (0 > device.OpenPropertyStore(StorageAccessMode.Read, out var properties))
                    return default(T);
                try {
                    if (0 > properties.GetValue(ref propertyKey, out var propVariant))
                        return default(T);

                    return PropVariantToObject<T>(ref propVariant);
                } finally {
                    Marshal.ReleaseComObject(properties);
                }
            } finally {
                Marshal.ReleaseComObject(device);
            }
        }

        private static string FormatTimeAgo(DateTime time) {
            var secs = (int)(DateTime.Now - time).TotalSeconds;
            if (secs <= 0)
                return "Now";
            if (secs < 60)
                return $"{secs}s ago";
            if (secs < 60 * 60)
                return $"{secs / 60}m ago";
            return $"{secs / 60 / 60}h ago";
        }

        private enum DeviceResetTriggerReason {
            Unknown,
            UserRequest,
            DefaultDeviceChange,
            DefaultDevicePropertyChange,
        }
    }
}
