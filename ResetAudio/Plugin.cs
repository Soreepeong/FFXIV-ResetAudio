using Dalamud.Game;
using Dalamud.Game.ClientState;
using Dalamud.Game.Command;
using Dalamud.Game.Gui;
using Dalamud.Hooking;
using Dalamud.Interface;
using Dalamud.IoC;
using Dalamud.Logging;
using Dalamud.Plugin;
using Dalamud.Plugin.Ipc;
using Dalamud.Plugin.Ipc.Exceptions;
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
        private class MemorySignatures {
            // When this signature breaks, look for references to the following GUIDs in the game binary.
            // IID_IMMDeviceEnumerator: A95664D2-9614-4F35-A746-DE8DB63617E6
            // CLSID_MMDeviceEnumerator: bcde0395-e52f-467c-8e3d-c4579291692e
            internal static readonly string XivAudioEnumerator_Initialize = string.Join(' ', new string[]{
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

            internal static readonly string AudioRenderClientThreadBody = string.Join(' ', new string[]{
                // Function preamble
                /* 0x00 */ "40 56"                      , // PUSH rsi
                /* 0x02 */ "41 57"                      , // PUSH r15
                /* 0x04 */ "48 81 ec ?? ?? ?? ??"       , // SUB rsp, 0x000000A8
            
                // Stack guard
                /* 0x0B */ "48 8b 05 ?? ?? ?? ??"       , // MOV rax, [ffxiv_dx11.exe+0x1EF5AB0]
                /* 0x12 */ "48 33 c4"                   , // XOR rax, rsp
                /* 0x15 */ "48 89 44 24 ??"             , // MOV [rsp+0x44], rax

                // CoInitializeEx(0, 0)
                /* 0x1A */ "33 d2"                      , // XOR edx, edx
                /* 0x1C */ "33 c9"                      , // XOR ecx, ecx
                /* 0x1E */ "40 32 f6"                   , // XOR sil, sil
                /* 0x21 */ "ff 15 ?? ?? ?? ??"          , // CALL CoInitializeEx
            
                // The event handle we want, which gets set on exit request.
                /* 0x27 */ "48 8b 05 ?? ?? ?? ??"       , // MOV rax, [ffxiv_dx11.exe+0x1EF6E40]
            });

            internal static readonly string MainAudioClass_Construct = string.Join(' ', new string[] {
                "48 89 5c 24 ??"                        , // MOV [rsp+0x08], rbx
                "48 89 6c 24 ??"                        , // MOV [rsp+0x10], rbp
                "48 89 74 24 ??"                        , // MOV [rsp+0x18], rsi
                "57"                                    , // PUSH rdi
                "41 56"                                 , // PUSH r14
                "41 57"                                 , // PUSH r15
                "48 83 ec ??"                           , // SUB rsp, 0x20
                "41 0f b6 f8"                           , // MOVZX edi, r8l
                "0f b6 f2"                              , // MOVZX esi, dl
                "4c 8b f1"                              , // MOV r14, rcx
            });

            internal static readonly string MainAudioClass_Initialize = string.Join(' ', new string[] {
                "48 89 5c 24 ??"                        , // MOV [rsp+0x10], rbx
                "55"                                    , // PUSH rbp
                "56"                                    , // PUSH rsi
                "57"                                    , // PUSH rdi
                "41 56"                                 , // PUSH r14
                "41 57"                                 , // PUSH r15
                "48 8d ac 24 ?? ?? ?? ??"               , // LEA rbp, [rsp-0x4f0]
                "48 81 ec ?? ?? ?? ??"                  , // SUB rsp, 0x5f0
                "48 8b 05 ?? ?? ?? ??"                  , // MOV rax, [ffxiv_dx11.exe+0x1EF5AB0]
                "48 33 c4"                              , // XOR rax, rsp
                "48 89 85 ?? ?? ?? ??"                  , // MOV [rbp+0x4e0], rax
                "48 8b 05 ?? ?? ?? ??"                  , // MOV rax, [ffxiv_dx11.exe+0x1F12FD8]
                "4d 8b c8"                              , // MOV r9, r8
                "0f b6 f2"                              , // MOVZX esi, dl
            });

            internal static readonly string MainAudioClass_Cleanup = string.Join(' ', new string[] {
                "48 89 5c 24 ??"                        , // MOV [rsp+0x08], rbx
                "48 89 6c 24 ??"                        , // MOV [rsp+0x10], rbp
                "48 89 74 24 ??"                        , // MOV [rsp+0x18], rsi
                "57"                                    , // PUSH rdi
                "48 83 ec ??"                           , // SUB rsp, 0x20
                "48 8b f1"                              , // MOV rci, rcx
                "33 ed"                                 , // XOR ebp, ebp
                "48 8b 89 ?? ?? ?? ??"                  , // MOV rcx, [rcx+0x288]
                "48 85 c9"                              , // TEST rcx, rcx
            });

            internal static readonly string MainAudioClass_SetStaticAddr2 = string.Join(' ', new string[] {
                "48 89 5c 24 ??"                        , // MOV [rsp+0x08], rbx
                "57"                                    , // PUSH rdi
                "48 83 ec ??"                           , // SUB rsp, 0x20
                "33 d2"                                 , // XOR edx, edx
                "48 8b f9"                              , // XOR rdi, rcx
                "45 33 c0"                              , // XOR r8d, r8d
                "8d 4a ??"                              , // LEA ecx, [rdx+0x30]
                "e8 ?? ?? ?? ??"                        , // CALL ffxiv_dx11.exe+0x60c80
                "48 8b d8"                              , // MOV rbx, rax
                "48 85 c0"                              , // TEST rax, rax
                "74 ??"                                 , // JE +0x22
                "48 8d 48 ??"                           , // LEA rcx, [rax+0x08]
                "ff 15 ?? ?? ?? ??"                     , // CALL ntdll.RtlInitializeCriticalSection
            });
        }

        public string Name => "Reset Audio";
        private readonly string SlashCommand = "/resetaudio";
        private readonly string SlashCommandHelpMessage = "Manually trigger game audio reset.\n* /resetaudio (r|reset): Reset audio right now.\n* /resetaudio (h|harder): Completely reloads audio.\n* /resetaudio c|configure: Open ResetAudio configuration window.\n* /resetaudio h|help: Print help message.";

        private readonly DalamudPluginInterface _pluginInterface;
        private readonly CommandManager _commandManager;
        private readonly ChatGui _chatGui;
        private readonly Framework _framework;
        private readonly Configuration _config;

        private readonly List<Tuple<IDisposable?, Action?>> _disposableList = new();
        private readonly CancellationTokenSource _disposeToken;

        private readonly IMMDeviceEnumerator _pDeviceEnumerator;

        private readonly IMMNotificationClientVtbl* _pNotificationClientVtbl;
        private readonly IMMNotificationClientVtbl _notificationClientVtblOriginalValue;

        private readonly bool* _resetAudio;
        private readonly ICallGateSubscriber<int, bool> _orchPlaySong;

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

        private DateTime? ReconstructAudioNextActionTimestamp = null;
        private int ReconstructAudioStep = 0;

        private IntPtr* _phAudioRenderClientThreadExitEvent;
        private IntPtr** _ppMainAudioClass;

        private delegate IntPtr MainAudioClass_Cleanup(IntPtr pThis);
        private delegate IntPtr MainAudioClass_Construct(IntPtr pThis, bool p0, bool p1);
        private delegate IntPtr MainAudioClass_Initialize(IntPtr pThis, bool p0, IntPtr pszAudioConfigString);
        private delegate IntPtr MainAudioClass_SetStaticAddr2(IntPtr pThis);

        private MainAudioClass_Cleanup _pfnMainAudioClass_Cleanup;
        private MainAudioClass_Construct _pfnMainAudioClass_Construct;
        private MainAudioClass_Initialize _pfnMainAudioClass_Initialize;
        private MainAudioClass_SetStaticAddr2 _pfnMainAudioClass_SetStaticAddr2;

        private Task? _resetAudioSoonTask;

        public Plugin(
            [RequiredVersion("1.0")] DalamudPluginInterface pluginInterface,
            [RequiredVersion("1.0")] CommandManager commandManager,
            [RequiredVersion("1.0")] ClientState clientState,
            [RequiredVersion("1.0")] ChatGui chatGui,
            [RequiredVersion("1.0")] Framework framework,
            [RequiredVersion("1.0")] SigScanner sigScanner) {
            try {
                _pluginInterface = pluginInterface;
                _commandManager = commandManager;
                _chatGui = chatGui;
                _framework = framework;

                _disposableList.Add(Tuple.Create<IDisposable?, Action?>(_disposeToken = new(), null));

                _config = _pluginInterface.GetPluginConfig() as Configuration ?? new Configuration();
                _config.Initialize(_pluginInterface);

                _pluginInterface.UiBuilder.Draw += DrawUI;
                _pluginInterface.UiBuilder.OpenConfigUi += () => { _config.ConfigVisible = !_config.ConfigVisible; };

                _orchPlaySong = _pluginInterface.GetIpcSubscriber<int, bool>("Orch.PlaySong");

                _framework.Update += OnFrameworkUpdate;
                _disposableList.Add(Tuple.Create<IDisposable?, Action?>(null, () => { _framework.Update -= OnFrameworkUpdate; }));

                var pAudioRenderClientThreadBody = sigScanner.ScanText(MemorySignatures.AudioRenderClientThreadBody);
                var pOpBase = pAudioRenderClientThreadBody + 0x2E;
                _phAudioRenderClientThreadExitEvent = (IntPtr*)(pOpBase + Marshal.ReadInt32(pOpBase - 0x04));
                PluginLog.Verbose($"phAudioRenderClientThreadExitEvent: {MainModuleRva((IntPtr)_phAudioRenderClientThreadExitEvent)}");

                var pfn = sigScanner.ScanText(MemorySignatures.MainAudioClass_Construct);
                _pfnMainAudioClass_Construct = Marshal.GetDelegateForFunctionPointer<MainAudioClass_Construct>(pfn);
                PluginLog.Verbose($"MainAudioClass_Construct: {MainModuleRva(pfn)}");

                pfn = sigScanner.ScanText(MemorySignatures.MainAudioClass_Initialize);
                _pfnMainAudioClass_Initialize = Marshal.GetDelegateForFunctionPointer<MainAudioClass_Initialize>(pfn);
                PluginLog.Verbose($"MainAudioClass_Initialize: {MainModuleRva(pfn)}");

                pfn = sigScanner.ScanText(MemorySignatures.MainAudioClass_Cleanup);
                _pfnMainAudioClass_Cleanup = Marshal.GetDelegateForFunctionPointer<MainAudioClass_Cleanup>(pfn);
                PluginLog.Verbose($"MainAudioClass_Cleanup: {MainModuleRva(pfn)}");

                pfn = sigScanner.ScanText(MemorySignatures.MainAudioClass_SetStaticAddr2);
                _pfnMainAudioClass_SetStaticAddr2 = Marshal.GetDelegateForFunctionPointer<MainAudioClass_SetStaticAddr2>(pfn);
                PluginLog.Verbose($"MainAudioClass_SetStaticAddr2: {MainModuleRva(pfn)}");
                pOpBase = pfn + 0x39;
                _ppMainAudioClass = (IntPtr**)(pOpBase + Marshal.ReadInt32(pOpBase - 0x04));
                PluginLog.Verbose($"ppMainAudioClass: {MainModuleRva((IntPtr)_ppMainAudioClass)}");
                PluginLog.Verbose($"*ppMainAudioClass: {MainModuleRva((IntPtr)(*_ppMainAudioClass))}");
                PluginLog.Verbose($"**ppMainAudioClass: {MainModuleRva(**_ppMainAudioClass)}");

                var pInitXivAudioEnumerator = sigScanner.ScanText(MemorySignatures.XivAudioEnumerator_Initialize);
                pOpBase = pInitXivAudioEnumerator + 0x34;
                _pNotificationClientVtbl = (IMMNotificationClientVtbl*)(pOpBase + Marshal.ReadInt32(pOpBase - 0x04));
                _notificationClientVtblOriginalValue = *_pNotificationClientVtbl;

                PluginLog.Verbose($"IMMNotificationClientVtbl: {MainModuleRva((IntPtr)_pNotificationClientVtbl)}");
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

                VirtualProtect((IntPtr)_pNotificationClientVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), MemoryProtection.PAGE_READWRITE, out var prevMemoryProtection);
                try {
                    _pNotificationClientVtbl->OnDeviceStateChanged = Marshal.GetFunctionPointerForDelegate(_newOnDeviceStateChanged = new(OnDeviceStateChanged));
                    _pNotificationClientVtbl->OnDeviceAdded = Marshal.GetFunctionPointerForDelegate(_newOnDeviceAdded = new(OnDeviceAdded));
                    _pNotificationClientVtbl->OnDeviceRemoved = Marshal.GetFunctionPointerForDelegate(_newOnDeviceRemoved = new(OnDeviceRemoved));
                    _pNotificationClientVtbl->OnDefaultDeviceChanged = Marshal.GetFunctionPointerForDelegate(_newOnDefaultDeviceChanged = new(OnDefaultDeviceChanged));
                    _pNotificationClientVtbl->OnPropertyValueChanged = Marshal.GetFunctionPointerForDelegate(_newOnPropertyValueChanged = new(OnPropertyValueChanged));
                } finally {
                    VirtualProtect((IntPtr)_pNotificationClientVtbl, (UIntPtr)Marshal.SizeOf<IMMNotificationClientVtbl>(), prevMemoryProtection, out _);
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
                    item.Item1?.Dispose();
                    item.Item2?.Invoke();
                } catch (Exception e) {
                    PluginLog.Warning(e, "{0}: Dispose failure", item);
                }
            }
            _disposableList.Clear();
        }

        private delegate int GetBufferDelegate(void* pThis, int numFramesWritten, out IntPtr dataBufferPointer);
        private delegate int ReleaseBufferDelegate(void* pThis, int numFramesWritten, AudioClientBufferFlags bufferFlags);

        private Hook<GetBufferDelegate>? _getBufferHook = null;
        private Hook<ReleaseBufferDelegate>? _releaseBufferHook = null;
        private float* _pSamples = null;

        private int GetBufferDetour(void* pThis, int numFramesWritten, out IntPtr dataBufferPointer) {
            var res = _getBufferHook!.Original(pThis, numFramesWritten, out dataBufferPointer);
            _pSamples = (float*)dataBufferPointer;
            PluginLog.Information("GetBufferDetour({0:X}): {1:X}, {2:X}", (IntPtr)pThis, numFramesWritten, dataBufferPointer);
            return res;
        }

        private int ReleaseBufferDetour(void* pThis, int numFramesWritten, AudioClientBufferFlags bufferFlags) {
            float minv = 1, maxv = -1;
            for (int i = 0; i < numFramesWritten * 2; i++) {
                _pSamples[i] *= 2;
                minv = Math.Min(minv, _pSamples[i]);
                maxv = Math.Max(maxv, _pSamples[i]);
            }
            PluginLog.Information("ReleaseBufferDetour({0:X}): {1:X}, {2}: {3} ~ {4}", (IntPtr)pThis, numFramesWritten, bufferFlags, minv, maxv);
            return _releaseBufferHook!.Original(pThis, numFramesWritten, bufferFlags);
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
                windowSize = new Vector2(500, 155) * scale;
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

                    if (ImGui.Button("Try Harder"))
                        ReconstructAudio();

                    ImGui.SameLine();
                    ImGui.TextUnformatted("This may crash, and restarting after this is still recommended.");

                    ImGuiHelpers.ScaledDummy(10);

                    var changed = false;
                    changed |= ImGui.Checkbox("Log audio reset notice message to default log channel", ref _config.PrintAudioResetToChat);
                    changed |= ImGui.Checkbox("Show advanced configuration", ref _config.AdvanceConfigExpanded);

                    if (_config.AdvanceConfigExpanded) {
                        ImGuiHelpers.ScaledDummy(10);

                        changed |= ImGui.Checkbox("Orchestrion Integration", ref _config.EnableOrchestrionIntegration);
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

        private void ReconstructAudio() {
            if (ReconstructAudioNextActionTimestamp != null)
                return;

            ReconstructAudioNextActionTimestamp = DateTime.Now;
            ReconstructAudioStep = 0;
            _chatGui.Print(string.Format("[{0}] Completely reloading audio.\n* Your sound settings from System Settings will not take effect until you change respective options again.\n* Your background music may stop playing until it changes.\n* Part of game audio may stop working until restart.\n* Restarting the game is still recommended.", Name));
        }

        private void OnFrameworkUpdate(Framework framework) {
            if (ReconstructAudioNextActionTimestamp == null || ReconstructAudioNextActionTimestamp >= DateTime.Now)
                return;

            if (_config.PrintAudioResetToChat)
                _chatGui.Print(string.Format("[{0}] Completely reloading audio. (Step {1})", Name, ReconstructAudioStep + 1));

            if (ReconstructAudioStep == 0) {
                ReconstructAudioNextActionTimestamp = DateTime.Now.AddMilliseconds(300);
                ReconstructAudioStep = 1;

                SetEvent(*_phAudioRenderClientThreadExitEvent);
                _pfnMainAudioClass_Cleanup(**_ppMainAudioClass);
                PluginLog.Information("Cleanup");

            } else if (ReconstructAudioStep == 1) {
                if (_config.EnableOrchestrionIntegration) {
                    ReconstructAudioNextActionTimestamp = DateTime.Now.AddMilliseconds(100);
                    ReconstructAudioStep = 2;
                } else {
                    ReconstructAudioNextActionTimestamp = null;
                    ReconstructAudioStep = 0;
                }

                _pfnMainAudioClass_Construct(**_ppMainAudioClass, true, false);
                PluginLog.Information("Construct");
                _pfnMainAudioClass_Initialize(**_ppMainAudioClass, false, IntPtr.Zero);
                PluginLog.Information("Initialize");
                _pfnMainAudioClass_SetStaticAddr2(**_ppMainAudioClass);
                PluginLog.Information("SetStaticAddr2");

            } else if (ReconstructAudioStep == 2) {
                ReconstructAudioNextActionTimestamp = DateTime.Now.AddMilliseconds(100);
                ReconstructAudioStep = 3;

                try {
                    if (_config.EnableOrchestrionIntegration)
                        _orchPlaySong.InvokeFunc(1);
                } catch (IpcNotReadyError) {
                    // pass
                }

            } else if (ReconstructAudioStep == 3) {
                ReconstructAudioNextActionTimestamp = null;
                ReconstructAudioStep = 0;

                try {
                    if (_config.EnableOrchestrionIntegration)
                        _orchPlaySong.InvokeFunc(0);
                } catch (IpcNotReadyError) {
                    // pass
                }
            }
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

            } else if (arguments.Length > 0 && arguments.Length <= 6 && "harder"[..arguments.Length] == arguments) {
                ReconstructAudio();

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
                dev = _pDeviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole);
                return dev.GetId() == deviceId;

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
            var device = _pDeviceEnumerator.GetDevice(deviceId);
            try {
                var properties = device.OpenPropertyStore(StorageAccessMode.Read);
                try {
                    return PropVariantToObjectAndFree<T>(properties.GetValue(ref propertyKey));
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

        [DllImport("kernel32.dll")]
        private static extern bool SetEvent(IntPtr hEvent);
    }
}
