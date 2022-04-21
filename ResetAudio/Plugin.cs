using Dalamud;
using Dalamud.Game;
using Dalamud.Game.ClientState;
using Dalamud.Game.Command;
using Dalamud.IoC;
using Dalamud.Logging;
using Dalamud.Plugin;
using ImGuiNET;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;

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

        private static readonly HashSet<PropertyKey> IGNORE_PROPERTY_KEYS = new() {
            // I have no idea when is this property change triggered but whatever
            new PropertyKey() {
                FmtId = new Guid(0x9855c4cd, 0xdf8c, 0x449c, 0xa1, 0x81, 0x81, 0x91, 0xb6, 0x8b, 0xd0, 0x6c),
                Pid = 0
            }
        };

        public string Name => "Reset Audio";
        private readonly string SlashCommand = "/resetaudio";

        private readonly DalamudPluginInterface _pluginInterface;
        private readonly CommandManager _commandManager;
        private readonly Configuration _config;

        private readonly List<IDisposable> _disposableList = new();

        private readonly IMMDeviceEnumerator _pDeviceEnumerator;

        private IMMNotificationClientVtbl* _pNotificationClientVtbl = null;
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

        public Plugin(
            [RequiredVersion("1.0")] DalamudPluginInterface pluginInterface,
            [RequiredVersion("1.0")] CommandManager commandManager,
            [RequiredVersion("1.0")] ClientState clientState,
            [RequiredVersion("1.0")] SigScanner sigScanner) {
            try {
                _pluginInterface = pluginInterface;
                _commandManager = commandManager;

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

                commandManager.AddHandler(SlashCommand, new((_, _) => ResetAudioNow()) { });
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
                    PluginLog.Warning(e, "Dispose failure");
                }
            }
            _disposableList.Clear();
        }

        private void DrawUI() {
            ImGui.SetNextWindowSize(new Vector2(400, 640), ImGuiCond.Once);

            if (_config.ConfigVisible) {
                if (ImGui.Begin("ResetAudio", ref _config.ConfigVisible, ImGuiWindowFlags.AlwaysAutoResize | ImGuiWindowFlags.NoResize | ImGuiWindowFlags.NoCollapse))
                    try {
                        if (ImGui.Button("Reset Audio Now")) {
                            ResetAudioNow();
                        }
                    } finally { ImGui.End(); }
                else {
                    Save();
                }
            }
        }

        private void ResetAudioNow() {
            PluginLog.Information("Resetting audio NOW!");
            *_resetAudio = true;
        }

        private int OnDeviceStateChanged(IMMNotificationClientVtbl* pNotificationClient, string deviceId, uint dwNewState) {
            PluginLog.Information("OnDeviceStateChanged({0}, {1})", GetDeviceFriendlyName(deviceId), dwNewState);
            return _originalOnDeviceStateChanged(pNotificationClient, deviceId, dwNewState);
        }

        private int OnDeviceAdded(IMMNotificationClientVtbl* pNotificationClient, string deviceId) {
            PluginLog.Information("OnDeviceAdded({0})", GetDeviceFriendlyName(deviceId));
            return _originalOnDeviceAdded(pNotificationClient, deviceId);
        }

        private int OnDeviceRemoved(IMMNotificationClientVtbl* pNotificationClient, string deviceId) {
            PluginLog.Information("OnDeviceRemoved({0})", GetDeviceFriendlyName(deviceId));
            return _originalOnDeviceRemoved(pNotificationClient, deviceId);
        }

        private int OnDefaultDeviceChanged(IMMNotificationClientVtbl* pNotificationClient, EDataFlow flow, ERole role, string defaultDeviceId) {
            PluginLog.Information("OnDefaultDeviceChanged({0}, {1}, {2})", flow, role, GetDeviceFriendlyName(defaultDeviceId));
            return _originalOnDefaultDeviceChanged(pNotificationClient, flow, role, defaultDeviceId);
        }

        private int OnPropertyValueChanged(IMMNotificationClientVtbl* pNotificationClient, string deviceId, PropertyKey propertyKey) {
            PluginLog.Information("OnPropertyValueChanged({0}, {1})", GetDeviceFriendlyName(deviceId), PropertyKeyToName(propertyKey) ?? propertyKey.ToString());
            if (!IGNORE_PROPERTY_KEYS.Contains(propertyKey)) {
                if (0 > _pDeviceEnumerator.GetDefaultAudioEndpoint(EDataFlow.eRender, ERole.eConsole, out var dev)) {
                    ResetAudioNow();
                } else {
                    if (0 > dev.GetId(out var defaultDeviceId) || defaultDeviceId == deviceId)
                        ResetAudioNow();

                    Marshal.ReleaseComObject(dev);
                }
            }
            return _originalOnPropertyValueChanged(pNotificationClient, deviceId, propertyKey);
        }

        private string GetDeviceFriendlyName(string deviceId) {
            return GetDeviceProperty<string>(deviceId, PKEY_Device_FriendlyName) ?? $"Unknown({deviceId})";
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

        private static T? PropVariantToObject<T>(ref PropVariant propVariant) {
            object? obj = null;
            var pVariant = Marshal.AllocCoTaskMem(24); // sizeof(VARIANT)
            VariantInit(pVariant);
            if (0 <= PropVariantToVariant(ref propVariant, pVariant)) {
                obj = Marshal.GetObjectForNativeVariant(pVariant);
                PropVariantClear(ref propVariant);
            }
            VariantClear(pVariant);
            return (T?)obj;
        }

        private static string? PropertyKeyToName(PropertyKey propKey) {
            if (0 > PSGetNameFromPropertyKey(ref propKey, out var ptr))
                return null;

            var name = Marshal.PtrToStringUni(ptr);
            Marshal.FreeCoTaskMem(ptr);
            return name;
        }

        private static string MainModuleRva(IntPtr ptr) {
            return $"[ffxiv_dx11.exe+0x{(long)ptr - (long)Process.GetCurrentProcess().MainModule!.BaseAddress:X}]";
        }

#pragma warning disable 0649

        private enum EDataFlow : uint {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count,
        }

        private enum ERole : uint {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count,
        }

        [Flags]
        private enum AudioEndpointState : uint {
            Active = 0x1,
            Disabled = 0x2,
            NotPresent = 0x4,
            Unplugged = 0x8,
            All = 0x0F,
        }

        private enum StorageAccessMode : uint {
            Read,
            Write,
            ReadWrite,
        }

        // https://referencesource.microsoft.com/#PresentationCore/Core/CSharp/system/windows/Media/Imaging/PropVariant.cs
        [StructLayout(LayoutKind.Sequential, Pack = 0)]
        private struct PropArray {
            public UInt32 cElems;
            public IntPtr pElems;
        }

        [StructLayout(LayoutKind.Explicit, Pack = 1)]
        private struct PropVariant {
            [FieldOffset(0)] public ushort varType;
            [FieldOffset(2)] public ushort wReserved1;
            [FieldOffset(4)] public ushort wReserved2;
            [FieldOffset(6)] public ushort wReserved3;

            [FieldOffset(8)] public byte bVal;
            [FieldOffset(8)] public sbyte cVal;
            [FieldOffset(8)] public ushort uiVal;
            [FieldOffset(8)] public short iVal;
            [FieldOffset(8)] public UInt32 uintVal;
            [FieldOffset(8)] public Int32 intVal;
            [FieldOffset(8)] public UInt64 ulVal;
            [FieldOffset(8)] public Int64 lVal;
            [FieldOffset(8)] public float fltVal;
            [FieldOffset(8)] public double dblVal;
            [FieldOffset(8)] public short boolVal;
            [FieldOffset(8)] public IntPtr pclsidVal; //this is for GUID ID pointer
            [FieldOffset(8)] public IntPtr pszVal; //this is for ansi string pointer
            [FieldOffset(8)] public IntPtr pwszVal; //this is for Unicode string pointer
            [FieldOffset(8)] public IntPtr punkVal; //this is for punkVal (interface pointer)
            [FieldOffset(8)] public PropArray ca;
            [FieldOffset(8)] public System.Runtime.InteropServices.ComTypes.FILETIME filetime;
        }

        [Guid("886d8eeb-8cf2-4446-8d02-cdba1dbdcf99"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyStore {
            public int GetCount(out int propCount);

            public int GetAt(int property, out PropertyKey key);

            public int GetValue(ref PropertyKey key, out PropVariant value);

            public int SetValue(ref PropertyKey key, ref PropVariant value);

            public int Commit();
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice {
            // activationParams is a propvariant
            public int Activate(ref Guid id, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object interfacePointer);

            public int OpenPropertyStore(StorageAccessMode stgmAccess, out IPropertyStore properties);

            public int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);

            public int GetState(out AudioEndpointState state);
        }

        [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceCollection {
            public int GetCount(out int numDevices);

            public int Item(int deviceNumber, out IMMDevice device);
        }

        [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator {
            public int EnumAudioEndpoints(EDataFlow dataFlow, AudioEndpointState stateMask, out IMMDeviceCollection devices);

            public int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice device);

            public int GetDevice(string id, out IMMDevice device);

            public int RegisterEndpointNotificationCallback(void* client);

            public int UnregisterEndpointNotificationCallback(void* client);
        }

        [ComImport, Guid("bcde0395-e52f-467c-8e3d-c4579291692e")]
        private class MMDeviceEnumerator {
        }

        private delegate int OnDeviceStateChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId, uint dwNewState);
        private delegate int OnDeviceAddedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId);
        private delegate int OnDeviceRemovedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId);
        private delegate int OnDefaultDeviceChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, EDataFlow flow, ERole role, [MarshalAs(UnmanagedType.LPWStr)] string defaultDeviceId);
        private delegate int OnPropertyValueChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId, PropertyKey propertyKey);

        private struct PropertyKey {
            public Guid FmtId;
            public uint Pid;

            public override string ToString() {
                return $"PropertyKey(GUID={FmtId}, PID={Pid})";
            }
        }

        private static PropertyKey PKEY_Device_FriendlyName = new() {
            FmtId = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0),
            Pid = 14
        };

        private struct IMMNotificationClientVtbl {
            public IntPtr QueryInterface;
            public IntPtr AddRef;
            public IntPtr Release;
            public IntPtr OnDeviceStateChanged;
            public IntPtr OnDeviceAdded;
            public IntPtr OnDeviceRemoved;
            public IntPtr OnDefaultDeviceChanged;
            public IntPtr OnPropertyValueChanged;
        }

        [Flags]
        public enum MemoryProtection : uint {
            PAGE_NOACCESS = 0x01,
            PAGE_READONLY = 0x02,
            PAGE_READWRITE = 0x04,
            PAGE_WRITECOPY = 0x08,
            PAGE_EXECUTE = 0x10,
            PAGE_EXECUTE_READ = 0x20,
            PAGE_EXECUTE_READWRITE = 0x40,
            PAGE_GUARD = 0x100,
            PAGE_NOCACHE = 0x200,
            MEM_WRITECOMBINE = 0x400,
            PAGE_TARGETS_INVALID = 0x40000000,
            PAGE_TARGETS_NO_UPDATE = 0x40000000,
        }

        [DllImport("kernel32.dll")]
        private static extern bool VirtualProtect(IntPtr lpAddress, UIntPtr dwSize, MemoryProtection flNewProtect, out MemoryProtection lpflOldProtect);

        [Flags]
        private enum CLSCTX : uint {
            INPROC_SERVER = 0x1,
            INPROC_HANDLER = 0x2,
            LOCAL_SERVER = 0x4,
            INPROC_SERVER16 = 0x8,
            REMOTE_SERVER = 0x10,
            INPROC_HANDLER16 = 0x20,
            RESERVED1 = 0x40,
            RESERVED2 = 0x80,
            RESERVED3 = 0x100,
            RESERVED4 = 0x200,
            NO_CODE_DOWNLOAD = 0x400,
            RESERVED5 = 0x800,
            NO_CUSTOM_MARSHAL = 0x1000,
            ENABLE_CODE_DOWNLOAD = 0x2000,
            NO_FAILURE_LOG = 0x4000,
            DISABLE_AAA = 0x8000,
            ENABLE_AAA = 0x10000,
            FROM_DEFAULT_CONTEXT = 0x20000,
            ACTIVATE_32_BIT_SERVER = 0x40000,
            ACTIVATE_64_BIT_SERVER = 0x80000,
            INPROC = INPROC_SERVER | INPROC_HANDLER,
            SERVER = INPROC_SERVER | LOCAL_SERVER | REMOTE_SERVER,
            ALL = SERVER | INPROC_HANDLER
        }

        [DllImport("ole32.dll")]
        private static extern int CoCreateInstance(
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter,
            CLSCTX dwClsCtx,
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [In, Out] ref IntPtr ppv);

        [DllImport("Propsys.dll", SetLastError = true)]
        private static extern int PropVariantToVariant(ref PropVariant propVariant, IntPtr pOutVariant);

        [DllImport("Propsys.dll", SetLastError = true)]
        private static extern int PSGetNameFromPropertyKey(ref PropertyKey key, out IntPtr ppszCanonicalName);

        [DllImport("ole32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int PropVariantClear(ref PropVariant propVariant);

        [DllImport("oleaut32.dll", SetLastError = true)]
        private static extern void VariantInit(IntPtr pVariant);

        [DllImport("oleaut32.dll", SetLastError = true)]
        private static extern void VariantClear(IntPtr pVariant);
#pragma warning restore 0649
    }
}
