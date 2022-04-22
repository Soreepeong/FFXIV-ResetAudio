using Dalamud.Logging;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

#pragma warning disable CA1069 // Enums values should not be duplicated
#pragma warning disable CA1401 // P/Invokes should not be visible
#pragma warning disable CA2211 // Non-constant fields should not be visible

namespace ResetAudio {
    public unsafe sealed class CantBelieveItsNotCpp {
        public enum EDataFlow : uint {
            eRender,
            eCapture,
            eAll,
            EDataFlow_enum_count,
        }

        public enum ERole : uint {
            eConsole,
            eMultimedia,
            eCommunications,
            ERole_enum_count,
        }

        [Flags]
        public enum AudioEndpointState : uint {
            Active = 0x1,
            Disabled = 0x2,
            NotPresent = 0x4,
            Unplugged = 0x8,
            All = 0x0F,
        }

        public enum StorageAccessMode : uint {
            Read,
            Write,
            ReadWrite,
        }

        // https://referencesource.microsoft.com/#PresentationCore/Core/CSharp/system/windows/Media/Imaging/PropVariant.cs
        [StructLayout(LayoutKind.Sequential, Pack = 0)]
        public struct PropArray {
            public UInt32 cElems;
            public IntPtr pElems;
        }

        [StructLayout(LayoutKind.Explicit, Pack = 1)]
        public struct PropVariant {
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
        public interface IPropertyStore {
            public int GetCount(out int propCount);

            public int GetAt(int property, out PropertyKey key);

            public int GetValue(ref PropertyKey key, out PropVariant value);

            public int SetValue(ref PropertyKey key, ref PropVariant value);

            public int Commit();
        }

        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDevice {
            // activationParams is a propvariant
            public int Activate(ref Guid id, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object interfacePointer);

            public int OpenPropertyStore(StorageAccessMode stgmAccess, out IPropertyStore properties);

            public int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);

            public int GetState(out AudioEndpointState state);
        }

        [Guid("0BD7A1BE-7A1A-44DB-8397-CC5392387B5E"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDeviceCollection {
            public int GetCount(out int numDevices);

            public int Item(int deviceNumber, out IMMDevice device);
        }

        [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMMDeviceEnumerator {
            public int EnumAudioEndpoints(EDataFlow dataFlow, AudioEndpointState stateMask, out IMMDeviceCollection devices);

            public int GetDefaultAudioEndpoint(EDataFlow dataFlow, ERole role, out IMMDevice device);

            public int GetDevice(string id, out IMMDevice device);

            public int RegisterEndpointNotificationCallback(void* client);

            public int UnregisterEndpointNotificationCallback(void* client);
        }

        [ComImport, Guid("bcde0395-e52f-467c-8e3d-c4579291692e")]
        public class MMDeviceEnumerator {
        }

        public delegate int OnDeviceStateChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId, uint dwNewState);
        public delegate int OnDeviceAddedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId);
        public delegate int OnDeviceRemovedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId);
        public delegate int OnDefaultDeviceChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, EDataFlow flow, ERole role, [MarshalAs(UnmanagedType.LPWStr)] string? defaultDeviceId);
        public delegate int OnPropertyValueChangedDelegate(IMMNotificationClientVtbl* pNotificationClient, [MarshalAs(UnmanagedType.LPWStr)] string deviceId, PropertyKey propertyKey);

        public struct PropertyKey : IEquatable<PropertyKey> {
            public Guid FmtId;

            public uint Pid;

            public bool Equals(PropertyKey other) => FmtId == other.FmtId && Pid == other.Pid;

            public override bool Equals(object? o) => o is PropertyKey other && Equals(other);

            public override int GetHashCode() => FmtId.GetHashCode() ^ Pid.GetHashCode();

            public override string ToString() => $"{{{FmtId}:{Pid}}}";

            public static bool operator ==(PropertyKey left, PropertyKey right) => left.Equals(right);

            public static bool operator !=(PropertyKey left, PropertyKey right) => !left.Equals(right);
        }

        public static PropertyKey PKEY_Device_FriendlyName = new() {
            FmtId = new Guid(0xa45c254e, 0xdf1c, 0x4efd, 0x80, 0x20, 0x67, 0xd1, 0x46, 0xa8, 0x50, 0xe0),
            Pid = 14
        };

        public struct IMMNotificationClientVtbl {
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
        public static extern bool VirtualProtect(IntPtr lpAddress, UIntPtr dwSize, MemoryProtection flNewProtect, out MemoryProtection lpflOldProtect);

        [Flags]
        public enum CLSCTX : uint {
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
        public static extern int CoCreateInstance(
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            [MarshalAs(UnmanagedType.IUnknown)] object pUnkOuter,
            CLSCTX dwClsCtx,
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid riid,
            [In, Out] ref IntPtr ppv);

        [DllImport("Propsys.dll", SetLastError = true)]
        public static extern int PropVariantToVariant(ref PropVariant propVariant, IntPtr pOutVariant);

        [DllImport("Propsys.dll", SetLastError = true)]
        public static extern int PSGetNameFromPropertyKey(ref PropertyKey key, out IntPtr ppszCanonicalName);

        [DllImport("ole32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern int PropVariantClear(ref PropVariant propVariant);

        [DllImport("oleaut32.dll", SetLastError = true)]
        public static extern void VariantInit(IntPtr pVariant);

        [DllImport("oleaut32.dll", SetLastError = true)]
        public static extern void VariantClear(IntPtr pVariant);

        public static T? PropVariantToObject<T>(ref PropVariant propVariant) {
            object? obj = null;
            var pVariant = Marshal.AllocCoTaskMem(24); // sizeof(VARIANT)
            VariantInit(pVariant);
            if (0 <= PropVariantToVariant(ref propVariant, pVariant)) {
                obj = Marshal.GetObjectForNativeVariant(pVariant);
                _ = PropVariantClear(ref propVariant);
            }
            VariantClear(pVariant);
            return (T?)obj;
        }

        public static string? PropertyKeyToName(PropertyKey propKey) {
            try {
                if (0 > PSGetNameFromPropertyKey(ref propKey, out var ptr))
                    return null;

                var name = Marshal.PtrToStringUni(ptr);
                Marshal.FreeCoTaskMem(ptr);
                return name;

            } catch (Exception ex) {
                PluginLog.Error(ex, "PropertyKeyToName({0}) failure", propKey);
                return null;
            }
        }

        public static string MainModuleRva(IntPtr ptr) {
            return $"[ffxiv_dx11.exe+0x{(long)ptr - (long)Process.GetCurrentProcess().MainModule!.BaseAddress:X}]";
        }
    }
}
