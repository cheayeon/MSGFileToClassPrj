using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using MSGFileToClassPrj.Enviroment.Mannager;

namespace MSGFileToClassPrj.Enviroment
{
    public class NativeCOMMethods
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GlobalLock(IntPtr hMem);

        [DllImport("ole32.DLL")]
        public static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, bool fDeleteOnRelease, out ILockBytes ppLkbyt);

        [DllImport("ole32.DLL", CharSet = CharSet.Auto, PreserveSig = false)]
        public static extern IntPtr GetHGlobalFromILockBytes(ILockBytes pLockBytes);

        [DllImport("ole32.DLL")]
        public static extern int StgIsStorageILockBytes(ILockBytes plkbyt);

        [DllImport("ole32.DLL")]
        public static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, STGM grfMode, uint reserved, out IStorage ppstgOpen);

        [DllImport("ole32.DLL")]
        public static extern void StgOpenStorageOnILockBytes(ILockBytes plkbyt, IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstgOpen);

        [DllImport("ole32.DLL")]
        public static extern int StgIsStorageFile([MarshalAs(UnmanagedType.LPWStr)] string wcsName);

        [DllImport("ole32.DLL")]
        public static extern int StgOpenStorage([MarshalAs(UnmanagedType.LPWStr)] string wcsName, IStorage pstgPriority, STGM grfMode, IntPtr snbExclude, int reserved, out IStorage ppstgOpen);

        [ComImport, Guid("0000000A-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface ILockBytes
        {
            void ReadAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbRead);
            void WriteAt([In, MarshalAs(UnmanagedType.U8)] long ulOffset, [In, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 1)] byte[] pv, [In, MarshalAs(UnmanagedType.U4)] int cb, [Out, MarshalAs(UnmanagedType.LPArray)] int[] pcbWritten);
            void Flush();
            void SetSize([In, MarshalAs(UnmanagedType.U8)] long cb);
            void LockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
            void UnlockRegion([In, MarshalAs(UnmanagedType.U8)] long libOffset, [In, MarshalAs(UnmanagedType.U8)] long cb, [In, MarshalAs(UnmanagedType.U4)] int dwLockType);
            void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, [In, MarshalAs(UnmanagedType.U4)] int grfStatFlag);
        }

        [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("0000000B-0000-0000-C000-000000000046")]
        public interface IStorage
        {
            [return: MarshalAs(UnmanagedType.Interface)]
            ComTypes.IStream CreateStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            ComTypes.IStream OpenStream([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr reserved1, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            IStorage CreateStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, [In, MarshalAs(UnmanagedType.U4)] int reserved1, [In, MarshalAs(UnmanagedType.U4)] int reserved2);
            [return: MarshalAs(UnmanagedType.Interface)]
            IStorage OpenStorage([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, IntPtr pstgPriority, [In, MarshalAs(UnmanagedType.U4)] STGM grfMode, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.U4)] int reserved);
            void CopyTo(int ciidExclude, [In, MarshalAs(UnmanagedType.LPArray)] Guid[] pIIDExclude, IntPtr snbExclude, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest);
            void MoveElementTo([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In, MarshalAs(UnmanagedType.Interface)] IStorage stgDest, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName, [In, MarshalAs(UnmanagedType.U4)] int grfFlags);
            void Commit(int grfCommitFlags);
            void Revert();
            void EnumElements([In, MarshalAs(UnmanagedType.U4)] int reserved1, IntPtr reserved2, [In, MarshalAs(UnmanagedType.U4)] int reserved3, [MarshalAs(UnmanagedType.Interface)] out IEnumSTATSTG ppVal);
            void DestroyElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsName);
            void RenameElement([In, MarshalAs(UnmanagedType.BStr)] string pwcsOldName, [In, MarshalAs(UnmanagedType.BStr)] string pwcsNewName);
            void SetElementTimes([In, MarshalAs(UnmanagedType.BStr)] string pwcsName, [In] System.Runtime.InteropServices.ComTypes.FILETIME pctime, [In] System.Runtime.InteropServices.ComTypes.FILETIME patime, [In] System.Runtime.InteropServices.ComTypes.FILETIME pmtime);
            void SetClass([In] ref Guid clsid);
            void SetStateBits(int grfStateBits, int grfMask);
            void Stat([Out]out System.Runtime.InteropServices.ComTypes.STATSTG pStatStg, int grfStatFlag);
        }

        [ComImport, Guid("0000000D-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IEnumSTATSTG
        {
            void Next(uint celt, [MarshalAs(UnmanagedType.LPArray), Out] System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt, out uint pceltFetched);
            void Skip(uint celt);
            void Reset();
            [return: MarshalAs(UnmanagedType.Interface)]
            IEnumSTATSTG Clone();
        }

        // https://learn.microsoft.com/en-us/windows/win32/stg/stgm-constants
        // 위의 경로 참고
        public enum STGM : int
        {
            DIRECT = 0x00000000,
            TRANSACTED = 0x00010000,
            SIMPLE = 0x08000000,
            READ = 0x00000000,
            WRITE = 0x00000001,
            READWRITE = 0x00000002,
            SHARE_DENY_NONE = 0x00000040,
            SHARE_DENY_READ = 0x00000030,
            SHARE_DENY_WRITE = 0x00000020,
            SHARE_EXCLUSIVE = 0x00000010,
            PRIORITY = 0x00040000,
            DELETEONRELEASE = 0x04000000,
            NOSCRATCH = 0x00100000,
            CREATE = 0x00001000,
            CONVERT = 0x00020000,
            FAILIFTHERE = 0x00000000,
            NOSNAPSHOT = 0x00200000,
            DIRECT_SWMR = 0x00400000
        }

        // https://learn.microsoft.com/en-us/office/client-developer/outlook/mapi/property-types
        // 위의 링크 참고
        public enum OutLookMAPI : ushort
        {
            PT_UNSPECIFIED = 0, /* (Reserved for interface use) type doesn't matter to caller */
            PT_NULL = 1,        /* NULL property value */
            PT_I2 = 2,          /* Signed 16-bit value */
            PT_LONG = 3,        /* Signed 32-bit value */
            PT_FLOAT = 4,       /* 4-byte floating point */
            PT_DOUBLE = 5,      /* Floating point double */
            PT_CURRENCY = 6,    /* Signed 64-bit int (decimal w/    4 digits right of decimal pt) */
            PT_APPTIME = 7,     /* Application time */
            PT_ERROR = 10,      /* 32-bit error value */
            PT_BOOLEAN = 11,    /* 16-bit boolean (non-zero true) */
            PT_OBJECT = 13,     /* Embedded object in a property */
            PT_I8 = 20,         /* 8-byte signed integer */
            PT_STRING8 = 30,    /* Null terminated 8-bit character string */
            PT_UNICODE = 31,    /* Null terminated Unicode string */
            PT_SYSTIME = 64,    /* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
            PT_CLSID = 72,      /* OLE GUID */
            PT_BINARY = 258,    /* Uninterpreted (counted byte array) */
        }

        public static IStorage CloneStorage(IStorage source, bool closeSource)
        {
            NativeCOMMethods.IStorage memoryStorage = null;
            NativeCOMMethods.ILockBytes memoryStorageBytes = null;
            try
            {
                //create a ILockBytes (unmanaged byte array) and then create a IStorage using the byte array as a backing store
                NativeCOMMethods.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out memoryStorageBytes);
                NativeCOMMethods.StgCreateDocfileOnILockBytes(memoryStorageBytes, NativeCOMMethods.STGM.CREATE | NativeCOMMethods.STGM.READWRITE | NativeCOMMethods.STGM.SHARE_EXCLUSIVE, 0, out memoryStorage);

                //copy the source storage into the new storage
                source.CopyTo(0, null, IntPtr.Zero, memoryStorage);
                memoryStorageBytes.Flush();
                memoryStorage.Commit(0);

                //ensure memory is released
                ReferenceManager.AddItem(memoryStorage);
            }
            catch
            {
                if (memoryStorage != null)
                {
                    Marshal.ReleaseComObject(memoryStorage);
                }
            }
            finally
            {
                if (memoryStorageBytes != null)
                {
                    Marshal.ReleaseComObject(memoryStorageBytes);
                }

                if (closeSource)
                {
                    Marshal.ReleaseComObject(source);
                }
            }

            return memoryStorage;
        }
    }
}
