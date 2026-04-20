using System;
using System.Runtime.InteropServices;
using IStream = System.Runtime.InteropServices.ComTypes.IStream;

namespace EllipticBit.RichEditorNET.Support
{
    [ComImport]
    [Guid("0000010A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IPersistStorage
    {
        [PreserveSig] int GetClassID(out Guid pClassID);
        [PreserveSig] int IsDirty();
        [PreserveSig] int InitNew(IStorage pstg);
        [PreserveSig] int Load(IStorage pstg);
        [PreserveSig] int Save(IStorage pstgSave, [MarshalAs(UnmanagedType.Bool)] bool fSameAsLoad);
        [PreserveSig] int SaveCompleted(IStorage pstgNew);
        [PreserveSig] int HandsOffStorage();
    }

    [ComImport]
    [Guid("0000000B-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IStorage
    {
        [PreserveSig] int CreateStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStream ppstm);
        [PreserveSig] int OpenStream([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr reserved1, uint grfMode, uint reserved2, out IStream ppstm);
        [PreserveSig] int CreateStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, uint grfMode, uint reserved1, uint reserved2, out IStorage ppstg);
        [PreserveSig] int OpenStorage([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr pstgPriority, uint grfMode, IntPtr snbExclude, uint reserved, out IStorage ppstg);
        [PreserveSig] int CopyTo(uint ciidExclude, IntPtr rgiidExclude, IntPtr snbExclude, IStorage pstgDest);
        [PreserveSig] int MoveElementTo([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IStorage pstgDest, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName, uint grfFlags);
        [PreserveSig] int Commit(uint grfCommitFlags);
        [PreserveSig] int Revert();
        [PreserveSig] int EnumElements(uint reserved1, IntPtr reserved2, uint reserved3, out IEnumSTATSTG ppenum);
        [PreserveSig] int DestroyElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsName);
        [PreserveSig] int RenameElement([MarshalAs(UnmanagedType.LPWStr)] string pwcsOldName, [MarshalAs(UnmanagedType.LPWStr)] string pwcsNewName);
        [PreserveSig] int SetElementTimes([MarshalAs(UnmanagedType.LPWStr)] string pwcsName, IntPtr pctime, IntPtr patime, IntPtr pmtime);
        [PreserveSig] int SetClass(ref Guid clsid);
        [PreserveSig] int SetStateBits(uint grfStateBits, uint grfMask);
        [PreserveSig] int Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, uint grfStatFlag);
    }

    [ComImport]
    [Guid("0000000D-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IEnumSTATSTG
    {
        [PreserveSig] int Next(uint celt,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 0)] System.Runtime.InteropServices.ComTypes.STATSTG[] rgelt,
            out uint pceltFetched);
        [PreserveSig] int Skip(uint celt);
        [PreserveSig] int Reset();
        [PreserveSig] int Clone(out IEnumSTATSTG ppenum);
    }

    [ComImport]
    [Guid("0000000A-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface ILockBytes
    {
        [PreserveSig] int ReadAt(ulong ulOffset, IntPtr pv, uint cb, out uint pcbRead);
        [PreserveSig] int WriteAt(ulong ulOffset, IntPtr pv, uint cb, out uint pcbWritten);
        [PreserveSig] int Flush();
        [PreserveSig] int SetSize(ulong cb);
        [PreserveSig] int LockRegion(ulong libOffset, ulong cb, uint dwLockType);
        [PreserveSig] int UnlockRegion(ulong libOffset, ulong cb, uint dwLockType);
        [PreserveSig] int Stat(out System.Runtime.InteropServices.ComTypes.STATSTG pstatstg, uint grfStatFlag);
    }

    internal static class StgConstants
    {
        internal const uint STGM_READ = 0x00000000;
        internal const uint STGM_WRITE = 0x00000001;
        internal const uint STGM_READWRITE = 0x00000002;
        internal const uint STGM_SHARE_EXCLUSIVE = 0x00000010;
        internal const uint STGM_CREATE = 0x00001000;
        internal const uint STGM_DIRECT = 0x00000000;
        internal const uint STGM_TRANSACTED = 0x00010000;
    }
}
