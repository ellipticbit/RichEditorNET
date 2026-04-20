using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.Support
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct SIZEL
    {
        public int cx;
        public int cy;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct REOBJECT
    {
        public int cbStruct;
        public int cp;
        public Guid clsid;
        public IntPtr poleobj;
        public IntPtr pstg;
        public IntPtr polesite;
        public SIZEL sizel;
        public uint dvaspect;
        public uint dwFlags;
        public uint dwUser;
    }

    internal static class ReoConstants
    {
        internal const uint REO_GETOBJ_NO_INTERFACES = 0x00000000;
        internal const uint REO_GETOBJ_POLEOBJ = 0x00000001;
        internal const uint REO_GETOBJ_PSTG = 0x00000002;
        internal const uint REO_GETOBJ_POLESITE = 0x00000004;
        internal const uint REO_GETOBJ_ALL_INTERFACES = 0x00000007;

        internal const int REO_IOB_SELECTION = -1;
        internal const int REO_IOB_USE_CP = -2;
        internal const int REO_CP_SELECTION = -1;
    }

    /// <summary>
    /// IRichEditOle interface exposed by RichEdit controls via EM_GETOLEINTERFACE.
    /// Only GetObject is used here; other vtable slots are stubbed in order.
    /// </summary>
    [ComImport]
    [Guid("00020D00-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IRichEditOle
    {
        [PreserveSig] int GetClientSite(out IntPtr lplpolesite);
        [PreserveSig] int GetObjectCount();
        [PreserveSig] int GetLinkCount();
        [PreserveSig] int GetObject(int iob, ref REOBJECT lpreobject, uint dwFlags);
        [PreserveSig] int InsertObject(IntPtr lpreobject);
        [PreserveSig] int ConvertObject(int iob, Guid rclsidNew, [MarshalAs(UnmanagedType.LPWStr)] string lpstrUserTypeNew);
        [PreserveSig] int ActivateAs(Guid rclsid, Guid rclsidAs);
        [PreserveSig] int SetHostNames([MarshalAs(UnmanagedType.LPStr)] string lpstrContainerApp, [MarshalAs(UnmanagedType.LPWStr)] string lpstrContainerObj);
        [PreserveSig] int SetLinkAvailable(int iob, int fAvailable);
        [PreserveSig] int SetDvaspect(int iob, uint dvaspect);
        [PreserveSig] int HandsOffStorage(int iob);
        [PreserveSig] int SaveCompleted(int iob, IntPtr lpstg);
        [PreserveSig] int InPlaceDeactivate();
        [PreserveSig] int ContextSensitiveHelp(int fEnterMode);
        [PreserveSig] int GetClipboardData(IntPtr lpchrg, uint reco, out IntPtr lplpdataobj);
        [PreserveSig] int ImportDataObject(IntPtr lpdataobj, short cf, IntPtr hMetaPict);
    }
}
