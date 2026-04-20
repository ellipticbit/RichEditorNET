using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IOleStream = System.Runtime.InteropServices.ComTypes.IStream;

namespace EllipticBit.RichEditorNET.Support
{
    [StructLayout(LayoutKind.Sequential)]
    internal struct RECTL
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct POINTL
    {
        public int x;
        public int y;
    }

    internal static class OleConstants
    {
        internal const uint DVASPECT_CONTENT = 1;

        internal const int OLECLOSE_SAVEIFDIRTY = 0;
        internal const int OLECLOSE_NOSAVE = 1;

        internal const uint OLEMISC_CANTLINKINSIDE = 0x00000010;
        internal const uint OLEMISC_RECOMPOSEONRESIZE = 0x00000001;
        internal const uint OLEMISC_INSIDEOUT = 0x00000080;
        internal const uint OLEMISC_ACTIVATEWHENVISIBLE = 0x00000100;
        internal const uint OLEMISC_STATIC = 0x00040000;

        internal const int OLE_E_BLANK = unchecked((int)0x80040007);
        internal const int OLE_E_NOTRUNNING = unchecked((int)0x80040005);
        internal const int OLE_E_ADVISENOTSUPPORTED = unchecked((int)0x80040003);
        internal const int OLE_S_USEREG = unchecked((int)0x00040000);
        internal const int E_NOTIMPL = unchecked((int)0x80004001);
        internal const int E_FAIL = unchecked((int)0x80004005);
        internal const int E_INVALIDARG = unchecked((int)0x80070057);
        internal const int E_NOINTERFACE = unchecked((int)0x80004002);
        internal const int S_OK = 0;
        internal const int S_FALSE = 1;
        internal const int DV_E_FORMATETC = unchecked((int)0x80040064);
        internal const int DATA_S_SAMEFORMATETC = unchecked((int)0x00040130);

        internal const short CF_BITMAP = 2;
        internal const short CF_DIB = 8;
        internal const short CF_ENHMETAFILE = 14;

        internal const uint TYMED_HGLOBAL = 1;
        internal const uint TYMED_GDI = 16;
        internal const uint TYMED_ENHMF = 64;
    }

    [ComImport]
    [Guid("00000118-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleClientSite
    {
        [PreserveSig] int SaveObject();
        [PreserveSig] int GetMoniker(uint dwAssign, uint dwWhichMoniker, out IntPtr ppmk);
        [PreserveSig] int GetContainer(out IntPtr ppContainer);
        [PreserveSig] int ShowObject();
        [PreserveSig] int OnShowWindow([MarshalAs(UnmanagedType.Bool)] bool fShow);
        [PreserveSig] int RequestNewObjectLayout();
    }

    [ComImport]
    [Guid("00000112-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleObject
    {
        [PreserveSig] int SetClientSite(IOleClientSite pClientSite);
        [PreserveSig] int GetClientSite(out IOleClientSite ppClientSite);
        [PreserveSig] int SetHostNames([MarshalAs(UnmanagedType.LPWStr)] string szContainerApp, [MarshalAs(UnmanagedType.LPWStr)] string szContainerObj);
        [PreserveSig] int Close(uint dwSaveOption);
        [PreserveSig] int SetMoniker(uint dwWhichMoniker, IntPtr pmk);
        [PreserveSig] int GetMoniker(uint dwAssign, uint dwWhichMoniker, out IntPtr ppmk);
        [PreserveSig] int InitFromData(IDataObject pDataObject, [MarshalAs(UnmanagedType.Bool)] bool fCreation, uint dwReserved);
        [PreserveSig] int GetClipboardData(uint dwReserved, out IDataObject ppDataObject);
        [PreserveSig] int DoVerb(int iVerb, IntPtr lpmsg, IOleClientSite pActiveSite, int lindex, IntPtr hwndParent, ref RECTL lprcPosRect);
        [PreserveSig] int EnumVerbs(out IntPtr ppEnumOleVerb);
        [PreserveSig] int Update();
        [PreserveSig] int IsUpToDate();
        [PreserveSig] int GetUserClassID(out Guid pClsid);
        [PreserveSig] int GetUserType(uint dwFormOfType, [MarshalAs(UnmanagedType.LPWStr)] out string pszUserType);
        [PreserveSig] int SetExtent(uint dwDrawAspect, ref SIZEL psizel);
        [PreserveSig] int GetExtent(uint dwDrawAspect, out SIZEL psizel);
        [PreserveSig] int Advise(IAdviseSink pAdvSink, out uint pdwConnection);
        [PreserveSig] int Unadvise(uint dwConnection);
        [PreserveSig] int EnumAdvise(out IntPtr ppenumAdvise);
        [PreserveSig] int GetMiscStatus(uint dwAspect, out uint pdwStatus);
        [PreserveSig] int SetColorScheme(IntPtr pLogpal);
    }

    [ComImport]
    [Guid("00000127-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IViewObject2
    {
        [PreserveSig] int Draw(uint dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hdcTargetDev, IntPtr hdcDraw, [In] ref RECTL lprcBounds, IntPtr lprcWBounds, IntPtr pfnContinue, IntPtr dwContinue);
        [PreserveSig] int GetColorSet(uint dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hicTargetDev, out IntPtr ppColorSet);
        [PreserveSig] int Freeze(uint dwDrawAspect, int lindex, IntPtr pvAspect, out uint pdwFreeze);
        [PreserveSig] int Unfreeze(uint dwFreeze);
        [PreserveSig] int SetAdvise(uint dwAspects, uint advf, IAdviseSink pAdvSink);
        [PreserveSig] int GetAdvise(IntPtr pAspects, IntPtr pAdvf, out IAdviseSink ppAdvSink);
        [PreserveSig] int GetExtent(uint dwDrawAspect, int lindex, IntPtr ptd, out SIZEL lpsizel);
    }

    [ComImport]
    [Guid("00020D03-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IRichEditOleCallback
    {
        [PreserveSig] int GetNewStorage(out IStorage ppstg);
        [PreserveSig] int GetInPlaceContext(out IntPtr ppipFrame, out IntPtr ppipDoc, IntPtr lpFrameInfo);
        [PreserveSig] int ShowContainerUI([MarshalAs(UnmanagedType.Bool)] bool fShow);
        [PreserveSig] int QueryInsertObject(ref Guid lpclsid, IStorage lpstg, int cp);
        [PreserveSig] int DeleteObject(IntPtr lpoleobj);
        [PreserveSig] int QueryAcceptData(IDataObject lpdataobj, IntPtr lpcfFormat, uint reco, [MarshalAs(UnmanagedType.Bool)] bool fReally, IntPtr hMetaPict);
        [PreserveSig] int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        [PreserveSig] int GetClipboardData(IntPtr lpchrg, uint reco, out IDataObject lplpdataobj);
        [PreserveSig] int GetDragDropEffect([MarshalAs(UnmanagedType.Bool)] bool fDrag, uint grfKeyState, ref uint pdwEffect);
        [PreserveSig] int GetContextMenu(ushort seltype, IntPtr lpoleobj, IntPtr lpchrg, out IntPtr hmenu);
    }

    }
