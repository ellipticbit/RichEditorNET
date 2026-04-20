using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace EllipticBit.RichEditorNET.Support
{
    /// <summary>
    /// Minimal IRichEditOleCallback implementation. The only meaningful method we provide is
    /// GetNewStorage, which RichEdit invokes whenever it needs backing storage for a new
    /// embedded object (paste, drag-drop, file load).
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    internal sealed class RichEditOleCallback : IRichEditOleCallback
    {
        public int GetNewStorage(out IStorage ppstg)
        {
            ppstg = null;
            ILockBytes lockBytes = null;
            try
            {
                int hr = PInvoke.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lockBytes);
                if (hr != 0 || lockBytes == null) return OleConstants.E_FAIL;
                const uint mode = StgConstants.STGM_READWRITE | StgConstants.STGM_SHARE_EXCLUSIVE | StgConstants.STGM_CREATE;
                hr = PInvoke.StgCreateDocfileOnILockBytes(lockBytes, mode, 0, out ppstg);
                return hr;
            }
            catch { return OleConstants.E_FAIL; }
            finally
            {
                if (lockBytes != null) Marshal.ReleaseComObject(lockBytes);
            }
        }

        public int GetInPlaceContext(out IntPtr ppipFrame, out IntPtr ppipDoc, IntPtr lpFrameInfo)
        {
            ppipFrame = IntPtr.Zero;
            ppipDoc = IntPtr.Zero;
            return OleConstants.E_NOTIMPL;
        }

        public int ShowContainerUI(bool fShow) => OleConstants.S_OK;

        public int QueryInsertObject(ref Guid lpclsid, IStorage lpstg, int cp) => OleConstants.S_OK;

        public int DeleteObject(IntPtr lpoleobj) => OleConstants.S_OK;

        public int QueryAcceptData(IDataObject lpdataobj, IntPtr lpcfFormat, uint reco, bool fReally, IntPtr hMetaPict) => OleConstants.S_OK;

        public int ContextSensitiveHelp(bool fEnterMode) => OleConstants.S_OK;

        public int GetClipboardData(IntPtr lpchrg, uint reco, out IDataObject lplpdataobj)
        {
            lplpdataobj = null;
            return OleConstants.E_NOTIMPL;
        }

        public int GetDragDropEffect(bool fDrag, uint grfKeyState, ref uint pdwEffect) => OleConstants.S_OK;

        public int GetContextMenu(ushort seltype, IntPtr lpoleobj, IntPtr lpchrg, out IntPtr hmenu)
        {
            hmenu = IntPtr.Zero;
            return OleConstants.S_OK;
        }
    }
}
