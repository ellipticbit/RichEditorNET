using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using IOleStream = System.Runtime.InteropServices.ComTypes.IStream;

namespace EllipticBit.RichEditorNET.Support
{
	/// <summary>
	/// Minimal in-process OLE server that wraps an embedded image so it can be inserted into
	/// RICHEDIT50W via IRichEditOle.InsertObject. The image bytes and alt text are persisted
	/// to a single named stream "EBImage" inside the storage that RichEdit hands us, so the
	/// data round-trips through any document operation that flows through IPersistStorage.
	/// Rendering is performed by IViewObject2.Draw using GDI+.
	/// </summary>
	[ComVisible(true)]
	[Guid("EB00010E-7E6A-4D11-8E45-0000000A6E01")]
	[ClassInterface(ClassInterfaceType.None)]
	internal sealed class ImageOleObject : IOleObject, IDataObject, IViewObject2, IPersistStorage
	{
		internal static readonly Guid ClassId = new Guid("EB00010E-7E6A-4D11-8E45-0000000A6E01");
		internal const string StreamName = "EBImage";

		private byte[] _bytes;
		private string _altText = string.Empty;
		private SIZEL _extentHimetric;
		private Image _renderImage;
		private IOleClientSite _clientSite;
		private IStorage _storage;
		private bool _dirty;

		internal byte[] Bytes => _bytes;
		internal string AltText => _altText;
		internal SIZEL Extent => _extentHimetric;

		internal ImageOleObject() { }

		internal ImageOleObject(byte[] bytes, string altText, SIZEL extentHimetric)
		{
			_bytes = bytes ?? Array.Empty<byte>();
			_altText = altText ?? string.Empty;
			_extentHimetric = extentHimetric;
			_dirty = true;
		}

		/// <summary>
		/// Inserts a new embedded image at the current selection / cp via IRichEditOle.InsertObject.
		///
		/// RichEdit renders embedded objects from the OLE presentation cache maintained by the
		/// default handler, not by calling IViewObject::Draw on a live in-proc object. To get
		/// the image to actually paint, we let OLE build a "static picture" object around our
		/// DIB rendering via OleCreateStaticFromData. That gives us a handler-wrapped IOleObject
		/// with a populated CF_DIB cache persisted into the IStorage RichEdit will reload on
		/// subsequent paints.
		///
		/// For clean round-trip of the original image bytes and alt text, we additionally write
		/// our own "EBImage" stream into the same IStorage alongside the OLE cache streams.
		/// </summary>
		internal static int Insert(IRichEditOle rich, int cp, byte[] bytes, string altText, int widthHimetric, int heightHimetric)
		{
			if (rich == null) return OleConstants.E_INVALIDARG;
			if (bytes == null || bytes.Length == 0) return OleConstants.E_INVALIDARG;

			IntPtr pSite = IntPtr.Zero;
			ILockBytes lockBytes = null;
			IStorage storage = null;
			IntPtr pStorage = IntPtr.Zero;
			IntPtr pOleObj = IntPtr.Zero;
			IntPtr pReo = IntPtr.Zero;
			ImageOleObject dataSrc = null;

			try
			{
				int hr = rich.GetClientSite(out pSite);
				if (hr != 0) return hr;

				hr = PInvoke.CreateILockBytesOnHGlobal(IntPtr.Zero, true, out lockBytes);
				if (hr != 0 || lockBytes == null) return OleConstants.E_FAIL;

				const uint mode = StgConstants.STGM_READWRITE | StgConstants.STGM_SHARE_EXCLUSIVE | StgConstants.STGM_CREATE;
				hr = PInvoke.StgCreateDocfileOnILockBytes(lockBytes, mode, 0, out storage);
				if (hr != 0 || storage == null) return OleConstants.E_FAIL;

				var sizeHimetric = new SIZEL { cx = widthHimetric, cy = heightHimetric };

				// Try to wrap the image in an OLE static-picture handler so RichEdit can paint
				// from its presentation cache. If anything goes wrong (GDI+ can't open the bytes,
				// the static helper rejects our IDataObject, etc.) fall back to inserting a raw
				// in-proc IOleObject - that preserves HTML round-trip of bytes/alt even if the
				// visual won't paint until save/close.
				dataSrc = new ImageOleObject(bytes, altText, sizeHimetric);
				Guid clsid = ClassId;
				bool staticSucceeded = false;

				try
				{
					var fmt = new FORMATETC
					{
						cfFormat = OleConstants.CF_DIB,
						ptd = IntPtr.Zero,
						dwAspect = DVASPECT.DVASPECT_CONTENT,
						lindex = -1,
						tymed = TYMED.TYMED_HGLOBAL
					};

					var iidOleObject = new Guid("00000112-0000-0000-C000-000000000046");
					int createHr = PInvoke.OleCreateStaticFromData(
						(System.Runtime.InteropServices.ComTypes.IDataObject)dataSrc,
						ref iidOleObject,
						PInvoke.OLERENDER_FORMAT,
						ref fmt,
						null,
						storage,
						out pOleObj);

					if (createHr >= 0 && pOleObj != IntPtr.Zero)
					{
						staticSucceeded = true;

						// OleCreateStaticFromData calls IPersistStorage::InitNew on our storage
						// but does NOT automatically Save. We must do it explicitly so that the
						// cached CF_DIB bits actually land inside the storage RichEdit will use.
						try
						{
							var persist = (IPersistStorage)Marshal.GetObjectForIUnknown(pOleObj);
							try
							{
								persist.GetClassID(out clsid);
								int sv = persist.Save(storage, true);
								if (sv >= 0) persist.SaveCompleted(null);
							}
							finally
							{
								Marshal.ReleaseComObject(persist);
							}
						}
						catch
						{
							// CLSID_STATICDIB
							clsid = new Guid("0003000A-0000-0000-C000-000000000046");
						}

						// Explicitly set the drawing extent on the wrapper. Some versions of the
						// OLE2 default/static handler initialize cached extent to {0,0}, which
						// causes RichEdit to paint the slot empty even at the correct size.
						try
						{
							var oleObj = (IOleObject)Marshal.GetObjectForIUnknown(pOleObj);
							try
							{
								var ext = sizeHimetric;
								oleObj.SetExtent(OleConstants.DVASPECT_CONTENT, ref ext);
							}
							finally
							{
								Marshal.ReleaseComObject(oleObj);
							}
						}
						catch { /* non-fatal */ }
					}
				}
				catch
				{
					staticSucceeded = false;
				}

				if (!staticSucceeded)
				{
					if (pOleObj != IntPtr.Zero) { Marshal.Release(pOleObj); pOleObj = IntPtr.Zero; }
					// Raw in-proc fallback: persist ourselves into the storage and hand our own
					// IOleObject to RichEdit.
					hr = ((IPersistStorage)dataSrc).Save(storage, true);
					if (hr != 0) return hr;
					((IPersistStorage)dataSrc).SaveCompleted(null);
					pOleObj = Marshal.GetIUnknownForObject(dataSrc);
					clsid = ClassId;
				}

				// Always write our EBImage stream alongside so TryRead can recover original bytes.
				WriteEBImageStream(storage, bytes, altText ?? string.Empty);
				storage.Commit(0);

				// Tell OLE the object is running inside a container. Without this the default
				// handler will not hit its render path when RichEdit asks for a presentation,
				// and the slot paints blank even though the cache is fully populated.
				try
				{
					var oleObjForContained = Marshal.GetObjectForIUnknown(pOleObj);
					try { PInvoke.OleSetContainedObject(oleObjForContained, true); }
					finally { Marshal.ReleaseComObject(oleObjForContained); }
				}
				catch { /* non-fatal */ }

				pStorage = Marshal.GetIUnknownForObject(storage);

				var reo = new REOBJECT
				{
					cbStruct = Marshal.SizeOf(typeof(REOBJECT)),
					cp = cp,
					clsid = clsid,
					poleobj = pOleObj,
					pstg = pStorage,
					polesite = pSite,
					sizel = sizeHimetric,
					dvaspect = OleConstants.DVASPECT_CONTENT,
					dwFlags = 0,
					dwUser = 0
				};

				pReo = Marshal.AllocHGlobal(Marshal.SizeOf(typeof(REOBJECT)));
				Marshal.StructureToPtr(reo, pReo, false);
				try
				{
					hr = rich.InsertObject(pReo);
				}
				finally
				{
					Marshal.DestroyStructure(pReo, typeof(REOBJECT));
				}
				return hr;
			}
			finally
			{
				if (pReo != IntPtr.Zero) Marshal.FreeHGlobal(pReo);
				if (pOleObj != IntPtr.Zero) Marshal.Release(pOleObj);
				if (pStorage != IntPtr.Zero) Marshal.Release(pStorage);
				if (pSite != IntPtr.Zero) Marshal.Release(pSite);
				if (storage != null) Marshal.ReleaseComObject(storage);
				if (lockBytes != null) Marshal.ReleaseComObject(lockBytes);
			}
		}

		private static void WriteEBImageStream(IStorage storage, byte[] bytes, string altText)
		{
			IOleStream stream = null;
			try
			{
				try { storage.DestroyElement(StreamName); } catch { }
				int hr = storage.CreateStream(StreamName,
					StgConstants.STGM_WRITE | StgConstants.STGM_SHARE_EXCLUSIVE | StgConstants.STGM_CREATE,
					0, 0, out stream);
				if (hr != 0 || stream == null) return;

				byte[] altBytes = string.IsNullOrEmpty(altText)
					? Array.Empty<byte>()
					: System.Text.Encoding.Unicode.GetBytes(altText);
				byte[] header = BitConverter.GetBytes(altBytes.Length);
				stream.Write(header, 4, IntPtr.Zero);
				if (altBytes.Length > 0) stream.Write(altBytes, altBytes.Length, IntPtr.Zero);
				if (bytes != null && bytes.Length > 0) stream.Write(bytes, bytes.Length, IntPtr.Zero);
				stream.Commit(0);
			}
			finally
			{
				if (stream != null) Marshal.ReleaseComObject(stream);
			}
		}

		/// <summary>
		/// Reads back an embedded image inserted by <see cref="Insert"/>. Returns true if the
		/// object at <paramref name="cp"/> is one of ours and the data could be parsed.
		/// </summary>
		internal static bool TryRead(IRichEditOle rich, int cp, out byte[] bytes, out string altText, out int widthHimetric, out int heightHimetric)
		{
			bytes = null;
			altText = string.Empty;
			widthHimetric = 0;
			heightHimetric = 0;
			if (rich == null) return false;

			int structSize = Marshal.SizeOf(typeof(REOBJECT));
			var reo = new REOBJECT { cbStruct = structSize, cp = cp };
			int hr = rich.GetObject(ReoConstants.REO_IOB_USE_CP, ref reo, ReoConstants.REO_GETOBJ_PSTG);
			if (hr != 0 || reo.pstg == IntPtr.Zero) { ReleaseReo(ref reo); return false; }

			widthHimetric = reo.sizel.cx;
			heightHimetric = reo.sizel.cy;

			IStorage storage = null;
			IOleStream stream = null;
			try
			{
				storage = (IStorage)Marshal.GetObjectForIUnknown(reo.pstg);
				int shr = storage.OpenStream(StreamName, IntPtr.Zero,
					StgConstants.STGM_READ | StgConstants.STGM_SHARE_EXCLUSIVE, 0, out stream);
				if (shr != 0 || stream == null) return false;

				var probe = new ImageOleObject();
				probe.ReadFromStream(stream);
				bytes = probe._bytes;
				altText = probe._altText ?? string.Empty;
				return bytes != null && bytes.Length > 0;
			}
			catch { return false; }
			finally
			{
				if (stream != null) Marshal.ReleaseComObject(stream);
				if (storage != null) Marshal.ReleaseComObject(storage);
				ReleaseReo(ref reo);
			}
		}

		private static void ReleaseReo(ref REOBJECT reo)
		{
			if (reo.pstg != IntPtr.Zero) { Marshal.Release(reo.pstg); reo.pstg = IntPtr.Zero; }
			if (reo.poleobj != IntPtr.Zero) { Marshal.Release(reo.poleobj); reo.poleobj = IntPtr.Zero; }
			if (reo.polesite != IntPtr.Zero) { Marshal.Release(reo.polesite); reo.polesite = IntPtr.Zero; }
		}

		private void EnsureRenderImage()
		{
			if (_renderImage != null || _bytes == null || _bytes.Length == 0) return;
			try
			{
				using (var ms = new MemoryStream(_bytes, writable: false))
					_renderImage = Image.FromStream(ms, useEmbeddedColorManagement: false, validateImageData: false);
			}
			catch
			{
				_renderImage = null;
			}
		}

		// ---------------- IOleObject ----------------

		public int SetClientSite(IOleClientSite pClientSite) { _clientSite = pClientSite; return OleConstants.S_OK; }
		public int GetClientSite(out IOleClientSite ppClientSite) { ppClientSite = _clientSite; return OleConstants.S_OK; }
		public int SetHostNames(string szContainerApp, string szContainerObj) => OleConstants.S_OK;
		public int Close(uint dwSaveOption) => OleConstants.S_OK;
		public int SetMoniker(uint dwWhichMoniker, IntPtr pmk) => OleConstants.E_NOTIMPL;
		int IOleObject.GetMoniker(uint dwAssign, uint dwWhichMoniker, out IntPtr ppmk) { ppmk = IntPtr.Zero; return OleConstants.E_NOTIMPL; }
		public int InitFromData(IDataObject pDataObject, bool fCreation, uint dwReserved) => OleConstants.E_NOTIMPL;
		public int GetClipboardData(uint dwReserved, out IDataObject ppDataObject) { ppDataObject = this; return OleConstants.S_OK; }
		public int DoVerb(int iVerb, IntPtr lpmsg, IOleClientSite pActiveSite, int lindex, IntPtr hwndParent, ref RECTL lprcPosRect) => OleConstants.S_OK;
		public int EnumVerbs(out IntPtr ppEnumOleVerb) { ppEnumOleVerb = IntPtr.Zero; return OleConstants.OLE_S_USEREG; }
		public int Update() => OleConstants.S_OK;
		public int IsUpToDate() => OleConstants.S_OK;
		public int GetUserClassID(out Guid pClsid) { pClsid = ClassId; return OleConstants.S_OK; }
		public int GetUserType(uint dwFormOfType, out string pszUserType) { pszUserType = "Embedded Image"; return OleConstants.S_OK; }

		public int SetExtent(uint dwDrawAspect, ref SIZEL psizel)
		{
			if (dwDrawAspect != OleConstants.DVASPECT_CONTENT) return OleConstants.E_INVALIDARG;
			_extentHimetric = psizel;
			return OleConstants.S_OK;
		}

		int IOleObject.GetExtent(uint dwDrawAspect, out SIZEL psizel)
		{
			psizel = _extentHimetric;
			return OleConstants.S_OK;
		}

		public int Advise(IAdviseSink pAdvSink, out uint pdwConnection) { pdwConnection = 0; return OleConstants.OLE_E_ADVISENOTSUPPORTED; }
		public int Unadvise(uint dwConnection) => OleConstants.S_OK;
		public int EnumAdvise(out IntPtr ppenumAdvise) { ppenumAdvise = IntPtr.Zero; return OleConstants.OLE_E_ADVISENOTSUPPORTED; }

		public int GetMiscStatus(uint dwAspect, out uint pdwStatus)
		{
			// Note: intentionally NOT setting OLEMISC_STATIC. That flag makes RichEdit render
			// exclusively via IDataObject::GetData(CF_DIB) and skip IViewObject::Draw. We want
			// RichEdit to call our Draw() so we can render the original image bytes via GDI+.
			pdwStatus = OleConstants.OLEMISC_CANTLINKINSIDE | OleConstants.OLEMISC_RECOMPOSEONRESIZE;
			return OleConstants.S_OK;
		}

		public int SetColorScheme(IntPtr pLogpal) => OleConstants.S_OK;

		// ---------------- IViewObject2 ----------------

		public int Draw(uint dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hdcTargetDev, IntPtr hdcDraw, ref RECTL lprcBounds, IntPtr lprcWBounds, IntPtr pfnContinue, IntPtr dwContinue)
		{
			if (dwDrawAspect != OleConstants.DVASPECT_CONTENT) return OleConstants.E_INVALIDARG;
			if (hdcDraw == IntPtr.Zero) return OleConstants.E_INVALIDARG;

			EnsureRenderImage();

			int width = lprcBounds.right - lprcBounds.left;
			int height = lprcBounds.bottom - lprcBounds.top;
			if (width <= 0 || height <= 0) return OleConstants.S_OK;

			try
			{
				using (var g = Graphics.FromHdc(hdcDraw))
				{
					var rect = new Rectangle(lprcBounds.left, lprcBounds.top, width, height);
					if (_renderImage != null)
					{
						g.DrawImage(_renderImage, rect);
					}
					else
					{
						using (var pen = new Pen(Color.Gray, 1f))
						{
							g.FillRectangle(Brushes.LightGray, rect);
							g.DrawRectangle(pen, rect.X, rect.Y, rect.Width - 1, rect.Height - 1);
							g.DrawLine(pen, rect.Left, rect.Top, rect.Right, rect.Bottom);
							g.DrawLine(pen, rect.Right, rect.Top, rect.Left, rect.Bottom);
						}
					}
				}
			}
			catch { return OleConstants.E_FAIL; }
			return OleConstants.S_OK;
		}

		public int GetColorSet(uint dwDrawAspect, int lindex, IntPtr pvAspect, IntPtr ptd, IntPtr hicTargetDev, out IntPtr ppColorSet)
		{
			ppColorSet = IntPtr.Zero;
			return OleConstants.S_FALSE;
		}

		public int Freeze(uint dwDrawAspect, int lindex, IntPtr pvAspect, out uint pdwFreeze) { pdwFreeze = 0; return OleConstants.E_NOTIMPL; }
		public int Unfreeze(uint dwFreeze) => OleConstants.E_NOTIMPL;
		public int SetAdvise(uint dwAspects, uint advf, IAdviseSink pAdvSink) => OleConstants.S_OK;
		public int GetAdvise(IntPtr pAspects, IntPtr pAdvf, out IAdviseSink ppAdvSink) { ppAdvSink = null; return OleConstants.S_OK; }

		public int GetExtent(uint dwDrawAspect, int lindex, IntPtr ptd, out SIZEL lpsizel)
		{
			lpsizel = _extentHimetric;
			return OleConstants.S_OK;
		}

		// ---------------- IDataObject ----------------
		// RichEdit calls these for clipboard / drag-drop. We expose a CF_DIB rendering of our image.

		public int DAdvise(ref FORMATETC pFormatetc, ADVF advf, IAdviseSink adviseSink, out int connection) { connection = 0; return OleConstants.OLE_E_ADVISENOTSUPPORTED; }
		public void DUnadvise(int connection) { }
		public int EnumDAdvise(out IEnumSTATDATA enumAdvise) { enumAdvise = null; return OleConstants.OLE_E_ADVISENOTSUPPORTED; }

		public IEnumFORMATETC EnumFormatEtc(DATADIR direction)
		{
			return new SingleFormatEnumerator();
		}

		public int GetCanonicalFormatEtc(ref FORMATETC formatIn, out FORMATETC formatOut)
		{
			formatOut = formatIn;
			formatOut.ptd = IntPtr.Zero;
			return OleConstants.DATA_S_SAMEFORMATETC;
		}

		public void GetData(ref FORMATETC format, out STGMEDIUM medium)
		{
			//System.Diagnostics.Debug.WriteLine($"ImageOleObject.GetData cf=0x{format.cfFormat:X4} aspect={format.dwAspect} tymed={format.tymed}");
			medium = default;
			if (format.cfFormat == OleConstants.CF_DIB && (format.tymed & TYMED.TYMED_HGLOBAL) != 0)
			{
				EnsureRenderImage();
				if (_renderImage == null) throw new COMException("No image data.", OleConstants.OLE_E_BLANK);
				IntPtr hDib = CreateDibSection(_renderImage);
				if (hDib == IntPtr.Zero) throw new COMException("Failed to create DIB.", OleConstants.E_FAIL);
				//System.Diagnostics.Debug.WriteLine($"ImageOleObject.GetData returning CF_DIB hDib=0x{hDib.ToInt64():X} image={_renderImage.Width}x{_renderImage.Height}");
				medium.tymed = TYMED.TYMED_HGLOBAL;
				medium.unionmember = hDib;
				medium.pUnkForRelease = null;
				return;
			}
			//System.Diagnostics.Debug.WriteLine($"ImageOleObject.GetData REJECTED cf=0x{format.cfFormat:X4}");
			throw new COMException("Format not supported.", OleConstants.DV_E_FORMATETC);
		}

		public void GetDataHere(ref FORMATETC format, ref STGMEDIUM medium) => Marshal.ThrowExceptionForHR(OleConstants.E_NOTIMPL);

		public int QueryGetData(ref FORMATETC format)
		{
			if (format.cfFormat == OleConstants.CF_DIB && (format.tymed & TYMED.TYMED_HGLOBAL) != 0)
				return OleConstants.S_OK;
			return OleConstants.DV_E_FORMATETC;
		}

		public void SetData(ref FORMATETC formatIn, ref STGMEDIUM medium, bool release) => Marshal.ThrowExceptionForHR(OleConstants.E_NOTIMPL);

		private static IntPtr CreateDibSection(Image image)
		{
			try
			{
				using (var bmp = new Bitmap(image))
				{
					// ole32's static picture handler renders from CF_DIB and is strict about
					// format: it expects 24bpp BI_RGB, bottom-up (positive biHeight), with rows
					// aligned on 4-byte boundaries. 32bpp or top-down DIBs render as blank.
					BitmapData bd = bmp.LockBits(new Rectangle(0, 0, bmp.Width, bmp.Height),
						ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
					try
					{
						int width = bd.Width;
						int height = bd.Height;
						int srcStride = bd.Stride;
						int dstStride = ((width * 3) + 3) & ~3; // 4-byte aligned
						int infoSize = 40;
						int dataSize = dstStride * height;
						int total = infoSize + dataSize;

						// CF_DIB HGLOBALs must be GMEM_MOVEABLE. Fixed HGLOBAL
						// (Marshal.AllocHGlobal) is silently rejected by the cache.
						IntPtr hMem = PInvoke.GlobalAlloc(PInvoke.GMEM_MOVEABLE | PInvoke.GMEM_ZEROINIT, (UIntPtr)total);
						if (hMem == IntPtr.Zero) return IntPtr.Zero;

						IntPtr locked = PInvoke.GlobalLock(hMem);
						if (locked == IntPtr.Zero) { PInvoke.GlobalFree(hMem); return IntPtr.Zero; }
						try
						{
							byte[] header = new byte[infoSize];
							Buffer.BlockCopy(BitConverter.GetBytes(infoSize), 0, header, 0, 4);
							Buffer.BlockCopy(BitConverter.GetBytes(width), 0, header, 4, 4);
							Buffer.BlockCopy(BitConverter.GetBytes(height), 0, header, 8, 4); // bottom-up (positive)
							Buffer.BlockCopy(BitConverter.GetBytes((short)1), 0, header, 12, 2);
							Buffer.BlockCopy(BitConverter.GetBytes((short)24), 0, header, 14, 2);
							Buffer.BlockCopy(BitConverter.GetBytes(0), 0, header, 16, 4); // BI_RGB
							Buffer.BlockCopy(BitConverter.GetBytes(dataSize), 0, header, 20, 4);
							Marshal.Copy(header, 0, locked, infoSize);

							// Read source rows (top-down from Bitmap) and write bottom-up into the DIB.
							byte[] srcRow = new byte[srcStride];
							byte[] dstRow = new byte[dstStride];
							IntPtr dstBits = IntPtr.Add(locked, infoSize);
							for (int y = 0; y < height; y++)
							{
								Marshal.Copy(IntPtr.Add(bd.Scan0, y * srcStride), srcRow, 0, srcStride);
								// 24bppRgb in GDI+ is BGR already, matching CF_DIB BI_RGB byte order.
								Buffer.BlockCopy(srcRow, 0, dstRow, 0, width * 3);
								// Zero any padding bytes beyond width*3.
								for (int pad = width * 3; pad < dstStride; pad++) dstRow[pad] = 0;
								int dstY = height - 1 - y; // flip vertically
								Marshal.Copy(dstRow, 0, IntPtr.Add(dstBits, dstY * dstStride), dstStride);
							}
						}
						finally
						{
							PInvoke.GlobalUnlock(hMem);
						}

						return hMem;
					}
					finally
					{
						bmp.UnlockBits(bd);
					}
				}
			}
			catch { return IntPtr.Zero; }
		}

		// ---------------- IPersistStorage ----------------

		public int GetClassID(out Guid pClassID) { pClassID = ClassId; return OleConstants.S_OK; }
		public int IsDirty() => _dirty ? OleConstants.S_OK : OleConstants.S_FALSE;

		public int InitNew(IStorage pstg)
		{
			_storage = pstg;
			_dirty = true;
			return OleConstants.S_OK;
		}

		public int Load(IStorage pstg)
		{
			_storage = pstg;
			IOleStream stream = null;
			try
			{
				int hr = pstg.OpenStream(StreamName, IntPtr.Zero,
					StgConstants.STGM_READ | StgConstants.STGM_SHARE_EXCLUSIVE, 0, out stream);
				if (hr != 0 || stream == null) return OleConstants.S_OK; // empty / new
				ReadFromStream(stream);
				_dirty = false;
				return OleConstants.S_OK;
			}
			catch { return OleConstants.E_FAIL; }
			finally
			{
				if (stream != null) Marshal.ReleaseComObject(stream);
			}
		}

		public int Save(IStorage pstgSave, bool fSameAsLoad)
		{
			IOleStream stream = null;
			try
			{
				try { pstgSave.DestroyElement(StreamName); } catch { }
				int hr = pstgSave.CreateStream(StreamName,
					StgConstants.STGM_WRITE | StgConstants.STGM_SHARE_EXCLUSIVE | StgConstants.STGM_CREATE, 0, 0, out stream);
				if (hr != 0 || stream == null) return OleConstants.E_FAIL;
				WriteToStream(stream);
				try { var clsid = ClassId; pstgSave.SetClass(ref clsid); } catch { }
				_dirty = false;
				return OleConstants.S_OK;
			}
			catch { return OleConstants.E_FAIL; }
			finally
			{
				if (stream != null) Marshal.ReleaseComObject(stream);
			}
		}

		private static Guid Unsafe_ClassId = ClassId;

		public int SaveCompleted(IStorage pstgNew)
		{
			if (pstgNew != null) _storage = pstgNew;
			return OleConstants.S_OK;
		}

		public int HandsOffStorage()
		{
			_storage = null;
			return OleConstants.S_OK;
		}

		private void WriteToStream(IOleStream stream)
		{
			byte[] altBytes = string.IsNullOrEmpty(_altText)
				? Array.Empty<byte>()
				: System.Text.Encoding.Unicode.GetBytes(_altText);
			byte[] header = BitConverter.GetBytes(altBytes.Length);
			stream.Write(header, 4, IntPtr.Zero);
			if (altBytes.Length > 0) stream.Write(altBytes, altBytes.Length, IntPtr.Zero);
			if (_bytes != null && _bytes.Length > 0) stream.Write(_bytes, _bytes.Length, IntPtr.Zero);
		}

		private void ReadFromStream(IOleStream stream)
		{
			// Determine total size via Stat
			stream.Stat(out var stat, 1 /* STATFLAG_NONAME */);
			long total = stat.cbSize;
			if (total < 4) { _bytes = Array.Empty<byte>(); _altText = string.Empty; return; }

			byte[] header = new byte[4];
			IntPtr readPtr = Marshal.AllocHGlobal(sizeof(int));
			try
			{
				stream.Read(header, 4, readPtr);
				int altLen = BitConverter.ToInt32(header, 0);
				if (altLen < 0 || altLen > total - 4) altLen = 0;

				if (altLen > 0)
				{
					byte[] altBuf = new byte[altLen];
					stream.Read(altBuf, altLen, readPtr);
					_altText = System.Text.Encoding.Unicode.GetString(altBuf);
				}
				else
				{
					_altText = string.Empty;
				}

				int remaining = (int)(total - 4 - altLen);
				if (remaining > 0)
				{
					_bytes = new byte[remaining];
					int read = 0;
					while (read < remaining)
					{
						byte[] chunk = new byte[remaining - read];
						stream.Read(chunk, chunk.Length, readPtr);
						int got = Marshal.ReadInt32(readPtr);
						if (got <= 0) break;
						Buffer.BlockCopy(chunk, 0, _bytes, read, got);
						read += got;
					}
					if (read < _bytes.Length)
					{
						byte[] trimmed = new byte[read];
						Buffer.BlockCopy(_bytes, 0, trimmed, 0, read);
						_bytes = trimmed;
					}
				}
				else
				{
					_bytes = Array.Empty<byte>();
				}
			}
			finally
			{
				Marshal.FreeHGlobal(readPtr);
			}

			_renderImage?.Dispose();
			_renderImage = null;
		}

		private sealed class SingleFormatEnumerator : IEnumFORMATETC
		{
			private int _index;
			private static readonly FORMATETC[] _formats = new[]
			{
				new FORMATETC {
					cfFormat = OleConstants.CF_DIB,
					ptd = IntPtr.Zero,
					dwAspect = DVASPECT.DVASPECT_CONTENT,
					lindex = -1,
					tymed = TYMED.TYMED_HGLOBAL
				}
			};

			public int Next(int celt, FORMATETC[] rgelt, int[] pceltFetched)
			{
				int written = 0;
				while (_index < _formats.Length && written < celt && written < rgelt.Length)
				{
					rgelt[written] = _formats[_index];
					_index++;
					written++;
				}
				if (pceltFetched != null && pceltFetched.Length > 0) pceltFetched[0] = written;
				return written == celt ? OleConstants.S_OK : OleConstants.S_FALSE;
			}

			public int Skip(int celt) { _index += celt; return _index <= _formats.Length ? OleConstants.S_OK : OleConstants.S_FALSE; }
			public int Reset() { _index = 0; return OleConstants.S_OK; }
			public void Clone(out IEnumFORMATETC ppenum) { ppenum = new SingleFormatEnumerator { _index = _index }; }
		}
	}
}
