using System;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Contains constants used throughout the Text Object Model (TOM) and TOM2 interfaces.
	/// These values correspond to the tomConstants enumeration defined in the Windows SDK header tom.h.
	/// </summary>
	public static class tomConstants
	{
		// Boolean and state values
		public const int tomFalse = 0;
		public const int tomTrue = -1;
		public const int tomUndefined = -9999999;
		public const int tomToggle = -9999998;
		public const int tomAutoColor = -9999997;
		public const int tomDefault = -9999996;
		public const int tomSuspend = -9999995;
		public const int tomResume = -9999994;

		// Apply modes
		public const int tomApplyNow = 0;
		public const int tomApplyLater = 1;
		public const int tomTrackParms = 2;
		public const int tomCacheParms = 3;
		public const int tomApplyTmp = 4;
		public const int tomDisableSmartFont = 8;
		public const int tomEnableSmartFont = 9;
		public const int tomUsePoints = 10;
		public const int tomUseTwips = 11;

		// Direction
		public const int tomBackward = unchecked((int)0xC0000001);
		public const int tomForward = 0x3FFFFFFF;

		// Move/Extend
		public const int tomMove = 0;
		public const int tomExtend = 1;

		// Collapse
		public const int tomCollapseEnd = 0;
		public const int tomCollapseStart = 1;

		// Start/End and point type flags
		public const int tomEnd = 0;
		public const int tomStart = 32;
		public const int tomClientCoord = 256;
		public const int tomAllowOffClient = 512;
		public const int tomTransform = 1024;
		public const int tomObjectArg = 2048;
		public const int tomAtEnd = 4096;

		// Selection types
		public const int tomNoSelection = 0;
		public const int tomSelectionIP = 1;
		public const int tomSelectionNormal = 2;
		public const int tomSelectionFrame = 3;
		public const int tomSelectionColumn = 4;
		public const int tomSelectionRow = 5;
		public const int tomSelectionBlock = 6;
		public const int tomSelectionInlineShape = 7;
		public const int tomSelectionShape = 8;

		// Selection flags
		public const int tomSelStartActive = 1;
		public const int tomSelAtEOL = 2;
		public const int tomSelOvertype = 4;
		public const int tomSelActive = 8;
		public const int tomSelReplace = 16;

		// Unit types
		public const int tomCharacter = 1;
		public const int tomWord = 2;
		public const int tomSentence = 3;
		public const int tomParagraph = 4;
		public const int tomLine = 5;
		public const int tomStory = 6;
		public const int tomScreen = 7;
		public const int tomSection = 8;
		public const int tomTableColumn = 9;
		public const int tomColumn = 9;
		public const int tomRow = 10;
		public const int tomWindow = 11;
		public const int tomCell = 12;
		public const int tomCharFormat = 13;
		public const int tomParaFormat = 14;
		public const int tomTable = 15;
		public const int tomObject = 16;
		public const int tomPage = 17;
		public const int tomHardParagraph = 128;
		public const int tomCluster = 129;
		public const int tomInlineObject = 130;
		public const int tomInlineObjectArg = 131;
		public const int tomLeafLine = 132;
		public const int tomLayoutColumn = 133;
		public const int tomProcessId = 0x40000001;

		// Find flags
		public const int tomMatchWord = 2;
		public const int tomMatchCase = 4;
		public const int tomMatchPattern = 8;

		// Change case types
		public const int tomLowerCase = 0;
		public const int tomUpperCase = 1;
		public const int tomTitleCase = 2;
		public const int tomSentenceCase = 4;
		public const int tomToggleCase = 5;

		// Underline styles
		public const int tomNone = 0;
		public const int tomSingle = 1;
		public const int tomWords = 2;
		public const int tomDouble = 3;
		public const int tomDotted = 4;
		public const int tomDash = 5;
		public const int tomDashDot = 6;
		public const int tomDashDotDot = 7;
		public const int tomWave = 8;
		public const int tomThick = 9;
		public const int tomHair = 10;
		public const int tomDoubleWave = 11;
		public const int tomHeavyWave = 12;
		public const int tomLongDash = 13;
		public const int tomThickDash = 14;
		public const int tomThickDashDot = 15;
		public const int tomThickDashDotDot = 16;
		public const int tomThickDotted = 17;
		public const int tomThickLongDash = 18;

		// Animation types
		public const int tomNoAnimation = 0;
		public const int tomLasVegasLights = 1;
		public const int tomBlinkingBackground = 2;
		public const int tomSparkleText = 3;
		public const int tomMarchingBlackAnts = 4;
		public const int tomMarchingRedAnts = 5;
		public const int tomShimmer = 6;
		public const int tomWipeDown = 7;
		public const int tomWipeRight = 8;
		public const int tomAnimationMax = 8;

		// Paragraph alignment
		public const int tomAlignLeft = 0;
		public const int tomAlignCenter = 1;
		public const int tomAlignRight = 2;
		public const int tomAlignJustify = 3;
		public const int tomAlignDecimal = 3;
		public const int tomAlignBar = 4;
		public const int tomDefaultTab = 5;
		public const int tomAlignInterWord = 3;
		public const int tomAlignNewspaperJustify = 4;
		public const int tomAlignInterLetter = 5;
		public const int tomAlignScaled = 6;

		// Tab alignment
		public const int tomAlignTabLeft = 0;
		public const int tomAlignTabCenter = 1;
		public const int tomAlignTabRight = 2;
		public const int tomAlignTabDecimal = 3;
		public const int tomAlignTabWord = 4;

		// Tab leader styles
		public const int tomSpaces = 0;
		public const int tomDots = 1;
		public const int tomDashes = 2;
		public const int tomLines = 3;
		public const int tomThickLines = 4;
		public const int tomEquals = 5;

		// Tab position constants
		public const int tomTabBack = -3;
		public const int tomTabNext = -1;
		public const int tomTabHere = -2;

		// Line spacing rules
		public const int tomLineSpaceSingle = 0;
		public const int tomLineSpace1pt5 = 1;
		public const int tomLineSpaceDouble = 2;
		public const int tomLineSpaceAtLeast = 3;
		public const int tomLineSpaceExactly = 4;
		public const int tomLineSpaceMultiple = 5;
		public const int tomLineSpacePercent = 6;

		// List types
		public const int tomListNone = 0;
		public const int tomListBullet = 1;
		public const int tomListNumberAsArabic = 2;
		public const int tomListNumberAsLCLetter = 3;
		public const int tomListNumberAsUCLetter = 4;
		public const int tomListNumberAsLCRoman = 5;
		public const int tomListNumberAsUCRoman = 6;
		public const int tomListNumberAsSequence = 7;
		public const int tomListNumberedCircle = 8;
		public const int tomListNumberedBlackCircleWingding = 9;
		public const int tomListNumberedWhiteCircleWingding = 10;
		public const int tomListNumberedArabicWide = 11;
		public const int tomListNumberedChS = 12;
		public const int tomListNumberedChT = 13;
		public const int tomListNumberedJpnChS = 14;
		public const int tomListNumberedJpnKor = 15;
		public const int tomListNumberedArabic1 = 16;
		public const int tomListNumberedArabic2 = 17;
		public const int tomListNumberedHebrew = 18;
		public const int tomListNumberedThaiAlpha = 19;
		public const int tomListNumberedThaiNum = 20;
		public const int tomListNumberedHindiAlpha = 21;
		public const int tomListNumberedHindiAlpha1 = 22;
		public const int tomListNumberedHindiNum = 23;

		// List format flags
		public const int tomListParentheses = 0x10000;
		public const int tomListPeriod = 0x20000;
		public const int tomListPlain = 0x30000;
		public const int tomListMinus = 0x80000;

		// Story types
		public const int tomUnknownStory = 0;
		public const int tomMainTextStory = 1;
		public const int tomFootnotesStory = 2;
		public const int tomEndnotesStory = 3;
		public const int tomCommentsStory = 4;
		public const int tomTextFrameStory = 5;
		public const int tomEvenPagesHeaderStory = 6;
		public const int tomPrimaryHeaderStory = 7;
		public const int tomEvenPagesFooterStory = 8;
		public const int tomPrimaryFooterStory = 9;
		public const int tomFirstPageHeaderStory = 10;
		public const int tomFirstPageFooterStory = 11;
		public const int tomScratchStory = 127;
		public const int tomFindStory = 128;
		public const int tomReplaceStory = 129;

		// Story active states
		public const int tomStoryInactive = 0;
		public const int tomStoryActiveDisplay = 1;
		public const int tomStoryActiveUI = 2;
		public const int tomStoryActiveDisplayUI = 3;

		// Open/Save format flags
		public const int tomRTF = 0x1;
		public const int tomText = 0x2;
		public const int tomHTML = 0x3;
		public const int tomWordDocument = 0x4;

		// Open/Save mode flags
		public const int tomCreateNew = 0x10;
		public const int tomCreateAlways = 0x20;
		public const int tomOpenExisting = 0x30;
		public const int tomOpenAlways = 0x40;
		public const int tomTruncateExisting = 0x50;
		public const int tomReadOnly = 0x100;
		public const int tomShareDenyRead = 0x200;
		public const int tomShareDenyWrite = 0x400;
		public const int tomPasteFile = 0x1000;

		// GetText2/SetText2 flags
		public const int tomUseCRLF = 0x1;
		public const int tomTextize = 0x2;
		public const int tomAllowFinalEOP = 0x4;
		public const int tomFoldMathAlpha = 0x8;
		public const int tomNoHidden = 0x20;
		public const int tomIncludeNumbering = 0x40;
		public const int tomTranslateTableCell = 0x80;
		public const int tomNoMathZoneBrackets = 0x100;
		public const int tomConvertMathChar = 0x200;
		public const int tomNoUCGreekItalic = 0x400;
		public const int tomAllowMathBold = 0x800;
		public const int tomLanguageTag = 0x1000;
		public const int tomConvertRTF = 0x2000;
		public const int tomApplyRtfDocProps = 0x4000;

		// Character effect flags (ITextFont/ITextFont2 effects masks)
		public const int tomBold = unchecked((int)0x80000001);
		public const int tomItalic = unchecked((int)0x80000002);
		public const int tomUnderline = unchecked((int)0x80000004);
		public const int tomStrikeout = unchecked((int)0x80000008);
		public const int tomProtected = unchecked((int)0x80000010);
		public const int tomLink = unchecked((int)0x80000020);
		public const int tomSmallCaps = unchecked((int)0x80000040);
		public const int tomAllCaps = unchecked((int)0x80000080);
		public const int tomHidden = unchecked((int)0x80000100);
		public const int tomOutline = unchecked((int)0x80000200);
		public const int tomShadow = unchecked((int)0x80000400);
		public const int tomEmboss = unchecked((int)0x80000800);
		public const int tomImprint = unchecked((int)0x80001000);
		public const int tomDisabled = unchecked((int)0x80002000);
		public const int tomRevised = unchecked((int)0x80004000);
		public const int tomSubscriptCF = unchecked((int)0x80010000);
		public const int tomSuperscriptCF = unchecked((int)0x80020000);
		public const int tomFontBound = unchecked((int)0x80100000);
		public const int tomLinkProtected = unchecked((int)0x80800000);
		public const int tomInlineObjectStart = unchecked((int)0x81000000);
		public const int tomExtendedChar = unchecked((int)0x82000000);
		public const int tomAutoBackColor = unchecked((int)0x84000000);
		public const int tomMathZoneNoBuildUp = unchecked((int)0x88000000);
		public const int tomMathZone = unchecked((int)0x90000000);
		public const int tomMathZoneOrdinary = unchecked((int)0xA0000000);
		public const int tomAutoTextColor = unchecked((int)0xC0000000);
		public const int tomMathZoneDisplay = 0x40000;

		// Paragraph effect flags (ITextPara2 effects masks)
		public const int tomParaEffectRTL = 0x1;
		public const int tomParaEffectKeep = 0x2;
		public const int tomParaEffectKeepNext = 0x4;
		public const int tomParaEffectPageBreakBefore = 0x8;
		public const int tomParaEffectNoLineNumber = 0x10;
		public const int tomParaEffectNoWidowControl = 0x20;
		public const int tomParaEffectDoNotHyphen = 0x40;
		public const int tomParaEffectSideBySide = 0x80;
		public const int tomParaEffectCollapsed = 0x100;
		public const int tomParaEffectOutlineLevel = 0x200;
		public const int tomParaEffectBox = 0x400;
		public const int tomParaEffectTableRowDelimiter = 0x1000;
		public const int tomParaEffectTable = 0x4000;

		// Font effect flags (ITextFont2 effects 2 masks)
		public const int tomModWidthPairs = 0x1;
		public const int tomModWidthSpace = 0x2;
		public const int tomAutoSpaceAlpha = 0x4;
		public const int tomAutoSpaceNumeric = 0x8;
		public const int tomAutoSpaceParens = 0x10;
		public const int tomEmbeddedFont = 0x20;
		public const int tomDoublestrike = 0x40;
		public const int tomOverlapping = 0x80;

		// Caret types
		public const int tomNormalCaret = 0;
		public const int tomKoreanBlockCaret = 0x1;
		public const int tomNullCaret = 0x2;
		public const int tomInclinesCaret = 0x3;
		public const int tomKoreanCaret = 0x4;

		// Gravity
		public const int tomGravityUI = 0;
		public const int tomGravityBack = 1;
		public const int tomGravityFore = 2;
		public const int tomGravityIn = 3;
		public const int tomGravityOut = 4;

		// TeX style
		public const int tomStyleDefault = 0;
		public const int tomStyleScriptScriptCramped = 1;
		public const int tomStyleScriptScript = 2;
		public const int tomStyleScriptCramped = 3;
		public const int tomStyleScript = 4;
		public const int tomStyleTextCramped = 5;
		public const int tomStyleText = 6;
		public const int tomStyleDisplayCramped = 7;
		public const int tomStyleDisplay = 8;

		// Character representation (CharRep)
		public const int tomAnsi = 0;
		public const int tomEastEurope = 1;
		public const int tomCyrillic = 2;
		public const int tomGreek = 3;
		public const int tomTurkish = 4;
		public const int tomHebrew = 5;
		public const int tomArabic = 6;
		public const int tomBaltic = 7;
		public const int tomVietnamese = 8;
		public const int tomDefaultCharRep = 9;
		public const int tomSymbol = 10;
		public const int tomThai = 11;
		public const int tomShiftJIS = 12;
		public const int tomGB2312 = 13;
		public const int tomHangul = 14;
		public const int tomBIG5 = 15;
		public const int tomPC437 = 16;
		public const int tomOEM = 17;
		public const int tomMac = 18;
		public const int tomArmenian = 19;
		public const int tomSyriac = 20;
		public const int tomThaana = 21;
		public const int tomDevanagari = 22;
		public const int tomBengali = 23;
		public const int tomGurmukhi = 24;
		public const int tomGujarati = 25;
		public const int tomOriya = 26;
		public const int tomTamil = 27;
		public const int tomTelugu = 28;
		public const int tomKannada = 29;
		public const int tomMalayalam = 30;
		public const int tomSinhala = 31;
		public const int tomLao = 32;
		public const int tomTibetan = 33;
		public const int tomMyanmar = 34;
		public const int tomGeorgian = 35;
		public const int tomJamo = 36;
		public const int tomEthiopic = 37;
		public const int tomCherokee = 38;
		public const int tomAboriginal = 39;
		public const int tomOgham = 40;
		public const int tomRunic = 41;
		public const int tomKhmer = 42;
		public const int tomMongolian = 43;
		public const int tomBraille = 44;
		public const int tomYi = 45;
		public const int tomLimbu = 46;
		public const int tomTaiLe = 47;
		public const int tomNewTaiLue = 48;
		public const int tomSylotiNagri = 49;
		public const int tomKharoshthi = 50;
		public const int tomKayahli = 51;
		public const int tomUsymbol = 52;
		public const int tomEmoji = 53;
		public const int tomGlagolitic = 54;
		public const int tomLisu = 55;
		public const int tomVai = 56;
		public const int tomNKo = 57;
		public const int tomOsmanya = 58;
		public const int tomPhagsPa = 59;
		public const int tomGothic = 60;
		public const int tomDeseret = 61;
		public const int tomTifinagh = 62;
		public const int tomCharRepMax = 63;

		// Text flow
		public const int tomTextFlowES = 0;
		public const int tomTextFlowSW = 1;
		public const int tomTextFlowWN = 2;
		public const int tomTextFlowNE = 3;
		public const int tomTextFlowMask = 0xF;

		// Phantom flags
		public const int tomPhantomShow = 1;
		public const int tomPhantomZeroWidth = 2;
		public const int tomPhantomZeroAscent = 4;
		public const int tomPhantomZeroDescent = 8;
		public const int tomPhantomTransparent = 16;
		public const int tomPhantomASmash = tomPhantomShow | tomPhantomZeroAscent;
		public const int tomPhantomDSmash = tomPhantomShow | tomPhantomZeroDescent;
		public const int tomPhantomHSmash = tomPhantomShow | tomPhantomZeroWidth;
		public const int tomPhantomSmash = tomPhantomShow | tomPhantomZeroAscent | tomPhantomZeroDescent;
		public const int tomPhantomHorz = tomPhantomZeroAscent | tomPhantomZeroDescent;
		public const int tomPhantomVert = tomPhantomZeroWidth;

		// Math display properties (GetMathProperties/SetMathProperties)
		public const int tomMathDispAlignMask = 3;
		public const int tomMathDispAlignCenterGroup = 0;
		public const int tomMathDispAlignCenter = 1;
		public const int tomMathDispAlignLeft = 2;
		public const int tomMathDispAlignRight = 3;
		public const int tomMathDispIntUnderOver = 4;
		public const int tomMathDispFracTeX = 8;
		public const int tomMathDispNaryGrow = 0x10;
		public const int tomMathDocEmptyArgMask = 0x60;
		public const int tomMathDocEmptyArgAuto = 0;
		public const int tomMathDocEmptyArgAlways = 0x20;
		public const int tomMathDocEmptyArgNever = 0x40;
		public const int tomMathDocSbSpOpUnchanged = 0x80;
		public const int tomMathDocDiffMask = 0x300;
		public const int tomMathDocDiffDefault = 0;
		public const int tomMathDocDiffUpright = 0x100;
		public const int tomMathDocDiffItalic = 0x200;
		public const int tomMathDocDiffOpenItalic = 0x300;
		public const int tomMathEnableRtl = 0x01000000;
		public const int tomMathBrkBinBefore = 0;
		public const int tomMathBrkBinAfter = 1;
		public const int tomMathBrkBinDup = 2;
		public const int tomMathBrkBinMask = 3;
		public const int tomMathBrkBinSubMask = 0xC;
		public const int tomMathBrkBinSubMM = 0;
		public const int tomMathBrkBinSubPM = 4;
		public const int tomMathBrkBinSubMP = 8;

		// Typography options (SetTypographyOptions)
		public const int tomAdvancedTypographyTraditional = 0;
		public const int tomAdvancedTypographyDefault = 1;

		// Font property types (ITextFont2.GetProperty/SetProperty)
		public const int tomFontPropTeXStyle = 0x33E;
		public const int tomFontPropAlign = 0x33F;
		public const int tomFontStretch = 0x300;
		public const int tomFontStyle = 0x301;
		public const int tomFontStyleUpright = 0;
		public const int tomFontStyleOblique = 1;
		public const int tomFontStyleItalic = 2;
		public const int tomFontStretchDefault = 0;
		public const int tomFontStretchUltraCondensed = 1;
		public const int tomFontStretchExtraCondensed = 2;
		public const int tomFontStretchCondensed = 3;
		public const int tomFontStretchSemiCondensed = 4;
		public const int tomFontStretchNormal = 5;
		public const int tomFontStretchSemiExpanded = 6;
		public const int tomFontStretchExpanded = 7;
		public const int tomFontStretchExtraExpanded = 8;
		public const int tomFontStretchUltraExpanded = 9;
		public const int tomFontWeightDefault = 0;
		public const int tomFontWeightThin = 100;
		public const int tomFontWeightExtraLight = 200;
		public const int tomFontWeightLight = 300;
		public const int tomFontWeightNormal = 400;
		public const int tomFontWeightMedium = 500;
		public const int tomFontWeightSemiBold = 600;
		public const int tomFontWeightBold = 700;
		public const int tomFontWeightExtraBold = 800;
		public const int tomFontWeightBlack = 900;
		public const int tomFontWeightHeavy = 900;
		public const int tomFontWeightExtraBlack = 950;
	}
}
