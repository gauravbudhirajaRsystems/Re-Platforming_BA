//********************************************************
//* © 2022 Litera Corp. All Rights Reserved.
//********************************************************

using System.Xml.Linq;

namespace SharedLibrary
{
	/// <summary>
	/// Helper class for interacting with the word open office xml
	/// </summary>
	public static class W
    {
        // ReSharper disable InconsistentNaming
        public static XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

		public static XName abstractNum = w + "abstractNum";
        public static XName abstractNumId = w + "abstractNumId";
        public static XName adjustRightInd = w + "adjustRightInd";
        public static XName after = w + "after";
        public static XName afterAutospacing = w + "afterAutospacing";
        public static XName aliases = w + "aliases";
        public static XName altName = w + "altName";
        public static XName ansiTheme = w + "ansiTheme";
        public static XName ascii = w + "ascii";
        public static XName asciiTheme = w + "asciiTheme";
        public static XName autoRedefine = w + "autoRedefine";
        public static XName autoSpaceDE = w + "autoSpaceDE";
        public static XName autoSpaceDN = w + "autoSpaceDN";
        public static XName b = w + "b";
        public static XName bCs = w + "bCs";
        public static XName basedOn = w + "basedOn";
        public static XName before = w + "before";
        public static XName beforeAutospacing = w + "beforeAutospacing";
        public static XName bidi = w + "bidi";
        public static XName body = w + "body";
        public static XName bookmarkStart = w + "bookmarkStart";
        public static XName bookmarkEnd = w + "bookmarkEnd";
        public static XName bottom = w + "bottom";
        public static XName br = w + "br";
        public static XName caps = w + "caps";
        public static XName cantSplit = w + "cantSplit";
        public static XName char_ = w + "char";
        public static XName color = w + "color";
        public static XName comment = w + "comment";
        public static XName commentRangeStart = w + "commentRangeStart";
        public static XName commentRangeEnd = w + "commentRangeEnd";
        public static XName commentReference = w + "commentReference";
        public static XName contextualSpacing = w + "contextualSpacing";
        public static XName cs = w + "cs";
        public static XName continuationSeparator = w + "continuationSeparator";
        public static XName customMarkFollows = w + "customMarkFollows";
        public static XName customStyle = w + "customStyle";
        public static XName default_ = w + "default";
        public static XName defaultTabStop = w + "defaultTabStop";
        public static XName del = w + "del";
        public static XName delText = w + "delText";
        public static XName docDefaults = w + "docDefaults";
        public static XName document = w + "document";
        public static XName drawing = w + "drawing";
        public static XName dstrike = w + "dstrike";
        public static XName eastAsia = w + "eastAsia";
        public static XName emboss = w + "emboss";
        public static XName endnote = w + "endnote";
        public static XName endnoteRef = w + "endnoteRef";
        public static XName endnoteReference = w + "endnoteReference";
        public static XName fareast = w + "fareast";
        public static XName firstLine = w + "firstLine";
        public static XName fldChar = w + "fldChar";
        public static XName fldCharType = w + "fldCharType";
        public static XName fldSimple = w + "fldSimple";
        public static XName font = w + "font";
        public static XName fonts = w + "fonts";
        public static XName footer = w + "footer";
        public static XName footnote = w + "footnote";
        public static XName footnoteRef = w + "footnoteRef";
        public static XName footnoteReference = w + "footnoteReference";
        public static XName gridCol = w + "gridCol";
        public static XName hAnsi = w + "hAnsi";
        public static XName ftr = w + "ftr";
        public static XName i = w + "i";
        public static XName iCs = w + "iCs";
        public static XName id = w + "id";
        public static XName ilvl = w + "ilvl";
        public static XName imprint = w + "imprint";
        public static XName ind = w + "ind";
        public static XName ins = w + "ins";
        public static XName instr = w + "instr";
        public static XName instrText = w + "instrText";
        public static XName isLgl = w + "isLgl";
        public static XName jc = w + "jc";
        public static XName hanging = w + "hanging";
        public static XName hdr = w + "hdr";
        public static XName header = w + "header";
        public static XName highlight = w + "highlight";
        public static XName hrule = w + "hRule";
        public static XName keepLines = w + "keepLines";
        public static XName keepNext = w + "keepNext";
        public static XName kinsoku = w + "kinsoku";
        public static XName lang = w + "lang";
        public static XName left = w + "left";
        public static XName leader = w + "leader";
        public static XName line = w + "line";
        public static XName lineRule = w + "lineRule";
        public static XName link = w + "link";
        public static XName lock_ = w + "lock";
        public static XName lvl = w + "lvl";
        public static XName lvlOverride = w + "lvlOverride";
        public static XName lvlText = w + "lvlText";
        public static XName mirrorIndents = w + "mirrorIndents";
        public static XName moveFrom = w + "moveFrom";
        public static XName moveFromRangeEnd = w + "moveFromRangeEnd";
        public static XName moveFromRangeStart = w + "moveFromRangeStart";
        public static XName moveTo = w + "moveTo";
        public static XName moveToRangeEnd = w + "moveToRangeEnd";
        public static XName moveToRangeStart = w + "moveToRangeStart";
        public static XName name = w + "name";
        public static XName next = w + "next";
        public static XName noProof = w + "noProof";
        public static XName num = w + "num";
        public static XName numbering = w + "numbering";
        public static XName numFmt = w + "numFmt";
        public static XName numId = w + "numId";
        public static XName numPr = w + "numPr";
        public static XName oMath = w + "oMath";
        public static XName oMathPara = w + "oMathPara";
        public static XName outline = w + "outline";
        public static XName outlineLvl = w + "outlineLvl";
        public static XName overflowPunct = w + "overflowPunct";
        public static XName p = w + "p";
        public static XName pBdr = w + "pBdr";
        public static XName pageBreakBefore = w + "pageBreakBefore";
        public static XName pgBorders = w + "pgBorders";
        public static XName pgMar = w + "pgMar";
        public static XName pict = w + "pict";
        public static XName pos = w + "pos";
        public static XName position = w + "position";
        public static XName pPr = w + "pPr";
        public static XName pPrDefault = w + "pPrDefault";
        public static XName pStyle = w + "pStyle";
        public static XName r = w + "r";
        public static XName rFonts = w + "rFonts";
        public static XName right = w + "right";
        public static XName rPr = w + "rPr";
        public static XName rPrDefault = w + "rPrDefault";
        public static XName rStyle = w + "rStyle";
        public static XName rtl = w + "rtl";
        public static XName sdt = w + "sdt";
        public static XName sdtPr = w + "sdtPr";
        public static XName sdtContent = w + "sdtContent";
        public static XName sectPr = w + "sectPr";
        public static XName separator = w + "separator";
        public static XName settings = w + "settings";
        public static XName shadow = w + "shadow";
        public static XName shd = w + "shd";
        public static XName spacing = w + "spacing";
        public static XName smallCaps = w + "smallCaps";
        public static XName smartTag = w + "smartTag";
        public static XName smartTagPr = w + "smartTagPr";
        public static XName snapToGrid = w + "snapToGrid";
        public static XName startOverride = w + "startOverride";
        public static XName strike = w + "strike";
        public static XName style = w + "style";
        public static XName styles = w + "styles";
        public static XName styleId = w + "styleId";
        public static XName suff = w + "suff";
        public static XName suppressOverlap = w + "suppressOverlap";
        public static XName suppressRef = w + "suppressRef";
        public static XName suppressAutoHyphens = w + "suppressAutoHyphens";
        public static XName suppressLineNumbers = w + "suppressLineNumbers";
        public static XName sym = w + "sym";
        public static XName sz = w + "sz";
        public static XName t = w + "t";
        public static XName tab = w + "tab";
        public static XName tabs = w + "tabs";
        public static XName tbl = w + "tbl";
        public static XName tblBorders = w + "tblBorders";
        public static XName tblGrid = w + "tblGrid";
        public static XName tblHeader = w + "tblHeader";
        public static XName tblInd = w + "tblInd";
        public static XName tblLayout = w + "tblLayout";
        public static XName tblOverlap = w + "tblOverlap";
        public static XName tblPr = w + "tblPr";
        public static XName tblStyle = w + "tblStyle";
        public static XName tblW = w + "tblW";
        public static XName tc = w + "tc";
        public static XName tcPr = w + "tcPr";
        public static XName top = w + "top";
        public static XName topLinePunct = w + "topLinePunct";
        public static XName trackRevisions = w + "trackRevisions";
        public static XName trHeight = w + "trHeight";
        public static XName tr = w + "tr";
        public static XName trPr = w + "trPr";
        public static XName txbxContent = w + "txbxContent";
        public static XName type = w + "type";
        public static XName u = w + "u";
        public static XName val = w + "val";
        public static XName vanish = w + "vanish";
        public static XName vAlign = w + "vAlign";
        public static XName vertAlign = w + "vertAlign";
        public static XName ww = w + "w";
        public static XName webHidden = w + "webHidden";
        public static XName widowControl = w + "widowControl";
        public static XName wordWrap = w + "wordWrap";
        public static XName hyperLink = w + "hyperlink";
    }

	public static class AP
	{
		public static XNamespace ap = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
		public static XName Properties = ap + "Properties";
	}

	public static class MC
    {
        public static XNamespace mc = "http://schemas.openxmlformats.org/markup-compatibility/2006";

        public static XName AlternateContent = mc + "AlternateContent";
        public static XName Choice = mc + "Choice";
    }

	public static class WP
    {
        public static XNamespace wp = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
        public static XName anchor = wp + "anchor";
    }

	public static class A
    {
        public static XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";

        public static XName graphic = a + "graphic";
        public static XName graphicData = a + "graphicData";
        public static XName xfrm = a + "xfrm";
        public static XName off = a + "off";
        public static XName ext = a + "ext";
        public static XName prstGeom = a + "prstGeom";

        // Related to themes.
        public static XName theme = a + "theme";
        public static XName fontScheme = a + "fontScheme";
        public static XName majorFont = a + "majorFont";
        public static XName minorFont = a + "minorFont";
        public static XName latin = a + "latin";
        public static XName font = a + "font";
    }

    public static class Pkg
    {
        public static XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";

        public static XName package = pkg + "package";
        public static XName part = pkg + "part";
        public static XName name = pkg + "name";
        public static XName contentType = pkg + "contentType";
        public static XName compression = pkg + "compression";
        public static XName binaryData = pkg + "binaryData";
        public static XName xmlData = pkg + "xmlData";
    }

    public static class Core
    {
        public static XNamespace cp = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";

        public static XName coreProperties = cp + "coreProperties";
    }

    public static class DS
    {
        public static XNamespace ds = "http://schemas.openxmlformats.org/officeDocument/2006/customXml";

        public static XName datastoreItem = ds + "datastoreItem";
        public static XName itemID = ds + "itemID";
    }

    public static class WPS
    {
        public static XNamespace wps = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";

        public static XName txbx = wps + "txtbx";
        public static XName wsp = wps + "wsp";
        public static XName spPr = wps + "spPr";

    }

    public static class REL
    {
        public static XNamespace rel = "http://schemas.openxmlformats.org/package/2006/relationships";

        public static XName Relationships = rel + "Relationships";
        public static XName Relationship = rel + "Relationship";
    }

    public static class V
    {
        public static XNamespace v = "urn:schemas-microsoft-com:vml";

        public static XName shapetype = v + "shapetype";
        public static XName shape = v + "shape";
        public static XName rect = v + "rect";
        public static XName textbox = v + "textbox";
        public static XName line = v + "line";
        public static XName stroke = v + "stroke";
    }

    public static class DCTerms
    {
        public static XNamespace dc = "http://purl.org/dc/terms/";

        public static XName created = dc + "created";
        public static XName modified = dc + "modified";
    }
}
