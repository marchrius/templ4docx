package pl.jsolve.templ4docx.utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import pl.jsolve.sweetener.collection.Maps;

public class DocxHandler {

  private static Map<String, Integer> EXTENSIONS_MAPPING = Maps.newHashMap();
  static {
    EXTENSIONS_MAPPING.put(".emf", XWPFDocument.PICTURE_TYPE_EMF);
    EXTENSIONS_MAPPING.put(".wmf", XWPFDocument.PICTURE_TYPE_WMF);
    EXTENSIONS_MAPPING.put(".pict", XWPFDocument.PICTURE_TYPE_PICT);
    EXTENSIONS_MAPPING.put(".jpeg", XWPFDocument.PICTURE_TYPE_JPEG);
    EXTENSIONS_MAPPING.put(".jpg", XWPFDocument.PICTURE_TYPE_JPEG);
    EXTENSIONS_MAPPING.put(".png", XWPFDocument.PICTURE_TYPE_PNG);
    EXTENSIONS_MAPPING.put(".dib", XWPFDocument.PICTURE_TYPE_DIB);
    EXTENSIONS_MAPPING.put(".gif", XWPFDocument.PICTURE_TYPE_GIF);
    EXTENSIONS_MAPPING.put(".tiff", XWPFDocument.PICTURE_TYPE_TIFF);
    EXTENSIONS_MAPPING.put(".eps", XWPFDocument.PICTURE_TYPE_EPS);
    EXTENSIONS_MAPPING.put(".bmp", XWPFDocument.PICTURE_TYPE_BMP);
    EXTENSIONS_MAPPING.put(".wpg", XWPFDocument.PICTURE_TYPE_WPG);
  }

  public static XWPFDocument load(String originalPath) {
    try {
      return new XWPFDocument(new FileInputStream(new File(originalPath)));
    }catch(IOException exception) {
      return null;
    }
  }

  public static XWPFDocument load(InputStream is) {
    try {
      return new XWPFDocument(is);
    }catch(IOException exception) {
      return null;
    }
  }

  public static void write(String writePath, XWPFDocument document)
      throws IOException, SecurityException {
    document.write(new FileOutputStream(new File(writePath)));
  }

  public static void clearRuns(XWPFParagraph paragraph) {
    int runsNo = paragraph.getRuns().size();
    for (int index = 0; index < runsNo; ++index) {
      paragraph.removeRun(0);
    }
  }

  public static void createPicture(XWPFParagraph paragraph, String blipId, int id, long width, long height) {
    createPicture(paragraph, blipId, id, width, height, 0, 0);
  }

  public static void createPictureCxCy(XWPFParagraph paragraph, String blipId, int id, long cx, long cy) {
    createPictureCxCy(paragraph, blipId, id, cx, cy, 0, 0);
  }

  public static void createPicture(XWPFParagraph paragraph, String blipId, int id, long width, long height, long x, long y) {
    long cx = Units.toEMU(width);
    long cy = Units.toEMU(height);

    createPictureCxCy(paragraph, blipId, id, cx, cy, x, y);
  }

  public static void createPictureCxCy(XWPFParagraph paragraph, String blipId, int id, long cx, long cy, long x, long y) {

    CTInline inline = paragraph.createRun().getCTR().addNewDrawing().addNewInline();

    String picXml = "" +
        "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "         <pic:nvPicPr>" +
        "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
        "            <pic:cNvPicPr/>" +
        "         </pic:nvPicPr>" +
        "         <pic:blipFill>" +
        "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
        "            <a:stretch>" +
        "               <a:fillRect/>" +
        "            </a:stretch>" +
        "         </pic:blipFill>" +
        "         <pic:spPr>" +
        "            <a:xfrm>" +
        "               <a:off x=\"" + x + "\" y=\"" + y + "\"/>" +
        "               <a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/>" +
        "            </a:xfrm>" +
        "            <a:prstGeom prst=\"rect\">" +
        "               <a:avLst/>" +
        "            </a:prstGeom>" +
        "         </pic:spPr>" +
        "      </pic:pic>" +
        "   </a:graphicData>" +
        "</a:graphic>";

//    CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();
    XmlToken xmlToken = null;
    try
    {
      xmlToken = XmlToken.Factory.parse(picXml);
    }
    catch(XmlException xe)
    {
      xe.printStackTrace();
    }
    inline.set(xmlToken);
//    graphicData.set(xmlToken);

    inline.setDistT(0);
    inline.setDistB(0);
    inline.setDistL(0);
    inline.setDistR(0);

    CTPositiveSize2D extent = inline.addNewExtent();
    extent.setCx(cx);
    extent.setCy(cy);

    CTNonVisualDrawingProps docPr = inline.addNewDocPr();
    docPr.setId(id);
    docPr.setName("Picture " + id);
    docPr.setDescr("Generated");
  }
}