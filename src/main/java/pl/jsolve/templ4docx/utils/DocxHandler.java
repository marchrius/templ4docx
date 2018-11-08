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

  public static void createPicture(
      XWPFRun sourceRun, XWPFRun destRun, String blipId, int id, long width, long height) {
    createPicture(sourceRun, destRun, blipId, id, width, height, 0, 0);
  }

  public static void createPictureCxCy(
      XWPFRun sourceRun, XWPFRun destRun, String blipId, int id, long cx, long cy) {
    createPictureCxCy(sourceRun, destRun, blipId, id, cx, cy, 0, 0);
  }

  public static void createPicture(
      XWPFRun sourceRun, XWPFRun destRun, String blipId, int id, long width, long height, long x, long y) {
    long cx = Units.toEMU(width);
    long cy = Units.toEMU(height);

    createPictureCxCy(sourceRun, destRun, blipId, id, cx, cy, x, y);
  }

  public static void createPictureCxCy(
      XWPFRun sourceRun, XWPFRun destRun, String blipId, int id, long cx, long cy, long x, long y) {

    CTDrawing drawing = sourceRun.getCTR().getDrawingArray(0);

    boolean isAnchor = false, isInline = false;

    CTAnchor aDst = null, aSrc = null;
    CTInline iDst = null, iSrc = null;

    CTGraphicalObject graphic = null;
    CTPositiveSize2D extent = null;
    CTNonVisualDrawingProps docPr = null;

    if (drawing.sizeOfAnchorArray() > 0) {
      aDst = destRun.getCTR().addNewDrawing().addNewAnchor();
      isAnchor = true;
    } else if (drawing.sizeOfInlineArray() > 0) {
      iDst = destRun.getCTR().addNewDrawing().addNewInline();
      isInline = true;
    }

    String picXml = "" +
        "<w:graphic xmlns:w=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
        "  <w:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "    <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
        "      <pic:nvPicPr>" +
        "        <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +
        "        <pic:cNvPicPr/>" +
        "      </pic:nvPicPr>" +
        "      <pic:blipFill>" +
        "        <w:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
        "        <w:stretch>" +
        "          <w:fillRect/>" +
        "        </w:stretch>" +
        "      </pic:blipFill>" +
        "      <pic:spPr>" +
        "        <w:xfrm>" +
        "          <w:off x=\"0\" y=\"0\"/>" +
        "          <w:ext cx=\"0\" cy=\"0\"/>" +
        "        </w:xfrm>" +
        "        <w:prstGeom prst=\"rect\">" +
        "          <w:avLst/>" +
        "        </w:prstGeom>" +
        "      </pic:spPr>" +
        "    </pic:pic>" +
        "  </w:graphicData>" +
        "</w:graphic>";

    XmlToken xmlToken = null;
    try
    {
      xmlToken = XmlToken.Factory.parse(picXml);
    }
    catch(XmlException xe)
    {
      xe.printStackTrace();
    }

    if (isAnchor) {
      aSrc = drawing.getAnchorArray(0);

//      graphic = aDst.addNewGraphic();
//      graphic.set(xmlToken);
//
//      aDst.setGraphic(graphic);

      aDst.set(xmlToken);

      aDst.setDistB(aSrc.getDistB());
      aDst.setDistL(aSrc.getDistL());
      aDst.setDistR(aSrc.getDistR());
      aDst.setDistT(aSrc.getDistT());
      aDst.setSimplePos(aSrc.getSimplePos());
      aDst.setSimplePos2(aSrc.getSimplePos2());
      aDst.setAllowOverlap(aSrc.getAllowOverlap());
      aDst.setBehindDoc(aSrc.getBehindDoc());
      aDst.setHidden(aSrc.getHidden());
      aDst.setPositionH(aSrc.getPositionH());
      aDst.setPositionV(aSrc.getPositionV());
      aDst.setRelativeHeight(aSrc.getRelativeHeight());
      aDst.setCNvGraphicFramePr(aSrc.getCNvGraphicFramePr());
      aDst.setDocPr(aSrc.getDocPr());
      aDst.setEffectExtent(aSrc.getEffectExtent());

      extent = aDst.getExtent() != null ? aDst.getExtent() : aDst.addNewExtent();
      docPr = aDst.getDocPr() != null ? aDst.getDocPr() : aDst.addNewDocPr();
    } else if (isInline) {
      iSrc = drawing.getInlineArray(0);

      graphic = iDst.addNewGraphic();
      graphic.set(xmlToken);

      iDst.setGraphic(graphic);

      iDst.setDistB(iSrc.getDistB());
      iDst.setDistL(iSrc.getDistL());
      iDst.setDistR(iSrc.getDistR());
      iDst.setDistT(iSrc.getDistT());
      iDst.setCNvGraphicFramePr(iSrc.getCNvGraphicFramePr());
      iDst.setDocPr(iSrc.getDocPr());

      extent = iDst.getExtent() != null ? iDst.getExtent() : iDst.addNewExtent();
      docPr = iDst.getDocPr() != null ? iDst.getDocPr() : iDst.addNewDocPr();
    }

    if (extent != null) {
      extent.setCx(cx);
      extent.setCy(cy);
    }

    if (docPr != null) {
//      docPr.setName("Picture " + id);
//      docPr.setDescr("Generated");
      docPr.setId(id);
    }
  }
}