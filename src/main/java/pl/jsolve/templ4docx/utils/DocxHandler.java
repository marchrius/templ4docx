package pl.jsolve.templ4docx.utils;

import static org.apache.poi.ooxml.POIXMLTypeLoader.DEFAULT_XML_OPTIONS;
import static org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute.Space;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import javax.xml.namespace.QName;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.IRunBody;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeaderFooter;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.values.XmlAnyTypeImpl;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlip;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBlipFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObject;
import org.openxmlformats.schemas.drawingml.x2006.main.CTGraphicalObjectData;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualPictureProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPoint2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPresetGeometry2D;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTransform2D;
import org.openxmlformats.schemas.drawingml.x2006.main.STShapeType;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPictureNonVisual;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTAnchor;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDrawing;
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

  public static XWPFPicture clonePicture(XWPFRun sourceRun, XWPFRun destRun, int index) throws InvalidFormatException, IOException {
    String relationId;
    XWPFPictureData picData;
    IRunBody parent = destRun.getParent();

    CTAnchor anchor = null;
    CTInline inline = null;
    boolean isAnchor;

    CTAnchor sAnchor = null;
    CTInline sInline = null;

    XWPFPicture sourcePicture = sourceRun.getEmbeddedPictures().get(index);
    XWPFPictureData sourcePictureData = sourcePicture.getPictureData();
    CTDrawing sourceDrawing = sourceRun.getCTR().getDrawingArray(0);

    if (sourceDrawing.sizeOfInlineArray() > 0) {
      isAnchor = false;
    } else if (sourceDrawing.sizeOfAnchorArray() > 0) {
      isAnchor = true;
    } else {
      return null;
    }

    ByteArrayInputStream pictureData = new ByteArrayInputStream(sourcePictureData.getData());

    if (parent.getPart() instanceof XWPFHeaderFooter) {
      XWPFHeaderFooter headerFooter = (XWPFHeaderFooter) parent.getPart();
      relationId = headerFooter.addPictureData(pictureData, sourcePictureData.getPictureType());
      picData = (XWPFPictureData) headerFooter.getRelationById(relationId);
    } else {
      @SuppressWarnings("resource")
      XWPFDocument doc = parent.getDocument();
      relationId = doc.addPictureData(pictureData, sourcePictureData.getPictureType());
      picData = (XWPFPictureData) doc.getRelationById(relationId);
    }

    // Create the drawing entry for it
//    try {
      CTDrawing destDrawing = destRun.getCTR().addNewDrawing();

      if (isAnchor) {
        sAnchor = sourceDrawing.getAnchorArray(0);
        anchor = destDrawing.addNewAnchor();
        anchor.set(sAnchor);
      } else {
        sInline = sourceDrawing.getInlineArray(0);
        inline = destDrawing.addNewInline();
        inline.set(sInline);
      }

      // Do the fiddly namespace bits on the inline
      // (We need full control of what goes where and as what)
//      String xml =
//          "<a:graphic xmlns:a=\"" + CTGraphicalObject.type.getName().getNamespaceURI() + "\">" +
//              "<a:graphicData uri=\"" + CTPicture.type.getName().getNamespaceURI() + "\">" +
//              "<pic:pic xmlns:pic=\"" + CTPicture.type.getName().getNamespaceURI() + "\" />" +
//              "</a:graphicData>" +
//              "</a:graphic>";
//
//      XmlToken xmlToken = null;
//      org.w3c.dom.Document doc;
//
//      InputSource is = new InputSource(new StringReader(xml));
//      doc = DocumentHelper.readDocument(is);
//      xmlToken = XmlToken.Factory.parse(doc.getDocumentElement(), DEFAULT_XML_OPTIONS);
//
//      if (xmlToken == null) {
//        throw new IOException();
//      }
//
//      if (isAnchor) {
//        anchor.set(xmlToken);
//      } else {
//        inline.set(xmlToken);
//      }

      // Setup the inline
      if (isAnchor) {
//        anchor.setDistT(0);
//        anchor.setDistR(0);
//        anchor.setDistB(0);
//        anchor.setDistL(0);
      } else {
//        inline.setDistT(0);
//        inline.setDistR(0);
//        inline.setDistB(0);
//        inline.setDistL(0);
      }

      CTNonVisualDrawingProps docPr = null;
      CTPositiveSize2D extent = null;
      CTGraphicalObject graphic = null;

      if (isAnchor) {
        docPr = anchor.getDocPr();
      } else {
        docPr = inline.getDocPr();
      }

      long id = parent.getDocument().getNextPicNameNumber(sourcePictureData.getPictureType());

      docPr.setId(id);
      /* This name is not visible in Word 2010 anywhere. */
      docPr.setName("Drawing " + id);
      docPr.setDescr(sourcePictureData.getFileName());

      if (isAnchor) {
        extent = anchor.getExtent();
      } else {
        extent = inline.getExtent();
      }

      extent.setCx(sourcePicture.getCTPicture().getSpPr().getXfrm().getExt().getCx());
      extent.setCy(sourcePicture.getCTPicture().getSpPr().getXfrm().getExt().getCy());

      // Grab the picture object
      if (isAnchor) {
        graphic = anchor.getGraphic();
      } else {
        graphic = inline.getGraphic();
      }

      CTGraphicalObjectData graphicData = graphic.getGraphicData();
      CTPicture pic = DocxHandler.getCTPictures(graphicData).get(0);

      // Set it up
      CTPictureNonVisual nvPicPr = pic.getNvPicPr();

      CTNonVisualDrawingProps cNvPr = nvPicPr.getCNvPr();

      /* use "0" for the id. See ECM-576, 20.2.2.3 */
      cNvPr.setId(0L);
      /* This name is not visible in Word 2010 anywhere */
      cNvPr.setName("Picture " + id);
      cNvPr.setDescr(sourcePictureData.getFileName());

      CTNonVisualPictureProperties cNvPicPr = nvPicPr.getCNvPicPr();

      CTBlipFillProperties blipFill = pic.getBlipFill();
      CTBlip blip = blipFill.getBlip();
//      blip.setEmbed(parent.getPart().getRelationId(picData));
      blip.setEmbed(relationId);
      if (blip.getExtLst() != null && blip.getExtLst().sizeOfExtArray() > 0) {
        for (int i = 0, len = blip.getExtLst().sizeOfExtArray(); i < len; i++) {
          blip.getExtLst().removeExt(i);
        }
      }

      blipFill.getStretch().getFillRect();

      CTShapeProperties spPr = pic.getSpPr() != null ? pic.getSpPr() : pic.addNewSpPr();
      CTTransform2D xfrm = spPr.getXfrm() != null ? spPr.getXfrm() : spPr.addNewXfrm();

      CTPoint2D off = xfrm.getOff() != null ? xfrm.getOff() : xfrm.addNewOff();
      off.setX(0);
      off.setY(0);

      CTPositiveSize2D ext = xfrm.getExt() != null ? xfrm.getExt() : xfrm.addNewExt();
      ext.setCx(extent.getCx());
      ext.setCy(extent.getCy());

      CTPresetGeometry2D prstGeom = spPr.getPrstGeom() != null ? spPr.getPrstGeom() : spPr.addNewPrstGeom();

      prstGeom.setPrst(STShapeType.RECT);
      if (prstGeom.getAvLst() == null) {
        prstGeom.addNewAvLst();
      }

      // Fix bug: Bug 55476 - after XWPFRun.addPicture() Word considers the document as corrupted
      XmlObject[] pics = graphicData.selectChildren(new QName(CTPicture.type.getName().getNamespaceURI(), "pic"));
      pics[0].set(pic);

      // Finish up
      return new XWPFPicture(pic, destRun);
//      XWPFPicture xwpfPicture = new XWPFPicture(pic, destRun);
//      pictures.add(xwpfPicture);
//      return xwpfPicture;
//    } catch (XmlException e) {
//      throw new IllegalStateException(e);
//    } catch (SAXException e) {
//      throw new IllegalStateException(e);
//    }
  }

  public static int getPictureFormat(String filename) {
    String _sanitizedFilename = StringUtils.trimToNull(filename);

    if (_sanitizedFilename == null) {
      return 0;
    }

    String ext = StringUtils.right(_sanitizedFilename, 4);

    if (EXTENSIONS_MAPPING.containsKey(ext)) {
      return EXTENSIONS_MAPPING.get(ext);
    }

    return 0;
  }

  public static Set<Integer> getAllPictureFormats() {
    return new HashSet<Integer>(EXTENSIONS_MAPPING.values());
  }

  public static List<CTPicture> getCTPictures(XmlObject o) {
      List<CTPicture> pics = new ArrayList<CTPicture>();
      XmlObject[] picts = o.selectPath("declare namespace pic='" + CTPicture.type.getName().getNamespaceURI() + "' .//pic:pic");
      for (XmlObject pict : picts) {
        if (pict instanceof XmlAnyTypeImpl) {
          // Pesky XmlBeans bug - see Bugzilla #49934
          try {
            pict = CTPicture.Factory.parse(pict.toString(), DEFAULT_XML_OPTIONS);
          } catch (XmlException e) {
            throw new POIXMLException(e);
          }
        }
        if (pict instanceof CTPicture) {
          pics.add((CTPicture) pict);
        }
      }
      return pics;
    }

  public static boolean isEmptyParagraph(XWPFParagraph paragraph, boolean b) {
    if (paragraph.isEmpty()) return true;

    for (XWPFRun run : paragraph.getRuns()) {
      if (run.getEmbeddedPictures() != null && run.getEmbeddedPictures().size() > 0) {
        return false;
      }
      for (int i = 0, len = run.getCTR().sizeOfTArray(); i < len; i++) {
        CTText text = run.getCTR().getTArray(i);
        if (text.getSpace().equals(Space.PRESERVE) || !text.getStringValue().isEmpty()) {
          return false;
        }
      }
    }
    return true;
  }
}