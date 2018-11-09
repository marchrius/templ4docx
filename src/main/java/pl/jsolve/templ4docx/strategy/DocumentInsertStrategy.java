package pl.jsolve.templ4docx.strategy;

import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.NumberingUtil;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTNumPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import pl.jsolve.sweetener.collection.Maps;
import pl.jsolve.templ4docx.cleaner.ParagraphCleaner;
import pl.jsolve.templ4docx.insert.DocumentInsert;
import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.insert.ParagraphInsert;
import pl.jsolve.templ4docx.utils.DocxHandler;
import pl.jsolve.templ4docx.variable.DocumentVariable;
import pl.jsolve.templ4docx.variable.Variable;

public class DocumentInsertStrategy implements InsertStrategy {

  private ParagraphCleaner paragraphCleaner;

  public DocumentInsertStrategy(ParagraphCleaner paragraphCleaner) {
    this.paragraphCleaner = paragraphCleaner;
  }

  @Override
  public void insert(Insert insert, Variable variable) {
    if (!(insert instanceof DocumentInsert)) {
      return;
    }
    if (!(variable instanceof DocumentVariable)) {
      return;
    }

    insertImpl((DocumentInsert) insert, (DocumentVariable) variable);
  }

  private void insertImpl(DocumentInsert insert, DocumentVariable variable) {
    XWPFParagraph templateParagraph = insert.getParagraph();
//    List<IBodyElement> bodyElements = getReverseListOfBodyElements(variable.getDocument());
    List<IBodyElement> bodyElements = variable.getDocument().getBodyElements();

    XWPFDocument mainDocument = templateParagraph.getDocument();
    XWPFDocument subDocument = insert.getDocumentParent();

    List<XWPFParagraph> paragraphs = insert.getDocumentParent().getParagraphs();
    XWPFParagraph prevParagraph = null;
    XWPFParagraph nextParagraph = null;

    for (int i = 0, len = paragraphs.size(); i < len; i++) {
      if (paragraphs.get(i).equals(templateParagraph)) {
        if (i - 1 >= 0) {
          prevParagraph = paragraphs.get(i - 1);
        }
        if (i + 1 < len) {
          nextParagraph = paragraphs.get(i + 1);
        }
        break;
      }
    }

//    if (nextParagraph == null && prevParagraph == null) {
//      System.out.println("Paragraph is first and last");
//    } else if (nextParagraph == null) {
//      System.out.println("Paragraph is last");
//    } else if (prevParagraph == null) {
//      System.out.println("Paragraph is first");
//    } else {
//      System.out.println("Paragraph is contained");
//    }

//		List<IBodyElement> bodyElements = getReverseListOfBodyElements(documentVariable.getDocument());

    XmlCursor templateCursor = templateParagraph.getCTP().newCursor();

    XWPFParagraph firstParagraph = null;

    String documentIdentifier = subDocument.getProperties().getCoreProperties().getIdentifier();

    for (IBodyElement bodyElement : bodyElements) {
      BodyElementType bodyElementType = bodyElement.getElementType();

      boolean clonedRunOnly = false;

      // Copying numbering from src document to match inserted paragraph styles
//      copyNumbering(documentIdentifier, subDocument, mainDocument);

      if (bodyElementType.equals(BodyElementType.PARAGRAPH)) {

        XWPFParagraph paragraph = (XWPFParagraph) bodyElement;

        XWPFParagraph newParagraph;

//        XWPFParagraph newParagraph = templateParagraph.getBody().insertNewParagraph(templateCursor);
//        XWPFParagraph newParagraph = templateParagraph.getDocument().createParagraph();
//        XWPFParagraph newParagraph = new XWPFParagraph(templateParagraph.getCTP(), templateParagraph.getBody());

        // Copying styles from src document to match inserted paragraph styles
        copyStyle(subDocument, mainDocument, subDocument.getStyles().getStyle(((XWPFParagraph) bodyElement).getStyleID()));

        // This will replace the template paragraph or, if necessary, add new one
        if (firstParagraph == null) {
          newParagraph = templateParagraph;
          firstParagraph = newParagraph;
          // move cursor to next for the next insertNewParagraph
          templateCursor = newParagraph.getCTP().newCursor();
        } else if (insert.isInAList() || firstParagraph.getNumID() != null) {
          newParagraph = firstParagraph;
          XWPFRun run = firstParagraph.createRun();
          run.addBreak();
          cloneParagraphRunInParagraph(firstParagraph, paragraph);
          clonedRunOnly = true;
        } else {
          // move the cursor to next
          templateCursor.toNextSibling();
          // and add the new paragraph
          newParagraph = templateParagraph.getBody().insertNewParagraph(templateCursor);
          // move cursor to next for the next insertNewParagraph
          templateCursor = newParagraph.getCTP().newCursor();
        }

        if (!clonedRunOnly) {
          cloneParagraph(newParagraph, paragraph, templateParagraph);
        }

        // if is first insertion, copy the numerating properties from paragraph, if any
        if (newParagraph == firstParagraph && (insert.isInAList() || newParagraph.getNumID() != null)) {
          clearParagraphNum(newParagraph);
          cloneParagraphNum(newParagraph, prevParagraph, nextParagraph);
//        } else {
//          clearParagraphNum(newParagraph);
        }

      } else if (bodyElementType.equals(BodyElementType.TABLE)) {
        templateCursor = templateParagraph.getCTP().newCursor();
        XWPFTable table = (XWPFTable) bodyElement;
        XWPFTable newTable = templateParagraph.getDocument().insertNewTbl(templateCursor);

        copyStyle(subDocument, mainDocument, subDocument.getStyles().getStyle(((XWPFTable) bodyElement).getStyleID()));

        cloneTable(newTable, table);
      }
    }
    clean(templateParagraph, insert);
  }

  private void clearParagraphNum(XWPFParagraph dest) {
    if (dest.getCTP() != null) {
      CTPPr ctpPr = dest.getCTP().getPPr();
      if (ctpPr == null) {
        ctpPr = dest.getCTP().addNewPPr();
      }
      CTNumPr ctNumPr = ctpPr.getNumPr();
      if (ctNumPr == null) {
        ctNumPr = ctpPr.addNewNumPr();
      }
      CTDecimalNumber ctNumIlvl = ctNumPr.getIlvl();
      if (ctNumIlvl == null) {
        ctNumIlvl = ctNumPr.addNewIlvl();
      }
      ctNumIlvl.setVal(null);
    }
    dest.setNumID(null);
    dest.setStyle(null);
  }

  private void cloneParagraphNum(XWPFParagraph dest, XWPFParagraph prevSource, XWPFParagraph nextSource) {
    if (dest == null) {
      return;
    }

    if (prevSource == null && nextSource == null) {
      dest.setNumID(BigInteger.ZERO);
      return;
    }

    XWPFParagraph source = nextSource != null && nextSource.getNumID() != null ? nextSource : prevSource != null && prevSource.getNumID() != null ? prevSource : null;

    if (source != null) {
      dest.setNumID(source.getNumID());
      dest.setStyle(source.getStyle());

      if (dest.getCTP() != null) {
        CTPPr ctpPr = dest.getCTP().getPPr();
        if (ctpPr == null) {
          ctpPr = dest.getCTP().addNewPPr();
        }
        CTNumPr ctNumPr = ctpPr.getNumPr();
        if (ctNumPr == null) {
          ctNumPr = ctpPr.addNewNumPr();
        }
        CTDecimalNumber ctNumIlvl = ctNumPr.getIlvl();
        if (ctNumIlvl == null) {
          ctNumIlvl = ctNumPr.addNewIlvl();
        }
        ctNumIlvl.setVal(source.getNumIlvl());
      }
    }
  }

  private void cloneParagraphRunInParagraph(XWPFParagraph clone, XWPFParagraph source) {
    if (clone == null || source == null) {
      return;
    }

    // we do not copy properties as the clone paragraph is meant to be already set up

//    CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
//    pPr.set(source.getCTP().getPPr());

    // Cloning all runs into this paragraph
    for (XWPFRun r : source.getRuns()) {
      XWPFRun newRun = clone.createRun();
      cloneRun(newRun, r);
    }
  }

  private void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
    cloneParagraph(clone, source, null);
  }

  private void cloneParagraph(XWPFParagraph clone, XWPFParagraph source, XWPFParagraph propertiesTemplate) {
    if (clone == null || source == null) {
      return;
    }

    String sourceIdentifier = source.getDocument().getProperties().getCoreProperties().getIdentifier();

    CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();

    // copy the source paragraph properties into cloned one
    pPr.set(source.getCTP().getPPr());

//    Map<BigInteger, XWPFNum> numbs = getNums(clone.getDocument());
//    XWPFNumbering cloneNumbs = clone.getDocument().getNumbering();

//    if (clone.getNumID() != null) {
//      XWPFAbstractNum abstractNum = cloneNumbs.getAbstractNum(clone.getNumID());
//      String str = sourceIdentifier + " " + abstractNum.getCTAbstractNum().getLvlArray(0).getPStyle().getVal();
//      XWPFNum newNum = numbs.get(str);
//      clone.setNumID(newNum.getCTNum().getNumId());
//    }

//    CTOnOff keepNext = CTOnOff.Factory.newInstance();
//    keepNext.setVal(STOnOff.ON);
//    pPr.setKeepNext(keepNext);

    for (XWPFRun r : source.getRuns()) {
      XWPFRun newRun = clone.createRun();
      cloneRun(newRun, r);
    }
  }

  private void cloneRun(XWPFRun clone, XWPFRun source) {
    if (clone == null || source == null) {
      return;
    }

    CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
    rPr.set(source.getCTR().getRPr());
    clone.setText(source.getText(0));

    // Checking for image duplication
    boolean hasImages = false;

    // Extract image from source docx file and insert into destination docx file.

    // You need next code when you want to call XWPFParagraph.removeRun().
//    dstPr.createRun();

    if (source.getEmbeddedPictures() != null && source.getEmbeddedPictures().size() > 0) {
      hasImages = true;
    }

    if (hasImages) {
      for (XWPFPicture pic : source.getEmbeddedPictures()) {

//        XWPFPictureData pictureData = pic.getPictureData();
//
//        byte[] data = pictureData.getData();
//
//        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
//        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
//        // This x and y are relative to cx and cy
//        long x = pic.getCTPicture().getSpPr().getXfrm().getExt().xgetCx().getLongValue();
//        long y = pic.getCTPicture().getSpPr().getXfrm().getExt().xgetCy().getLongValue();
//
        try {

//          // not working. DO NOT USE
//          clone.addPicture(new ByteArrayInputStream(data), pictureData.getPictureType(), pictureData.getFileName(), Units.pointsToPixel(Units.toPoints(cx)), Units.pointsToPixel(Units.toPoints(cy)));

//          String blipId = clone.getDocument().addPictureData(data, pictureData.getPictureType());
//          DocxHandler.createPictureCxCy(source, clone, blipId,
//              clone.getDocument().getNextPicNameNumber(pictureData.getPictureType()), cx, cy);

          DocxHandler.clonePicture(source, clone, source.getEmbeddedPictures().indexOf(pic));

//        } catch (IOException e1) {
//          e1.printStackTrace();
        } catch (InvalidFormatException e1) {
          e1.printStackTrace();
        } catch (IOException e) {
          e.printStackTrace();
        }
      }

      for (XWPFPicture sPic : source.getEmbeddedPictures()) {
        for (XWPFPicture cPic : clone.getEmbeddedPictures()) {
          if (cPic.getPictureData().getChecksum().equals(sPic.getPictureData().getChecksum())) {
            cPic.getCTPicture().set(sPic.getCTPicture());
          }
        }
      }
    }
//    else {
//      int pos = destDoc.getParagraphs().size() - 1;
//      destDoc.setParagraph(srcPr, pos);
//    }
  }

  private void cloneTable(XWPFTable clone, XWPFTable source) {
    if (clone == null || source == null) {
      return;
    }

    CTTblPr tblPr = clone.getCTTbl().addNewTblPr();
    tblPr.set(source.getCTTbl().getTblPr());

    for (int i = 0; i < source.getRows().size(); i++) {

      XWPFTableRow newRow = clone.getRow(i);

      if (newRow == null) {
        newRow = clone.createRow();
      }

      cloneRow(newRow, source.getRows().get(i));
    }
  }

  private void  cloneRow(XWPFTableRow clone, XWPFTableRow source) {
    if (clone == null || source == null) {
      return;
    }

    CTRow ctRow = clone.getCtRow();
    ctRow.set(source.getCtRow());
  }

  private List<IBodyElement> getReverseListOfBodyElements(XWPFDocument document) {
    List<IBodyElement> bodyElements = new ArrayList<IBodyElement>(document.getBodyElements());
    Collections.reverse(bodyElements);
    return bodyElements;
  }

  private void clean(XWPFParagraph paragraph, DocumentInsert insert) {
    for (XWPFRun run : paragraph.getRuns()) {
      String text = run.getText(0);
      if (StringUtils.contains(text, insert.getKey().getKey())) {
        text = StringUtils.replace(text, insert.getKey().getKey(), "");
        run.setText(text, 0);
      }
    }
    if (DocxHandler.isEmptyParagraph(paragraph, true)) {
      paragraphCleaner.add(new DocumentInsert(insert.getKey(), paragraph, insert.getCellParent(), insert.getDocumentParent(), insert.isInAList()));
//      paragraph.getDocument().removeBodyElement(paragraph.getDocument().getPosOfParagraph(paragraph));
    }
  }

  // Copy Styles of Table and Paragraph.
  private static void copyStyle(XWPFDocument srcDoc, XWPFDocument destDoc, XWPFStyle style)
  {
    if (destDoc == null || style == null)
      return;

    if (destDoc.getStyles() == null) {
      destDoc.createStyles();
    }

    List<XWPFStyle> usedStyleList = srcDoc.getStyles().getUsedStyleList(style);
    for (XWPFStyle xwpfStyle : usedStyleList) {
      destDoc.getStyles().addStyle(xwpfStyle);
    }
  }

  // Copy Numbering of Table and Paragraph.
  private static void copyNumbering(String prefix, XWPFDocument srcDoc, XWPFDocument destDoc)
  {
    if (destDoc == null || srcDoc == null)
      return;

    if (destDoc.getNumbering() == null) {
      destDoc.createNumbering();
    }

    String sanitizedPrefix = StringUtils.trimToEmpty(prefix);

    sanitizedPrefix = sanitizedPrefix.isEmpty() ? "" : sanitizedPrefix + " ";

    Map<BigInteger, XWPFNum> mapNumberings = getNums(srcDoc);

    for (Map.Entry<BigInteger, XWPFNum> entry : mapNumberings.entrySet()) {
      XWPFNum num = entry.getValue();
      XWPFAbstractNum abstractNum = srcDoc.getNumbering().getAbstractNum(num.getCTNum().getAbstractNumId().getVal());
      BigInteger newNumId =  destDoc.getNumbering().addNum(num);
      XWPFNum newNum = destDoc.getNumbering().getNum(newNumId);
      XWPFAbstractNum newAbstractNum = destDoc.getNumbering().getAbstractNum(num.getCTNum().getAbstractNumId().getVal());

//      .getCTNum().getLvlOverrideArray(0).getLvl().getPStyle().setVal(sanitizedPrefix + " " + newNum.getCTNum().getLvlOverrideArray(0).getLvl().getPStyle().getVal());
    }

  }

  private static Map<BigInteger, XWPFNum> getNums(XWPFDocument doc) {

    Map<BigInteger, XWPFNum> styles = Maps.newHashMap();
    XWPFNumbering numbering = doc.getNumbering();

    for (XWPFNum num : NumberingUtil.getNums(numbering)) {
      if (num != null) {
        styles.put(num.getCTNum().getNumId(), num);
      }
    }

    return styles;
  }

  public void cleanParagraphs() {
    for (ParagraphInsert paragraph : paragraphCleaner.getParagraphs()) {
      try {
//        paragraph.deleteMe();
      } catch (Exception ex) {
        // do nothing, row doesn't exist
      }
    }
  }
}
