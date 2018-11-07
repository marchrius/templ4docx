package pl.jsolve.templ4docx.strategy;

import java.io.ByteArrayInputStream;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
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

import pl.jsolve.templ4docx.insert.DocumentInsert;
import pl.jsolve.templ4docx.insert.Insert;
import pl.jsolve.templ4docx.utils.DocxHandler;
import pl.jsolve.templ4docx.variable.DocumentVariable;
import pl.jsolve.templ4docx.variable.Variable;

public class DocumentInsertStrategy implements InsertStrategy {

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

    if (nextParagraph == null && prevParagraph == null) {
      System.out.println("Paragraph is first and last");
    } else if (nextParagraph == null) {
      System.out.println("Paragraph is last");
    } else if (prevParagraph == null) {
      System.out.println("Paragraph is first");
    } else {
      System.out.println("Paragraph is contained");
    }

//		List<IBodyElement> bodyElements = getReverseListOfBodyElements(documentVariable.getDocument());

    boolean isFirst = true;

    XmlCursor templateCursor = templateParagraph.getCTP().newCursor();

    for (IBodyElement bodyElement : bodyElements) {
      BodyElementType bodyElementType = bodyElement.getElementType();

      if (bodyElementType.equals(BodyElementType.PARAGRAPH)) {
        XWPFParagraph paragraph = (XWPFParagraph) bodyElement;

        // Copying styles from src document to match inserted paragraph styles
//        copyStyle(subDocument, mainDocument, subDocument.getStyles().getStyle(paragraph.getStyleID()));

        XWPFParagraph newParagraph;

//        XWPFParagraph newParagraph = templateParagraph.getBody().insertNewParagraph(templateCursor);
//        XWPFParagraph newParagraph = templateParagraph.getDocument().createParagraph();
//        XWPFParagraph newParagraph = new XWPFParagraph(templateParagraph.getCTP(), templateParagraph.getBody());

        // This will replace the template paragraph or, if necessary, add new one


        if (isFirst) {
          newParagraph = templateParagraph;
        } else {
          templateCursor.toNextSibling();
          newParagraph = templateParagraph.getBody().insertNewParagraph(templateCursor);
        }

        templateCursor = newParagraph.getCTP().newCursor();

        cloneParagraph(newParagraph, paragraph);


        if (isFirst) {
          cloneParagraphNum(newParagraph, prevParagraph, nextParagraph);
        }

      } else if (bodyElementType.equals(BodyElementType.TABLE)) {
        templateCursor = templateParagraph.getCTP().newCursor();
         XWPFTable table = (XWPFTable) bodyElement;
        XWPFTable newTable = templateParagraph.getDocument().insertNewTbl(templateCursor);
        cloneTable(newTable, table);
      }


      isFirst = false;
    }
    clean(templateParagraph, insert);
  }

  private void cloneParagraphNum(XWPFParagraph dest, XWPFParagraph prevSource, XWPFParagraph nextSource) {
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

  private void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
    CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();

    pPr.set(source.getCTP().getPPr());
    for (XWPFRun r : source.getRuns()) {
      XWPFRun newRun = clone.createRun();
      cloneRun(newRun, r);
    }
  }

  private void cloneRun(XWPFRun clone, XWPFRun source) {
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

        XWPFPictureData pictureData = pic.getPictureData();

        byte[] data = pictureData.getData();

        long cx = pic.getCTPicture().getSpPr().getXfrm().getExt().getCx();
        long cy = pic.getCTPicture().getSpPr().getXfrm().getExt().getCy();
        // This x and y are relative to cx and cy
        long x = pic.getCTPicture().getSpPr().getXfrm().getExt().xgetCx().getLongValue();
        long y = pic.getCTPicture().getSpPr().getXfrm().getExt().xgetCy().getLongValue();

        try {

          String blipId = clone.getDocument().addPictureData(data, pictureData.getPictureType());
          DocxHandler.createPictureCxCy((XWPFParagraph) clone.getParent(), blipId, clone.getDocument().getNextPicNameNumber(pictureData.getPictureType()), cx, cy);

        } catch (InvalidFormatException e1) {
          e1.printStackTrace();
        }
      }
    }
//    else {
//      int pos = destDoc.getParagraphs().size() - 1;
//      destDoc.setParagraph(srcPr, pos);
//    }
  }

  private void cloneTable(XWPFTable clone, XWPFTable source) {
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

  private void cloneRow(XWPFTableRow clone, XWPFTableRow source) {
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
}
