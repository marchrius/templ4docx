package pl.jsolve.templ4docx.strategy;

import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.NumberingUtil;
import org.apache.poi.xwpf.usermodel.XWPFAbstractNum;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFNum;
import org.apache.poi.xwpf.usermodel.XWPFNumbering;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
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

    XWPFDocument subDocument = variable.getDocument();
    XWPFDocument mainDocument = templateParagraph.getDocument();

    List<IBodyElement> bodyElements = subDocument.getBodyElements();
//    List<IBodyElement> bodyElements = getReverseListOfBodyElements(variable.getDocument());
    List<XWPFParagraph> paragraphs = templateParagraph.getDocument().getParagraphs();
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

        // Copying styles from src document to match inserted paragraph styles
        DocxHandler.copyStyle(subDocument, mainDocument, subDocument.getStyles().getStyle(((XWPFParagraph) bodyElement).getStyleID()));

        // This will replace the template paragraph or, if necessary, add new one
        if (firstParagraph == null) {
          newParagraph = templateParagraph;
          firstParagraph = newParagraph;
          // move cursor to next for the next insertNewParagraph
          templateCursor = newParagraph.getCTP().newCursor();
        } else if (variable.isAsUniqueParagraph() && (insert.isInAList() || firstParagraph.getNumID() != null)) {
          newParagraph = firstParagraph;
          XWPFRun run = firstParagraph.createRun();
          run.addBreak();
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
          DocxHandler.cloneParagraph(newParagraph, paragraph, templateParagraph);
        } else {
          DocxHandler.cloneParagraph(newParagraph, paragraph, null);
        }

        // if is first insertion, copy the numerating properties from paragraph, if any
        if ((newParagraph == firstParagraph && (insert.isInAList() || newParagraph.getNumID() != null)) ||
            !variable.isAsUniqueParagraph() && (insert.isInAList() || newParagraph.getNumID() != null)) {
          clearParagraphNum(newParagraph);
          cloneParagraphNum(newParagraph, prevParagraph, nextParagraph);
          if (firstParagraph != newParagraph && !variable.isAsUniqueParagraph()) {
            keepIndentOnlyParagraphNum(newParagraph);
          }
        }

        if (newParagraph != firstParagraph || !variable.isAsUniqueParagraph()) {
          prevParagraph = newParagraph;
        }

      } else if (bodyElementType.equals(BodyElementType.TABLE)) {
        templateCursor = templateParagraph.getCTP().newCursor();
        XWPFTable table = (XWPFTable) bodyElement;
        XWPFTable newTable = templateParagraph.getDocument().insertNewTbl(templateCursor);

        DocxHandler.copyStyle(subDocument, mainDocument, subDocument.getStyles().getStyle(((XWPFTable) bodyElement).getStyleID()));

        cloneTable(newTable, table);
      }
    }
    clean(templateParagraph, insert);
  }

  private void clearParagraphNum(XWPFParagraph dest) {
    dest.setNumID(null);
    dest.setStyle(null);
    if (dest.getCTP().getPPr().getNumPr().getIlvl() == null)
      dest.getCTP().getPPr().getNumPr().addNewIlvl();
    dest.getCTP().getPPr().getNumPr().getIlvl().setVal(null);
  }

  private void keepIndentOnlyParagraphNum(XWPFParagraph dest) {
    dest.setNumID(null);
    dest.setIndentationHanging(-1);
//    if (dest.getIndentationLeft() > -1) {
//      dest.setFirstLineIndent(dest.getFirstLineIndent() + dest.getIndentationHanging());
//    } else {
//      dest.setFirstLineIndent(dest.getIndentationHanging());
//    }
//    if (dest.getIndentationLeft() > -1) {
//      dest.setIndentationLeft(dest.getIndentationLeft() + dest.getIndentationHanging());
//    } else {
//      dest.setIndentationLeft(dest.getIndentationHanging());
//    }
//    dest.setStyle(null);
//    if (dest.getCTP().getPPr().getNumPr().getIlvl() == null)
//      dest.getCTP().getPPr().getNumPr().addNewIlvl();
//    dest.getCTP().getPPr().getNumPr().getIlvl().setVal(null);
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
      if (dest.getCTP().getPPr().getNumPr().getIlvl() == null)
        dest.getCTP().getPPr().getNumPr().addNewIlvl();
      dest.getCTP().getPPr().getNumPr().getIlvl().setVal(source.getNumIlvl());
    }
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
