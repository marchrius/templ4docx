package pl.jsolve.templ4docx.insert;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import pl.jsolve.templ4docx.util.Key;

public class DocumentInsert extends ParagraphInsert {

    public DocumentInsert(Key key, XWPFParagraph paragraph, XWPFTableCell cellParent, XWPFDocument documentParent, boolean inAList) {
        super(key, paragraph, cellParent, documentParent);
        this.setInAList(inAList);
    }

    public DocumentInsert(Key key, XWPFParagraph paragraph, XWPFTableCell cellParent, XWPFDocument documentParent) {
        super(key, paragraph, cellParent, documentParent);
        this.setInAList(false);
    }
}
