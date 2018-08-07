package pl.jsolve.templ4docx.insert;

import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import pl.jsolve.templ4docx.util.Key;

public class DocumentInsert extends Insert{

	 private XWPFParagraph paragraph;

    public DocumentInsert(Key key, XWPFParagraph paragraph) {
        super(key);
        this.paragraph = paragraph;
    }

    public XWPFParagraph getParagraph() {
        return paragraph;
    }
}
