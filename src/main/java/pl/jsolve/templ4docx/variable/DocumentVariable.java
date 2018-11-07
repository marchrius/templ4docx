package pl.jsolve.templ4docx.variable;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pl.jsolve.templ4docx.core.Docx;

public class DocumentVariable implements Variable {

  private final String key;
  private final XWPFDocument document;

  public DocumentVariable(String key, Docx document) {
    this.key = key;
    this.document = document.getXWPFDocument();
  }

  public DocumentVariable(String key, XWPFDocument document) {
    this.key = key;
    this.document = document;
  }

  public String getKey() {
    return key;
  }

  public XWPFDocument getDocument() {
    return document;
  }

}
