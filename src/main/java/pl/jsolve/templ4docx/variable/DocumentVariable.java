package pl.jsolve.templ4docx.variable;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import pl.jsolve.templ4docx.core.Docx;

public class DocumentVariable implements Variable {

  private final String key;
  private final XWPFDocument document;
  private final boolean asUniqueParagraph;

  public DocumentVariable(String key, Docx document, boolean asUniqueParagraph) {
    this.key = key;
    this.document = document.getXWPFDocument();
    this.asUniqueParagraph = asUniqueParagraph;
  }

  public DocumentVariable(String key, Docx document) {
    this(key, document, true);
  }

  public DocumentVariable(String key, XWPFDocument document, boolean asUniqueParagraph) {
    this.key = key;
    this.document = document;
    this.asUniqueParagraph = asUniqueParagraph;
  }

  public DocumentVariable(String key, XWPFDocument document) {
    this(key, document, true);
  }

  public String getKey() {
    return key;
  }

  public XWPFDocument getDocument() {
    return document;
  }

  public boolean isAsUniqueParagraph() {
    return asUniqueParagraph;
  }
}
