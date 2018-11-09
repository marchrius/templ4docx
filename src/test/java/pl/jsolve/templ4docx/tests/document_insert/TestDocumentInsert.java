package pl.jsolve.templ4docx.tests.document_insert;

import org.junit.Test;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.DocumentVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 */
public class TestDocumentInsert extends AbstractDocumentInsertTest {

  @Test
  public void test() throws IOException {
    String documentFileName = "document-insert/document-template";
    InputStream is = getClass().getClassLoader().getResourceAsStream(documentFileName + ".docx");

    String subDocumentFileName = "document-insert/long-names";
    InputStream subIs = getClass().getClassLoader()
        .getResourceAsStream(subDocumentFileName + ".docx");

    Docx docx = new Docx(is);
    is.close();

    Docx subDocx = new Docx(subIs);
    subIs.close();

    docx.setVariablePattern(new VariablePattern("#{", "}"));

    Variables var = new Variables();

    var.addDocumentVariable("#{document.1}", subDocx);

    var.addDocumentVariable("#{document.2}", subDocx);

    var.addDocumentVariable("#{document.3}", subDocx);

    var.addTextVariable(new TextVariable("#{cost}", "1234.56"));

    var.addTextVariable("#{variableWithVeryVeryLongName01}", "Short");

    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryLongName02}", "Medium");

    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}",
        "Long");

    var.addTextVariable(new TextVariable("#{form.bankIBAN}", "1234.56"));

    var.addTextVariable("#{var01}", "Oh");
    var.addTextVariable("#{var02}", "Welcome");
    var.addTextVariable("#{var03}", "Guest");

    List<String> placeholders = docx.findVariables();

    docx.fillTemplate(var);

    String processedPath = getRelative("document-template-processed.docx");

    System.out.println(processedPath);

    docx.save(processedPath);

    String text = docx.readTextContent();

    System.out.println(text);

    assertEquals(
        "This is test simple template with three variables: Oh, Welcome, Guest.\n"
            + "\n"
            + "This is test simple template with three variables with long names: Short, Medium, Long.\n"
            + "\n"
            + "This is test simple template with three variables with long names: Short, Medium, Long.\n"
            + "\n"
            + "This is test simple template with three variables with long names: Short, Medium, Long.\n"
            + "\n"
            + "This document will cost you $ 1234.56\n"
            + "\n"
            + "IBAN / Account #:  1234.56",
        text.trim());
  }

}
