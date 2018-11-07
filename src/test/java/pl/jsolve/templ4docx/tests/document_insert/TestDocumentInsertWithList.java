package pl.jsolve.templ4docx.tests.document_insert;

import org.junit.Test;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import static org.junit.Assert.assertEquals;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 */
public class TestDocumentInsertWithList extends AbstractDocumentInsertTest {

  @Test
  public void test() throws IOException {
    InputStream is = loadDocx("document-insert/document-template-with-list");
    InputStream subAIs = loadDocx("document-insert/document-template-list-item-a");
    InputStream subBIs = loadDocx("document-insert/document-template-list-item-b");

    logger.info("Starting...");
    Docx docx = new Docx(is);
    is.close();

    Docx subADocx = new Docx(subAIs);
    subAIs.close();

    Docx subBDocx = new Docx(subBIs);
    subBIs.close();

    /* Configuration */
    docx.setVariablePattern(new VariablePattern("#{", "}"));

    /* END Configuration */

    Variables var = new Variables();

    var.addTextVariable("#{variableWithVeryVeryLongName01}", "Ciao");
    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryLongName02}", "Mondo");
    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}",
        (new Date()).toString());

    var.addDocumentVariable("#{document_a:document}", subADocx);
    var.addDocumentVariable("#{document_b:document}", subBDocx);

    List<Variable> bulletsList = new ArrayList<Variable>();

    bulletsList.add(new DocumentVariable("#{list}", subADocx));
    bulletsList.add(new TextVariable("#{list}", "Testo 1"));
    bulletsList.add(new TextVariable("#{list}", "Testo 2"));
    bulletsList.add(new TextVariable("#{list}", "Testo 3"));
    bulletsList.add(new TextVariable("#{list}", "Testo 4"));
    bulletsList.add(new TextVariable("#{list}", "Testo 5"));

    var.addBulletListVariable("#{list}", bulletsList);

    List<String> placeholders = docx.findVariables();

    docx.fillTemplate(var);

    String processedPath = getRelative((new Date()).getTime() + "-processed" + ".docx");

    System.out.println(processedPath);

    docx.save(processedPath);

//    String text = docx.readTextContent();
//
//    System.out.println(text);
//
//        assertEquals("This is test simple template with three variables: #{var01}, #{var02}, #{var03}.\n" +
//                "\n" +
//                "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.\n" +
//                "\n" +
//                "\n" +
//                "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.\n" +
//                "\n" +
//                "\n" +
//                "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.\n" +
//                "\n" +
//                "\n" +
//                "This document will cost you $ 1234.56\n" +
//                "\n" +
//                "IBAN / Account #:  1234.56",
//                text.trim());
  }

}
