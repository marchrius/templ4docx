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

import static org.junit.Assert.assertEquals;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class TestDocumentInsert extends AbstractDocumentInsertTest {

    @Test
    public void test() throws IOException {
        String documentFileName = "document-insert/document-template";
        InputStream is = getClass().getClassLoader().getResourceAsStream(documentFileName + ".docx");

        String subDocumentFileName = "document-insert/long-names";
        InputStream subIs = getClass().getClassLoader().getResourceAsStream(subDocumentFileName + ".docx");

        Docx docx = new Docx(is);
        is.close();

        Docx subDocx = new Docx(subIs);
        subIs.close();

        docx.setVariablePattern(new VariablePattern("${", "}"));

        Variables var = new Variables();
        var.addDocumentVariable(new DocumentVariable("${document.1}", subDocx.getXWPFDocument()));

        var.addDocumentVariable(new DocumentVariable("${document.2}", subDocx.getXWPFDocument()));

        var.addDocumentVariable(new DocumentVariable("${document.3}", subDocx.getXWPFDocument()));

        docx.fillTemplate(var);

        String tmpPath = System.getProperty("user.dir");

        String processedPath = String.format("%s%s%s", tmpPath, File.separator,
                documentFileName + "-processed" + ".docx");

        File parentFile = new File(processedPath);
        parentFile = parentFile.getParentFile();

        if (!parentFile.exists()) {
            parentFile.mkdirs();
        }

        System.out.println(processedPath);

        docx.save(processedPath);

        String text = docx.readTextContent();

        System.out.println(text);

        assertEquals(
                "This is test simple template with three variables: ${var01}, ${var02}, ${var03}.\n" +
                        "\n" +
                        "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.\n" +
                        "\n" +
                        "\n" +
                        "\n" +
                        "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.\n" +
                        "\n" +
                        "\n" +
                        "This is test simple template with three variables with long names: ${variableWithVeryVeryLongName01}, ${variableWithVeryVeryVeryVeryVeryVeryLongName02}, ${variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}.",
                text.trim());
    }

}
