package pl.jsolve.templ4docx.tests.meta;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import org.junit.Test;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class TestInvalidMetaAfterMSOfficeEdit extends AbstractMetaTest {

    @Test
    public void test() throws IOException {
        String templateFileName = "meta/invalid-meta-edited-in-ms-office";
        InputStream is = getClass().getClassLoader().getResourceAsStream(templateFileName + ".docx");

        Docx docx = new Docx(is);

        String text = docx.readTextContent();
        assertEquals(
                "This is test simple template with three variables: valu invalid_data is here e01, valu ${var02} e02, valu ${var01} e03. File edited in MS Office.",
                text.trim());

        docx.setProcessMetaInformation(true);
        docx.setVariablePattern(new VariablePattern("${", "}"));

        Variables var = new Variables();
        var.addTextVariable(new TextVariable("${var01}", "valueAfterEditInMSOffice01"));
        var.addTextVariable(new TextVariable("${var02}", "valueAfterEditInMSOffice02"));
        var.addTextVariable(new TextVariable("${var03}", "valueAfterEditInMSOffice03"));

        docx.fillTemplate(var);

        String processedPath = getRelative("invalid-meta-edited-in-ms-office-processed" + ".docx");

        docx.save(processedPath);

        text = docx.readTextContent();
        assertEquals(
                "This is test simple template with three variables: valueAfterEditInMSOffice01, valueAfterEditInMSOffice02, valueAfterEditInMSOffice03. File edited in MS Office.",
                text.trim());

        is.close();

    }

}
