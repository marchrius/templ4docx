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
 *
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class TestRepeatedVariables extends AbstractMetaTest {

    @Test
    public void test() throws IOException {
        String templateFileName = "meta/repeated-variables";
        InputStream is = getClass().getClassLoader().getResourceAsStream(templateFileName + ".docx");

        Docx docx = new Docx(is);
        docx.setProcessMetaInformation(true);
        docx.setVariablePattern(new VariablePattern("${", "}"));

        Variables var = new Variables();
        var.addTextVariable(new TextVariable("${var01}", "value01"));
        var.addTextVariable(new TextVariable("${var02}", "value02"));
        var.addTextVariable(new TextVariable("${var03}", "value03"));

        docx.fillTemplate(var);

        String processedPath = getRelative("repeated-variables-processed" + ".docx");

        docx.save(processedPath);

        String text = docx.readTextContent();
        assertEquals(
                "This is test simple template with repeated variables: value01, value02, value03. Then again: value01, value02, value03. And only one variable: value02.",
                text.trim());

        is.close();

        docx = new Docx(processedPath);
        docx.setProcessMetaInformation(true);
        docx.setVariablePattern(new VariablePattern("${", "}"));

        var = new Variables();
        var.addTextVariable(new TextVariable("${var01}", "value011"));
        var.addTextVariable(new TextVariable("${var02}", "value022"));
        var.addTextVariable(new TextVariable("${var03}", "value033"));

        docx.fillTemplate(var);

        String processedPath2 = getRelative("repeated-variables-processed2" + ".docx");

        docx.save(processedPath2);

        text = docx.readTextContent();
        assertEquals(
                "This is test simple template with repeated variables: value011, value022, value033. Then again: value011, value022, value033. And only one variable: value022.",
                text.trim());
    }

}
