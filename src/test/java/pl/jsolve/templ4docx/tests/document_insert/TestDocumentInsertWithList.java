package pl.jsolve.templ4docx.tests.document_insert;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.TimeZone;
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

    Date date = new Date(704750400000L);

    SimpleDateFormat sdf = new SimpleDateFormat("dd-MM-yyyy 'at' HH:mm");

    var.addTextVariable("#{variableWithVeryVeryLongName01}", "Ciao");
    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryLongName02}", "Mondo");
    var.addTextVariable("#{variableWithVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryVeryLongName03}",
        sdf.format(date));

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
    var.addBulletListVariable("#{this_will_be_removed:document}", Collections.<Variable>emptyList());

    List<String> placeholders = docx.findVariables();

    docx.fillTemplate(var);

    String processedPath = getRelative("document-insert-with-list-processed" + ".docx");

    System.out.println(processedPath);

    docx.save(processedPath);

    String text = docx.readTextContent();

    System.out.println(text);

        assertEquals("Document insert with list\n"
                + "Insert a document into as a list element into pre-existent list\n"
                + "This is a simple text followed by a multi-level list\n"
                + "\n"
                + "Level 1\n"
                + "Level 1.1\n"
                + "Level 2\n"
                + "Level 2.1\n"
                + "Level 2.2\n"
                + "Level 2.2.1\n"
                + "Level 2.2.1.1\n"
                + "Level 3\n"
                + "Level 3.1\n"
                + "Level 3.1.1\n"
                + "Level 3.1.2\n"
                + "Level 4\n"
                + "Level 4.1\n"
                + "Level 5\n"
                + "Level 5.1\n"
                + "Level 5.1.1\n"
                + "nullSimple text. Bold text. Italic text. Strike text. Red text. Glowing text. Text withnullcarriage return. This text will be inserted into a numerating list with an image\n"
                + "The sun shines as it has never shone. This sentence is wrong, can you tell me why?\n"
                + "Level 5.1.3\n"
                + " nullSimple text. Bold text. Italic text. Strike text. Red text. Glowing text. Text withnullcarriage return. This text will be inserted into a numerating list with an image\n"
                + "The sun shines as it has never shone. This sentence is wrong, can you tell me why?\n"
                + "Testo 1 \n"
                + "Testo 2 \n"
                + "Testo 3 \n"
                + "Testo 4 \n"
                + "Testo 5 \n"
                + "Level 6\n"
                + "nullSimple text. Bold text. Italic text. Strike text. Red text. Glowing text. Text withnullcarriage return. This text will be inserted into a numerating list with an image\n"
                + "The sun shines as it has never shone. This sentence is wrong, can you tell me why?\n"
                + "This is test simple template with three variables with long names but with short values: Ciao, Mondo, " + sdf.format(date) + ".\n"
                + "\n"
                + "Here we have the document-a\n"
                + "nullSimple text. Bold text. Italic text. Strike text. Red text. Glowing text. Text withnullcarriage return. This text will be inserted into a numerating list with an image\n"
                + "The sun shines as it has never shone. This sentence is wrong, can you tell me why?\n"
                + "\n"
                + "Here we have the document b\n"
                + "This is test simple template with three variables with long names but with short values: Ciao, Mondo, " + sdf.format(date) + ".\n"
                + "\n"
                + "Level 1\n"
                + "Level 2\n"
                + "Level 3",
                text.trim());
  }

}
