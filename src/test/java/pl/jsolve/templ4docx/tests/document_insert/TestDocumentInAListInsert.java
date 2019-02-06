package pl.jsolve.templ4docx.tests.document_insert;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.junit.Test;
import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.core.VariablePattern;
import pl.jsolve.templ4docx.variable.DocumentVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variable;
import pl.jsolve.templ4docx.variable.Variables;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 */
public class TestDocumentInAListInsert extends AbstractDocumentInsertTest {

  @Test
  public void test() throws IOException {
    String documentFileName = "document-insert/document-template-with-document-in-list";
    InputStream is = getClass().getClassLoader().getResourceAsStream(documentFileName + ".docx");

    String subDocumentFileName = "document-insert/document-template-document-with-paragraph";
    InputStream subIs = getClass().getClassLoader()
        .getResourceAsStream(subDocumentFileName + ".docx");

    Docx docx = new Docx(is);
    is.close();

    Docx subDocx = new Docx(subIs);
    subIs.close();

    docx.setVariablePattern(new VariablePattern("#{", "}"));

    Variables var = new Variables();

    var.addDocumentVariable("#{document.1.1}", subDocx, true);
    var.addDocumentVariable("#{document.1.2}", subDocx, false);

    List<Variable> listAsUniqueParam = new ArrayList<Variable>();
    List<Variable> listAsNotUniqueParam = new ArrayList<Variable>();

    listAsUniqueParam.add(new DocumentVariable("#{document.uniquepraram}", subDocx, true));
    listAsNotUniqueParam.add(new DocumentVariable("#{document.nonuniquepraram}", subDocx, false));

    var.addBulletListVariable("#{document.uniquepraram}", listAsUniqueParam);
    var.addBulletListVariable("#{document.nonuniquepraram}", listAsNotUniqueParam);

    List<String> placeholders = docx.findVariables();

    docx.fillTemplate(var);

    String processedPath = getRelative("document-template-processed.docx");

    System.out.println(processedPath);

    docx.save(processedPath);

    String text = docx.readTextContent();

    System.out.println(text);

    assertEquals(
        "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Item 1\n"
            + "Item 2\n"
            + "Item 3\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Item 1\n"
            + "Item 2\n"
            + "Item 3\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.",
        text.trim());
  }

}
