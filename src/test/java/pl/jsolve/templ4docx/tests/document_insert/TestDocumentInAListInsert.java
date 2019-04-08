package pl.jsolve.templ4docx.tests.document_insert;

import static org.junit.Assert.assertEquals;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.util.IOUtils;
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

    String subDocumentAsListFileName = "document-insert/document-template-document-with-paragraph-as-list";
    InputStream subAsListIs = getClass().getClassLoader()
        .getResourceAsStream(subDocumentAsListFileName + ".docx");

    Docx docx = new Docx(is);
    IOUtils.closeQuietly(is);

    Docx subDocx = new Docx(subIs);
    IOUtils.closeQuietly(subIs);

    Docx subAsListDocx = new Docx(subAsListIs);
    IOUtils.closeQuietly(subAsListIs);

    docx.setVariablePattern(new VariablePattern("#{", "}"));

    Variables var = new Variables();

    var.addDocumentVariable("#{document.1.1}", subDocx, true);
    var.addDocumentVariable("#{document.1.2}", subDocx, false);

    var.addDocumentVariable("#{document.2.1}", subAsListDocx, true);
    var.addDocumentVariable("#{document.2.2}", subAsListDocx, false);

    List<Variable> listAsUniqueParam = new ArrayList<Variable>();
    List<Variable> listAsNotUniqueParam = new ArrayList<Variable>();

    List<Variable> listAsListAsUniqueParam = new ArrayList<Variable>();
    List<Variable> listAsListAsNotUniqueParam = new ArrayList<Variable>();

    listAsUniqueParam.add(new DocumentVariable("#{document.uniquepraram}", subDocx, true));
    listAsNotUniqueParam.add(new DocumentVariable("#{document.nonuniquepraram}", subDocx, false));

    listAsListAsUniqueParam.add(new DocumentVariable("#{document.uniquepraram.aslist}", subAsListDocx, true));
    listAsListAsNotUniqueParam.add(new DocumentVariable("#{document.nonuniquepraram.aslist}", subAsListDocx, false));

    var.addBulletListVariable("#{document.uniquepraram}", listAsUniqueParam);
    var.addBulletListVariable("#{document.nonuniquepraram}", listAsNotUniqueParam);

    var.addBulletListVariable("#{document.uniquepraram.aslist}", listAsListAsUniqueParam);
    var.addBulletListVariable("#{document.nonuniquepraram.aslist}", listAsListAsNotUniqueParam);

    List<String> placeholders = docx.findVariables();

    docx.fillTemplate(var);

    String processedPath = getRelative("document-template-processed.docx");

    System.out.println(processedPath);

    docx.save(processedPath);

    String text = docx.readTextContent();

    System.out.println(text);

    assertEquals(
        "Dump document insertion for document.1.1/2\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Dump document insertion for document.2.1/2\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Forced unique paragraph for document.uniqueparam\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Item 1\n"
            + "Item 2\n"
            + "Item 3\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Forced unique paragraph for document.uniqueparam.aslist\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Item 1\n"
            + "Item 2\n"
            + "Item 3\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Non-unique paragraph for document.nonuniqueparam\n"
            + "\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Item 1\n"
            + "Item 2\n"
            + "Item 3\n"
            + "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Pellentesque vulputate ultricies felis, a euismod leo suscipit eget\n"
            + "Quisque maximus dictum interdum. Nulla facilisi. Suspendisse gravida est sed auctor ullamcorper.\n"
            + "Non-unique paragraph for document.nonuniqueparam.aslist\n"
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
