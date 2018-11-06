package pl.jsolve.templ4docx.tests.document_insert;

import org.junit.Before;

import java.io.File;
import java.io.InputStream;
import pl.jsolve.templ4docx.tests.AbstractTest;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class AbstractDocumentInsertTest extends AbstractTest {

  @Before
  public void createTmpDirectory() {
    this.setTestPath(
        String.format("%s%s", File.separator, "document-insert"));
    super.createTmpDirectory();
  }
}
