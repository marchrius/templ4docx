package pl.jsolve.templ4docx.tests.cleaner;

import java.io.File;

import org.junit.Before;
import pl.jsolve.templ4docx.tests.AbstractTest;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class AbstractCleanerTest extends AbstractTest {

  @Before
  public void createTmpDirectory() {
    this.setTestPath(
        String.format("%s%s", File.separator, "cleaner"));
    super.createTmpDirectory();
  }
}
