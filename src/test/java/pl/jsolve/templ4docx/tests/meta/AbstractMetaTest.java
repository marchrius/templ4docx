package pl.jsolve.templ4docx.tests.meta;

import java.io.File;

import org.junit.Before;
import pl.jsolve.templ4docx.tests.AbstractTest;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public abstract class AbstractMetaTest extends AbstractTest {

  @Before
  public void createTmpDirectory() {
    this.setTestPath(
        String.format("%s%s", File.separator, "meta"));
    super.createTmpDirectory();
  }

}
