package pl.jsolve.templ4docx.tests.variable.object;

import java.io.File;

import org.junit.Before;
import pl.jsolve.templ4docx.tests.AbstractTest;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 */
public class AbstractVariableObjectTest extends AbstractTest {

  @Before
  public void createTmpDirectory() {
    this.setTestPath(
        String.format("%s%s%s%s", File.separator, "variable", File.separator, "object"));
    super.createTmpDirectory();
  }

}
