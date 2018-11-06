package pl.jsolve.templ4docx.tests;

import java.io.File;
import java.io.InputStream;
import java.util.logging.Logger;

public class AbstractTest {

  protected Logger logger = Logger.getLogger(AbstractTest.class.getSimpleName());

  private String tmpDir = null;
  private String testPath = null;

  public void createTmpDirectory() {
    this.tmpDir = System.getProperty("java.io.tmpdir");
    File testDir = new File(String.format("%s%s", getTmpDir(), getTestPath()));
    if (!testDir.exists()) {
      if (!testDir.mkdirs()) {
        logger.warning("Directory \"" + testDir + "\" could not be created");
      }
    }
  }

  protected InputStream loadDocx(String filename) {
    return getClass().getClassLoader().getResourceAsStream(filename + ".docx");
  }

  protected String getTestPath() {
    return testPath;
  }

  protected void setTestPath(String testPath) {
    this.testPath = testPath;
  }

  public String getTmpDir() {
    return tmpDir;
  }
}
