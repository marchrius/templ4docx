package pl.jsolve.templ4docx.tests;

import java.io.File;
import java.io.InputStream;
import java.util.logging.Logger;

public class AbstractTest {

  protected Logger logger = Logger.getLogger(AbstractTest.class.getSimpleName());

  protected String tmpDir = null;

  private String instanceTempDirectory = null;
  private String testPath = null;

  public void createTmpDirectory() {
    this.tmpDir = tmpDir == null ? System.getProperty("java.io.tmpdir") : tmpDir;
    this.tmpDir = this.tmpDir.charAt(this.tmpDir.length() - 1) == File.separatorChar ? this.tmpDir.substring(0, this.tmpDir.length() - 1) : this.tmpDir;
    this.instanceTempDirectory = String.format("%s%s%s%s", tmpDir, File.separator, "templ4docx-tests", getTestPath());
    File testDir = new File(this.instanceTempDirectory);
    if (!testDir.exists()) {
      if (!testDir.mkdirs()) {
        logger.warning("Directory \"" + testDir + "\" could not be created");
      }
    }
    logger.info("Instance temp directory: " + this.getTempDirectory());
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

  public File getTempDirectory() {
    return new File(instanceTempDirectory);
  }

  public String getRelative(String... paths) {
    StringBuilder sb = new StringBuilder();
    for (String s : paths) {
      sb.append(s).append(File.separator);
    }
    sb.reverse().deleteCharAt(0).reverse();
    return String.format("%s%s%s", instanceTempDirectory, File.separator, sb);
  }
}
