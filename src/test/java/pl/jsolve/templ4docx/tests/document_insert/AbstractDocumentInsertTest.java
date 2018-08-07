package pl.jsolve.templ4docx.tests.document_insert;

import org.junit.Before;

import java.io.File;

/**
 * @author indvd00m (gotoindvdum[at]gmail[dot]com)
 *
 */
public class AbstractDocumentInsertTest {

    @Before
    public void createTmpDirectory() {
        String tmpPath = System.getProperty("java.io.tmpdir");
        String testPath = String.format("%s%s%s", tmpPath, File.separator, "cleaner");
        File testDir = new File(testPath);
        if (!testDir.exists())
            testDir.mkdirs();
    }

}
