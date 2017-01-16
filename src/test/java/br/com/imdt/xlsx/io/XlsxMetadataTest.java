package br.com.imdt.xlsx.io;

import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import static org.hamcrest.CoreMatchers.is;
import org.junit.After;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Before;
import org.junit.Rule;
import org.junit.rules.ExpectedException;
import org.xml.sax.SAXException;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxMetadataTest {

    private XlsxMetadata metadata;

    @Rule
    public ExpectedException thrown = ExpectedException.none();

    public XlsxMetadataTest() {
    }

    @Before
    public void prepareMetadata() {
        try {
            metadata = new XlsxMetadata(ClassLoader.getSystemResourceAsStream("TestFile.xlsx"));
        } catch (InvalidFormatException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @After
    public void closesConection() {
        try {
            metadata.close();
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @Test
    public void testSheetCount() {
        try {
            assertEquals(metadata.getSheetCount(), 2);

        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }

    @Test
    public void testSheetNames() {
        try {
            List<String> sheetNames = metadata.getSheetNames();
            assertTrue("Expected to contain Sheet 1", sheetNames.contains("Sheet 1"));
            assertTrue("Expected to contain Sheet 2", sheetNames.contains("Sheet 2"));
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }
    
    @Test
    public void testFetchIndexByValidName() {
        try {
            assertEquals("Expected to be 1", 1 ,metadata.getSheetIndexBySheetName("Sheet 1"));
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }
       
    @Test
    public void testFetchIndexByInvalidName() {
        try {
            assertEquals("Expected to be -1", -1 ,metadata.getSheetIndexBySheetName("Sheet 23"));
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }
    
    @Test
    public void testFetchIndexByNullString() {

        thrown.expect(IllegalArgumentException.class);
        thrown.expectMessage(is("SheetName can't be null!"));

        try {
            metadata.getSheetIndexBySheetName(null);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }

    @Test
    public void testFetchIndexByEmptyName() {

        thrown.expect(IllegalArgumentException.class);
        thrown.expectMessage(is("SheetName can't be empty!"));

        try {
            metadata.getSheetIndexBySheetName("");
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }
}
