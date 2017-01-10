package br.com.imdt.xlsx.io;

import java.io.IOException;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.Test;
import static org.junit.Assert.*;
import org.xml.sax.SAXException;

/**
 *
 * @author imdt-klaus
 */
public class XlsxMetadataTest {
    
    public XlsxMetadataTest() {
    }

    @Test
    public void testSheetCount() {
        try {
            XlsxMetadata metadata = new XlsxMetadata(ClassLoader.getSystemResourceAsStream("TestFile.xlsx"));
            try {
                assertEquals(metadata.getSheetCount(), 2);
                metadata.close();
            }   catch (SAXException ex) {
                Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
                metadata.close();
                fail("Fail to process file");
            } catch (OpenXML4JException ex) {
                Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
                metadata.close();
                fail("OpenXML fail");
            }
        } catch (InvalidFormatException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    @Test
    public void testSheetNames() {
        try {
            XlsxMetadata metadata = new XlsxMetadata(ClassLoader.getSystemResourceAsStream("TestFile.xlsx"));
            try {
                List<String> sheetNames = metadata.getSheetNames();
                assertTrue("Expected to contain Sheet 1",sheetNames.contains("Sheet 1"));
                assertTrue("Expected to contain Sheet 2",sheetNames.contains("Sheet 2"));
            } catch (SAXException ex) {
                Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
                metadata.close();
                fail("Fail to process file");
            } catch (OpenXML4JException ex) {
                Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
                metadata.close();
                fail("OpenXML fail");
            }
        } catch (InvalidFormatException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
