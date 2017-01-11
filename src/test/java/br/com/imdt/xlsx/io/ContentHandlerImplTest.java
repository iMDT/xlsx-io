package br.com.imdt.xlsx.io;

import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.junit.Test;
import static org.junit.Assert.*;
import org.xml.sax.SAXException;

/**
 *
 * @author imdt-klaus
 */
public class ContentHandlerImplTest {

    public ContentHandlerImplTest() {
    }

    @Test
    public void testStream() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            XlsxStreamer streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.addAll(rawValues);
                }
            });
            streamer.stream();

            assertArrayEquals("Expected to contain { ASD , DD  \n DDD , \"\" ,ASD \n \"\" , FAS , TAS }", new String[]{"ASD", "DD", "DDD", "", "ASD", "", "FAS", "TAS"}, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
