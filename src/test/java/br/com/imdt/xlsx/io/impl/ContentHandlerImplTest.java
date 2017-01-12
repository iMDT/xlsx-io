package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.DefaultCallback;
import br.com.imdt.xlsx.io.XlsxStreamer;
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
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class ContentHandlerImplTest {

    public ContentHandlerImplTest() {
    }

    @Test
    public void testWithRawValues() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            XlsxStreamer streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.addAll(rawValues);

                    for (String s : rawValues) {
                        System.out.println(s);
                    }
                }
            });
            streamer.stream();

            assertArrayEquals("Expected to contain { ASD , DD  \n DDD , \"\" ,ASD \n \"\" , FAS , TAS \n 123 , DAS , AAA \n 42715, 23.05, 50}", new String[]{"ASD", "DD", "DDD", "", "ASD", "", "FAS", "TAS", "123", "DAS", "AAA", "42715", "23.05", "50"}, allRowValues.toArray());
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

    @Test
    public void testWithFormattedValues() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            XlsxStreamer streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.addAll(formattedValues);

                    for (String s : formattedValues) {
                        System.out.println(s);
                    }
                }
            });
            streamer.stream();

            assertArrayEquals("Expected to contain { \"ASD\" , \"DD\"  \n \"DDD\" , \"\" ,\"ASD\" \n \"\" , \"FAS\" , \"TAS\" \n \"123\" , \"DAS\" , \"AAA\" \n 11/12/16, 23,05 , R$ 50}", new String[]{"\"ASD\"", "\"DD\"", "\"DDD\"", "", "\"ASD\"", "", "\"FAS\"", "\"TAS\"", "123", "\"DAS\"", "\"AAA\"", "11/12/16", "23.05", "R$ 50"}, allRowValues.toArray());
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
