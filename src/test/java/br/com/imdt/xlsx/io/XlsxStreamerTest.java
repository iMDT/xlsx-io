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
public class XlsxStreamerTest {
    
    public XlsxStreamerTest() {
    }

    @Test
    public void testStream() {
        try {
            XlsxStreamer streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"),new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    for(String s : rawValues){
                        System.out.println("DATA:" + s);
                    }
                }
            });
            
            streamer.stream();
        } catch (InvalidFormatException ex) {
            Logger.getLogger(XlsxStreamerTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(XlsxStreamerTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxStreamerTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxStreamerTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(XlsxStreamerTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
}
