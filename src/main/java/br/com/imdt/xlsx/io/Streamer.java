package br.com.imdt.xlsx.io;

import br.com.imdt.xlsx.io.exception.SheetNotFoundException;
import java.io.IOException;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public interface Streamer {

    /**
     * Stream all sheets presented in the document.
     *
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     */
    public void stream() throws IOException, SAXException, OpenXML4JException, ParserConfigurationException;
    
    
    /**
     * Stream the sheet by its sheetName
     *
     * @param sheetName
     *
     * @throws IllegalArgumentException If sheet name is null or empty
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SheetNotFoundException If the sheet with the given name couldn't be found.
     */
    public void streamSheetByName(String sheetName) throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, SheetNotFoundException;
    
    
    /**
     * Fetch the {@link SheetIterator} by its index, index is zero based.
     *
     * @param index
     * @throws IllegalArgumentException If sheet index is less than zero.
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     * @throws javax.xml.parsers.ParserConfigurationException
     * @throws SheetNotFoundException If the index couldn't be found.
     */
    public void streamSheetByIndex(int index) throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, SheetNotFoundException;
    
}
