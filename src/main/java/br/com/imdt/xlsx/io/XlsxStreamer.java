package br.com.imdt.xlsx.io;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 *
 * @author imdt-klaus
 */
public class XlsxStreamer implements Closeable{

    private DataCallback dataCallback;
    private OPCPackage pack;

    private XlsxStreamer(DataCallback dataCallback) {
        this.dataCallback = dataCallback;
    }

    public XlsxStreamer(String fileLocation, DataCallback dataCallback) throws InvalidFormatException, IOException {
        this(dataCallback);
        this.pack = OPCPackage.open(fileLocation, PackageAccess.READ);
    }

    public XlsxStreamer(File file, DataCallback dataCallback) throws InvalidFormatException, IOException {
        this(dataCallback);
        this.pack = OPCPackage.open(file, PackageAccess.READ);
    }

    public XlsxStreamer(InputStream inputStream, DataCallback dataCallback) throws InvalidFormatException, IOException {
        this(dataCallback);
        this.pack = OPCPackage.open(inputStream);
    }

    public void stream() throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(pack);
        XSSFReader streamer = new XSSFReader(pack);
        StylesTable styles = streamer.getStylesTable();
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();
        dataCallback.onBegin();
        int sheetNumber = 1;
        while (sheetIterator.hasNext()) {
            InputStream stream = sheetIterator.next();
            processSheet(styles, sharedStringsTable, stream,sheetNumber);
            stream.close();
            sheetNumber++;
        }
        dataCallback.onEnd();
    }
    
    /**
     * Parses and shows the content of one sheet using the specified styles and
     * shared-strings tables.
     *
     * @param styles
     * @param strings
     * @param sheetInputStream
     */
    public void processSheet(
            StylesTable styles,
            ReadOnlySharedStringsTable strings,
            InputStream sheetInputStream, int sheetNumber)
            throws IOException, ParserConfigurationException, SAXException {

        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        ContentHandler handler = new ContentHandlerImpl(sheetNumber,dataCallback,styles, strings);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

    public void close() throws IOException {
        pack.close();
    }
}
