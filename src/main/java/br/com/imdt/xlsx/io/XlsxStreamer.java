package br.com.imdt.xlsx.io;

import br.com.imdt.xlsx.io.exception.SheetNotFoundException;
import br.com.imdt.xlsx.io.impl.ContentHandlerImpl;
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
 * This class is used to stream all sheets in a document, or a specific sheet
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxStreamer implements Streamer, Closeable {

    private DataCallback dataCallback;
    private OPCPackage pack;
    
    public boolean ignoreEmptyRow = true;

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

    public boolean isIgnoringEmptyRows() {
        return ignoreEmptyRow;
    }

    public void setIgnoreEmptyRow(boolean ignoreEmptyRow) {
        this.ignoreEmptyRow = ignoreEmptyRow;
    }

    /**
     * Parses and shows the content of one sheet using the specified styles and
     * shared-strings tables.
     *
     * @param styles
     * @param strings
     * @param sheetInputStream
     * @param sheetNumber
     * @throws java.io.IOException
     * @throws javax.xml.parsers.ParserConfigurationException
     * @throws org.xml.sax.SAXException
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
        ContentHandler handler = new ContentHandlerImpl(sheetNumber, dataCallback, styles, strings, null,ignoreEmptyRow);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

    /**
     * Call
     * {@link #processSheet(org.apache.poi.xssf.model.StylesTable, org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable, java.io.InputStream, int)}
     * with the given stream.
     *
     * @param streamer
     * @param stream
     * @param sheetNumber
     * @throws IOException
     * @throws ParserConfigurationException
     * @throws InvalidFormatException
     * @throws SAXException
     */
    private void streamSheet(XSSFReader streamer, InputStream stream, int sheetNumber) throws IOException, ParserConfigurationException, InvalidFormatException, SAXException {
        ReadOnlySharedStringsTable sharedStringsTable = new ReadOnlySharedStringsTable(pack);
        StylesTable styles = streamer.getStylesTable();
        dataCallback.onSheetBegin();
        processSheet(styles, sharedStringsTable, stream, sheetNumber);
        dataCallback.onSheetEnd();
        stream.close();
    }

    @Override
    public void stream() throws IOException, SAXException, OpenXML4JException, ParserConfigurationException {
        
        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();
        dataCallback.onBegin();
        int sheetNumber = 1;
        while (sheetIterator.hasNext()) {
            streamSheet(streamer, sheetIterator.next(), sheetNumber);
            sheetNumber++;
        }
        dataCallback.onEnd();
    }

    @Override
    public void streamSheetByName(String sheetName) throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, SheetNotFoundException {
        if (sheetName == null) {
            throw new IllegalArgumentException("SheetName can't be null!");
        } else if (sheetName.isEmpty()) {
            throw new IllegalArgumentException("SheetName can't be empty!");
        }
        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();

        int sheetNumber = 0;
        InputStream stream;

        while (sheetIterator.hasNext()) {
            stream = sheetIterator.next();
            if (sheetIterator.getSheetName().contentEquals(sheetName)) {
                dataCallback.onBegin();
                streamSheet(streamer, stream, sheetNumber);
                dataCallback.onEnd();
                return;
            }
            sheetNumber++;
        }
        throw new SheetNotFoundException(sheetName);
    }

    @Override
    public void streamSheetByIndex(int index) throws IOException, SAXException, OpenXML4JException, ParserConfigurationException, SheetNotFoundException {
        if (index < 0) {
            throw new IllegalArgumentException("Index must be higher than -1!");
        }

        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();
        int sheetNumber = 0;
        InputStream stream;

        while (sheetIterator.hasNext()) {
            stream = sheetIterator.next();

            if (sheetNumber == index) {
                dataCallback.onBegin();
                streamSheet(streamer, stream, sheetNumber);
                dataCallback.onEnd();
                return;
            }
            sheetNumber++;
        }

        throw new SheetNotFoundException(index);
    }

    @Override
    public void close() throws IOException {
        pack.close();
    }
}
