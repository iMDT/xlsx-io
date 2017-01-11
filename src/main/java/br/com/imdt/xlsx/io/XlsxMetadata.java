package br.com.imdt.xlsx.io;

import java.io.Closeable;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.xml.sax.SAXException;

/**
 * Responsible for get the Xlsx Metadata
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxMetadata implements Closeable {

    private OPCPackage pack;

    public XlsxMetadata(String fileLocation) throws InvalidFormatException, IOException {

        this.pack = OPCPackage.open(fileLocation, PackageAccess.READ);
    }

    public XlsxMetadata(File file) throws InvalidFormatException, IOException {

        this.pack = OPCPackage.open(file, PackageAccess.READ);
    }

    public XlsxMetadata(InputStream inputStream) throws InvalidFormatException, IOException {

        this.pack = OPCPackage.open(inputStream);
    }

    /**
     * Fetch the list of sheet names presented in a XLSX
     *
     * @return list of sheet names.
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    public List<String> getSheetNames() throws IOException, SAXException, OpenXML4JException {
        ArrayList<String> sheetNames = new ArrayList<String>();

        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();

        while (sheetIterator.hasNext()) {
            sheetIterator.next();
            sheetNames.add(sheetIterator.getSheetName());
        }

        return sheetNames;
    }

    /**
     * Fetch the {@link SheetIterator} by its sheetName
     *
     * @param sheetName
     * @return The sheet if found, null otherwise.
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    public SheetIterator getSheetByName(String sheetName) throws IOException, SAXException, OpenXML4JException {
        if (sheetName == null) {
            throw new IllegalArgumentException("SheetName can't be null!");
        } else if (sheetName.isEmpty()) {
            throw new IllegalArgumentException("SheetName can't be empty!");
        }
        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();

        while (sheetIterator.hasNext()) {
            sheetIterator.next();
            if (sheetIterator.getSheetName().contentEquals(sheetName)) {
                return sheetIterator;
            }
        }

        return null;
    }

    /**
     * Fetch the {@link SheetIterator} by its index, index is zero based.
     *
     * @param index
     * @return The sheet if found, null otherwise.
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    public SheetIterator getSheetByIndex(int index) throws IOException, SAXException, OpenXML4JException {
        if (index < 0) {
            throw new IllegalArgumentException("Index must be higher than -1!");
        }
        
        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();
        int currentIndex = 0;

        while (sheetIterator.hasNext()) {
            sheetIterator.next();
            if (currentIndex == index) {
                return sheetIterator;
            }
            currentIndex++;
        }

        return null;
    }

    /**
     * Fetch the index of the sheet by its name.
     *
     * @param sheetName
     * @return The sheet index, -1 otherwise
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    public int getSheetIndexBySheetName(String sheetName) throws IOException, SAXException, OpenXML4JException {
        if (sheetName == null) {
            throw new IllegalArgumentException("SheetName can't be null!");
        } else if (sheetName.isEmpty()) {
            throw new IllegalArgumentException("SheetName can't be empty!");
        }
        int index = 0;
        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();

        while (sheetIterator.hasNext()) {
            sheetIterator.next();
            if (sheetIterator.getSheetName().contentEquals(sheetName)) {
                return index;
            }
            index++;
        }

        return -1;
    }

    /**
     * Get the sheet count of the document.
     *
     *
     * @return The sheet count.
     * @throws IOException
     * @throws SAXException
     * @throws OpenXML4JException
     */
    public int getSheetCount() throws IOException, SAXException, OpenXML4JException {
        int count = 0;

        XSSFReader streamer = new XSSFReader(pack);
        SheetIterator sheetIterator = (SheetIterator) streamer.getSheetsData();

        while (sheetIterator.hasNext()) {
            sheetIterator.next();
            count++;
        }

        return count;
    }

    /**
     * Closes the file lock and release memory.
     *
     * @throws IOException
     */
    public void close() throws IOException {
        pack.close();
    }
}
