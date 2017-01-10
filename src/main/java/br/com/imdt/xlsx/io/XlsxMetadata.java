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
 * @author imdt-klaus
 */
public class XlsxMetadata implements Closeable{
    
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

    public void close() throws IOException {
        pack.close();
    }
}
