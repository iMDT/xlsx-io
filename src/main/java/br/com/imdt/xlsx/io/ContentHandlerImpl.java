package br.com.imdt.xlsx.io;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 * @author imdt-klaus
 */
public class ContentHandlerImpl extends DefaultHandler {

    /**
     * Table with styles
     */
    private final StylesTable stylesTable;

    /**
     * Table with unique strings
     */
    private final ReadOnlySharedStringsTable sharedStringsTable;

    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;

    // Gathers characters as they are seen.
    private final StringBuffer value;

    private xssfDataType nextDataType;

    private String cellReference;

    private final long sheetNumber;
    private long rowNumber = 0;

    private final DataCallback dataCallback;

    private ArrayList<String> rawValues = new ArrayList<String>(30);
    private ArrayList<String> formattedValues = new ArrayList<String>(30);
   // Set when V start element is seen
   private boolean vIsOpen;
   // Set when F start element is seen
   private boolean fIsOpen;
   // Set when an Inline String "is" is seen
   private boolean isIsOpen;
   // Set when a header/footer element is seen
   private boolean hfIsOpen;

    /**
     * The type of the data value is indicated by an attribute on the cell. The
     * value is usually in a "v" element within the cell.
     */
    enum xssfDataType {
        BOOL,
        ERROR,
        FORMULA,
        INLINESTR,
        SSTINDEX,
        NUMBER,
        SHEETDATA
    }

    /**
     * Accepts objects needed while parsing.
     *
     * @param sheetNumber
     * @param dataCallback
     * @param styles Table of styles
     * @param strings Table of shared strings
     */
    public ContentHandlerImpl(long sheetNumber,
            DataCallback dataCallback,
            StylesTable styles,
            ReadOnlySharedStringsTable strings) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.value = new StringBuffer();
        this.formatter = new DataFormatter();
        this.sheetNumber = sheetNumber;
        this.dataCallback = dataCallback;
    }


    /*
    * (non-Javadoc)
    * @see org.xml.sax.helpers.DefaultHandler#startElement(java.lang.String, java.lang.String, java.lang.String, org.xml.sax.Attributes)
     */
    @Override
    public void startElement(String uri, String localName, String name,
            Attributes attributes) throws SAXException {

       if (isTextTag(name)) {
           vIsOpen = true;
           // Clear contents cache
           value.setLength(0);
       } else if ("is".equals(name)) {
          // Inline string outer tag
          isIsOpen = true;
       } else if("oddHeader".equals(name) || "evenHeader".equals(name) ||
             "firstHeader".equals(name) || "firstFooter".equals(name) ||
             "oddFooter".equals(name) || "evenFooter".equals(name)) {
          hfIsOpen = true;
       }
       else if("row".equals(name)) {
           int rowNum = Integer.parseInt(attributes.getValue("r")) - 1;
       }
       // c => cell
       else if ("c".equals(name)) {
           // Set up defaults.
           this.nextDataType = xssfDataType.NUMBER;
           this.formatIndex = -1;
           this.formatString = null;
           cellReference = attributes.getValue("r");
           String cellType = attributes.getValue("t");
           String cellStyleStr = attributes.getValue("s");
           System.out.println("CellType: " + cellType);
           if ("b".equals(cellType))
               nextDataType = xssfDataType.BOOL;
           else if ("e".equals(cellType))
               nextDataType = xssfDataType.ERROR;
           else if ("inlineStr".equals(cellType))
               nextDataType = xssfDataType.INLINESTR;
           else if ("s".equals(cellType))
               nextDataType = xssfDataType.SSTINDEX;
           else if ("str".equals(cellType))
               nextDataType = xssfDataType.FORMULA;
           else if (cellStyleStr != null) {
              // Number, but almost certainly with a special style or format
               int styleIndex = Integer.parseInt(cellStyleStr);
               XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
               this.formatIndex = style.getDataFormat();
               this.formatString = style.getDataFormatString();
               if (this.formatString == null)
                   this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
           }
       }

    }

    /*
    * (non-Javadoc)
    * @see org.xml.sax.helpers.DefaultHandler#endElement(java.lang.String, java.lang.String, java.lang.String)
     */
    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {

        String thisStr = null;
        System.out.println("Name: " + name);
        System.out.println("Value: " + value);

        // v => contents of a cell
        if ("v".equals(name)) {
            // Process the value contents as required.
            // Do now, as characters() may be called more than once
            switch (nextDataType) {

                case BOOL:
                    char first = value.charAt(0);
                    rawValues.add( value.toString());
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    formattedValues.add( thisStr);

                    break;

                case ERROR:
                    rawValues.add( value.toString());
                    thisStr = "\"ERROR:" + value.toString() + '"';
                    formattedValues.add(  thisStr);

                    break;

                case FORMULA:
                    // A formula could result in a string value,
                    // so always add double-quote characters.
                    rawValues.add( value.toString());
                    thisStr = '"' + value.toString() + '"';
                    formattedValues.add( thisStr);

                    break;

                case INLINESTR:
                    // TODO: have seen an example of this, so it's untested.
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    rawValues.add( rtsi.toString().toUpperCase());
                    thisStr = '"' + rtsi.toString() + '"';
                    formattedValues.add( thisStr);

                    break;

                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                        rawValues.add( rtss.toString());
                        thisStr = '"' + rtss.toString() + '"';
                        formattedValues.add( thisStr.toUpperCase());

                    } catch (NumberFormatException ex) {
                        rawValues.add("ERROR");
                        formattedValues.add("ERROR");
                    }
                    break;

                case NUMBER:
                    String n = value.toString();
                    rawValues.add(  n);
                    if (this.formatString != null) {
                        formattedValues.add(formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString));
                        thisStr = n;
                        if (formatString.contentEquals("D/M/YYYY")) {
                            rawValues.add(  formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString));
                        }
                    }
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    formattedValues.add( thisStr.toUpperCase());
                    rawValues.add(  value.toString());

                    break;
            }
            if(cellReference == null){
                rawValues.add("");
                formattedValues.add("");
            }
        } else if ("row".equals(name)) {
            dataCallback.onRow(sheetNumber, rowNumber++, rawValues, formattedValues);
            rawValues.clear();
            formattedValues.clear();
        }
    }
    
    
   private boolean isTextTag(String name) {
      if("v".equals(name)) {
         // Easy, normal v text tag
         return true;
      }
      if("inlineStr".equals(name)) {
         // Easy inline string
         return true;
      }
      if("t".equals(name) && isIsOpen) {
         // Inline string <is><t>...</t></is> pair
         return true;
      }
      // It isn't a text tag
      return false;
   }
}
