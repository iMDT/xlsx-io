package br.com.imdt.xlsx.io;

import java.util.ArrayList;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
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

    private XlsxDataType xlsxDataType;

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

    private long currentCol;

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
        this.currentCol = -1;
    }

    @Override
    public void startElement(String uri, String localName, String name,
            Attributes attributes) throws SAXException {

        if (isTextType(name)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        } else if ("is".equals(name)) {
            // Inline string outer tag
            isIsOpen = true;
        } else if ("oddHeader".equals(name) || "evenHeader".equals(name)
                || "firstHeader".equals(name) || "firstFooter".equals(name)
                || "oddFooter".equals(name) || "evenFooter".equals(name)) {
            hfIsOpen = true;
        } else if ("row".equals(name)) {

        } // c => cell
        else if ("c".equals(name)) {
            // Set up defaults.
            this.xlsxDataType = XlsxDataType.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            cellReference = attributes.getValue("r");
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");

            xlsxDataType = XlsxDataType.getByCellType(cellType);

            if (cellStyleStr != null) {
                // Number, but almost certainly with a special style or format
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null) {
                    this.formatString = BuiltinFormats.getBuiltinFormat(this.formatIndex);
                }
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {

        String thisStr = null;
        // v => contents of a cell
        if ("v".equals(name)) {

            System.out.println("currentCol: " + currentCol);

            int thisCol = (new CellReference(cellReference)).getCol();
            long missedCols = thisCol - currentCol - 1;

            for (int i = 0; i < missedCols; i++) {
                rawValues.add("");
                formattedValues.add("");
            }

            this.currentCol = thisCol;

            switch (xlsxDataType) {

                case BOOL:
                    char first = value.charAt(0);
                    rawValues.add(value.toString());
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    formattedValues.add(thisStr);

                    break;

                case ERROR:
                    rawValues.add(value.toString());
                    thisStr = "\"ERROR:" + value.toString() + '"';
                    formattedValues.add(thisStr);

                    break;

                case FORMULA:
                    // A formula could result in a string value,
                    // so always add double-quote characters.
                    rawValues.add(value.toString());
                    thisStr = '"' + value.toString() + '"';
                    formattedValues.add(thisStr);

                    break;

                case INLINESTR:
                    // TODO: have seen an example of this, so it's untested.
                    XSSFRichTextString rtsi = new XSSFRichTextString(value.toString());
                    rawValues.add(rtsi.toString().toUpperCase());
                    thisStr = '"' + rtsi.toString() + '"';
                    formattedValues.add(thisStr);

                    break;

                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
                        rawValues.add(rtss.toString());
                        thisStr = '"' + rtss.toString() + '"';
                        formattedValues.add(thisStr.toUpperCase());

                    } catch (NumberFormatException ex) {
                        rawValues.add("ERROR");
                        formattedValues.add("ERROR");
                    }
                    break;

                case NUMBER:
                    String n = value.toString();
                    rawValues.add(n);
                    if (this.formatString != null) {
                        formattedValues.add(formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString));
                        thisStr = n;
                        if (formatString.contentEquals("D/M/YYYY")) {
                            rawValues.add(formatter.formatRawCellContents(Double.parseDouble(n), this.formatIndex, this.formatString));
                        }
                    }
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + xlsxDataType + ")";
                    formattedValues.add(thisStr.toUpperCase());
                    rawValues.add(value.toString());

                    break;
            }
        } else if ("row".equals(name)) {
            dataCallback.onRow(sheetNumber, rowNumber++, rawValues, formattedValues);
            this.currentCol = -1;
            rawValues.clear();
            formattedValues.clear();
        }
    }

    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen) {
            value.append(ch, start, length);
        }
    }

    private boolean isTextType(String name) {
        if (name == null) {
            return false;
        } else if (name.isEmpty()) {
            return false;
        }
        return "v".equals(name) || "inlineStr".equals(name) || "t".equals(name) && isIsOpen;
    }
}
