package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.DataCallback;
import br.com.imdt.xlsx.io.DataHandler;
import br.com.imdt.xlsx.io.XlsxDataType;
import br.com.imdt.xlsx.io.XlsxDataTypeService;
import java.util.ArrayList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class ContentHandlerImpl extends DefaultHandler {

    private final StylesTable stylesTable;

    private final ReadOnlySharedStringsTable sharedStringsTable;

    private XSSFCellStyle xSSFCellStyle;

    private final StringBuffer bufferedReadedValue;

    private XlsxDataType xlsxDataType;

    /**
     * Inform the cell reference
     *
     * @example A2 , C2 , AA.
     */
    private String cellReference;

    /**
     * Points the index of the sheet number.
     */
    private final long sheetNumber;

    /**
     * Points the number of rows readed.
     */
    private long rowNumber = 0;

    /**
     * The data callback used to inform when an sheet started or finished been
     * readed, and when a data row has itinarated
     */
    private final DataCallback dataCallback;

    /**
     * List of raw values presented in the row.
     */
    private final NotNullList rawValues;

    /**
     * List of formatted values presented in the row.
     */
    private final NotNullList formattedValues;

    /**
     * Inform when the element is a cell value.
     */
    private boolean isCellValue;
    /**
     * Inform when the element is a formula value.
     */
    private boolean isFormulaValue;
    /**
     * Inform when the element is a inline value value.
     */
    private boolean isInlineString;
    /**
     * Inform when the element is a footer or header value.
     */
    private boolean isFooterOrHeader;

    /**
     * Indicates the current column we are reading.
     */
    private long currentCol;

    /**
     * An handler to send when readed cell data is ready to be used.
     */
    private final DataHandler dataHandler;
    
    private final XlsxDataTypeService dataTypeService;

    @Override
    public void startElement(String uri, String localName, String name,
            Attributes attributes) throws SAXException {

        if (dataTypeService.isTextElement(name, isInlineString)) {
            isCellValue = true;
            bufferedReadedValue.setLength(0);
        } else if (dataTypeService.isHeaderOrFooter(name)) {
            isFooterOrHeader = true;

        } else if (XlsxDataType.INLINE_STRING_OUTER_TAG.getCellType().equals(name)) {
            isInlineString = true;
        } else if (XlsxDataType.CELL.getCellType().equals(name)) {
            xSSFCellStyle = null;
            this.xlsxDataType = XlsxDataType.NUMBER;
            cellReference = attributes.getValue("r");
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");

            xlsxDataType = dataTypeService.getByCellType(cellType);

            if (cellStyleStr != null) {
                int styleIndex = Integer.parseInt(cellStyleStr);
                xSSFCellStyle = stylesTable.getStyleAt(styleIndex);
            }
        }
    }

    @Override
    public void endElement(String uri, String localName, String name)
            throws SAXException {

        if (XlsxDataType.CELL_VALUE.getCellType().equals(name)) {

            int thisCol = (new CellReference(cellReference)).getCol();
            long missedCols = thisCol - currentCol - 1;

            for (int i = 0; i < missedCols; i++) {
                rawValues.add("");
                formattedValues.add("");
            }

            this.currentCol = thisCol;

            switch (xlsxDataType) {

                case BOOL:
                    dataHandler.handleBoolean(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                case ERROR:
                    dataHandler.handleError(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                case FORMULA:
                    dataHandler.handleFormula(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                case INLINE_STRING:
                    dataHandler.handleInlineString(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                case SSTINDEX:
                    dataHandler.handleSharedStringsTableIndex(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                case NUMBER:
                    dataHandler.handleNumber(xSSFCellStyle, bufferedReadedValue.toString());
                    break;

                default:
                    dataHandler.handleUnknow(xSSFCellStyle, bufferedReadedValue.toString());
                    break;
            }
        } else if (XlsxDataType.ROW.getCellType().equals(name)) {
            dataCallback.onRow(sheetNumber, rowNumber++, rawValues, formattedValues);
            this.currentCol = -1;
            rawValues.clear();
            formattedValues.clear();
        }
    }

    @Override
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (isCellValue) {
            bufferedReadedValue.append(ch, start, length);
        }
    }

    /**
     * Default constructor of {@link ContentHandlerImpl}
     *
     * @param sheetNumber The sheet number
     * @param dataCallback
     * @param styles Table of styles
     * @param sharedStringsTable Table of shared strings
     * @param dataHandler The handler to be delivered when data cell is ready to
     * be used
     */
    public ContentHandlerImpl(long sheetNumber,
            DataCallback dataCallback,
            StylesTable styles,
            ReadOnlySharedStringsTable sharedStringsTable,
            DataHandler dataHandler) {

        this.stylesTable = styles;
        this.sharedStringsTable = sharedStringsTable;
        this.bufferedReadedValue = new StringBuffer();
        this.sheetNumber = sheetNumber;
        this.dataCallback = dataCallback;
        this.currentCol = -1;
        this.rawValues = new NotNullList();
        this.formattedValues = new NotNullList();
        this.dataTypeService = new XlsxDataTypeServiceImpl();
        if (dataHandler == null) {
            this.dataHandler = new DefaultDataHandlerImpl(rawValues, formattedValues, sharedStringsTable);
        }else{
            this.dataHandler = dataHandler;
        }
    }
}
