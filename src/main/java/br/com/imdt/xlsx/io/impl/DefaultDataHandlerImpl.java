package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.DataHandler;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

/**
 * Default implementation of {@link DataHandler}
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class DefaultDataHandlerImpl implements DataHandler {

    private final ArrayList<String> rawValues;

    private final ArrayList<String> formattedValues;

    private final ReadOnlySharedStringsTable sharedStringsTable;

    private final DataFormatter formatter;

    public DefaultDataHandlerImpl(ArrayList<String> rawValues, ArrayList<String> formattedValues, ReadOnlySharedStringsTable sharedStringsTable) {
        this.rawValues = rawValues;
        this.formattedValues = formattedValues;
        this.sharedStringsTable = sharedStringsTable;
        this.formatter = new DataFormatter();
    }

    @Override
    public void handleBoolean(XSSFCellStyle style, String value) {
        char first = value.charAt(0);
        rawValues.add(value);
        formattedValues.add(first == '0' ? "FALSE" : "TRUE");
    }

    @Override
    public void handleError(XSSFCellStyle style, String error) {
        rawValues.add(error);
        formattedValues.add("\"ERROR:" + error + '"');

    }

    @Override
    public void handleFormula(XSSFCellStyle style, String formula) {
        rawValues.add(formula);
        formattedValues.add('"' + formula + '"');

    }

    @Override
    public void handleInlineString(XSSFCellStyle style, String inlineString) {
        XSSFRichTextString rtsi = new XSSFRichTextString(inlineString);
        rawValues.add(rtsi.toString().toUpperCase());
        formattedValues.add('"' + rtsi.toString() + '"');

    }

    @Override
    public void handleSharedStringsTableIndex(XSSFCellStyle style, String sharedStringsTableIndex) {
        try {
            int idx = Integer.parseInt(sharedStringsTableIndex);
            XSSFRichTextString rtss = new XSSFRichTextString(sharedStringsTable.getEntryAt(idx));
            rawValues.add(rtss.toString());
            formattedValues.add('"' + rtss.toString() + '"');

        } catch (NumberFormatException ex) {
            rawValues.add("ERROR");
            formattedValues.add("ERROR");
        }

    }

    @Override
    public void handleNumber(XSSFCellStyle style, String number) {
        rawValues.add(number);
        if (style != null) {
            short formatIndex = style.getDataFormat();
            String formatString = style.getDataFormatString();
            if (formatString == null) {
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }
            if (formatString != null) {
                formattedValues.add(formatter.formatRawCellContents(Double.parseDouble(number), formatIndex, formatString));
                if (formatString.contentEquals("D/M/YYYY")) {
                    rawValues.add(formatter.formatRawCellContents(Double.parseDouble(number), formatIndex, formatString));
                }
            }
        }else{
           formattedValues.add(number);
        }
    }

    @Override
    public void handleUnknow(XSSFCellStyle style, String unknow) {
        formattedValues.add("Unknow type of data: " + unknow);
        rawValues.add("Unknow type of data: " + unknow);
    }
}
