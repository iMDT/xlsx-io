package br.com.imdt.xlsx.io;

import org.apache.poi.xssf.usermodel.XSSFCellStyle;

/**
 * Deliver the data when its detected on {@link ContentHandlerImpl}
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public interface DataHandler {

    public void handleBoolean(XSSFCellStyle style ,String value);

    public void handleError(XSSFCellStyle style ,String error);

    public void handleFormula(XSSFCellStyle style ,String formula);

    public void handleInlineString(XSSFCellStyle style ,String inlineString);

    public void handleSharedStringsTableIndex(XSSFCellStyle style ,String sharedStringsTableIndex);

    public void handleNumber(XSSFCellStyle style , String number);

    public void handleUnknow(XSSFCellStyle style ,String unknow);

}
