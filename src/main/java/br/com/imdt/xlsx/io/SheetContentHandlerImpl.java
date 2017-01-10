package br.com.imdt.xlsx.io;

import java.util.ArrayList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;

/**
 *
 * @author imdt-klaus
 */
public class SheetContentHandlerImpl implements XSSFSheetXMLHandler.SheetContentsHandler {

    private boolean firstCellOfRow = false;
    private int currentRow = -1;
    private int currentCol = -1;
    private ArrayList<String> values = new ArrayList<String>();

    @Override
    public void startRow(int rowNum) {
        // Prepare for this row
        firstCellOfRow = true;
        currentRow = rowNum;
        currentCol = -1;
    }

    @Override
    public void endRow() {

    }

    @Override
    public void cell(String cellReference, String formattedValue) {
        if (firstCellOfRow) {
            firstCellOfRow = false;
        } else {
            values.add("");
        }

        // gracefully handle missing CellRef here in a similar way as XSSFCell does
        if (cellReference == null) {

        }

        // Did we miss any cells?
        int thisCol = (new CellReference(cellReference)).getCol();
        int missedCols = thisCol - currentCol - 1;
        for (int i = 0; i < missedCols; i++) {
            values.add("");
        }
        currentCol = thisCol;

        values.add(formattedValue);
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        // Skip, no headers or footers in CSV
    }
}
