package br.com.imdt.xlsx.io;

import java.util.regex.Pattern;

/**
 * Util class used in {@link XlsxMetadata} and {@link XlsxStreamer}
 * 
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public abstract class XlsxUtils {

    /**
     * Converts an Excel column name like "C" to a zero-based index.
     *
     * @param name
     * @return Index corresponding to the specified name
     */
    public static int nameToColumn(String name) {
        int column = -1;
        for (int i = 0; i < name.length(); ++i) {
            int c = name.charAt(i);
            column = (column + 1) * 26 + c - 'A';
        }
        return column;
    }
    
    /**
     * Return the row number based on the first column.
     * @param columnName
     * @return The number of the row.
     */
    public static long getRowNumber(String columnName){
       if(!isFirstColumn(columnName)){
           throw new IllegalArgumentException("The column '" + columnName + "' is not the first column");
       }
       
       return Long.valueOf(columnName.substring(1));
    }
    
    /**
     * Return true if the column name matches the first column
     * 
     * @param columnName
     * @return true if the column name matches the first column
     */
    public static boolean isFirstColumn(String columnName){
        return Pattern.matches("^[A-Z]\\d{1,}",columnName);
    }
}
