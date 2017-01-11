package br.com.imdt.xlsx.io;

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
}
