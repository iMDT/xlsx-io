package br.com.imdt.xlsx.io;

import java.util.ArrayList;

/**
 * This interface is responsible for deliver the data when streaming over the xlsx
 * 
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public interface DataCallback {
    
    /**
     * Called when the documented has started being read
     */
    public void onBegin();
    
    /**
     * Called when an row has finished been readed.
     * @param sheetNumber
     * @param rowNum
     * @param rawValues
     * @param formattedValues 
     */
    public void onRow(Long sheetNumber,Long rowNum,ArrayList<String> rawValues, ArrayList<String> formattedValues);
    
    /**
     * Called when the documented has ended being read
     */
    public void onEnd();
    
}
