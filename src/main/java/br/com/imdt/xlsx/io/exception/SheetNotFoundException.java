package br.com.imdt.xlsx.io.exception;

/**
 * Exception throw when an sheet could not be found to proccess.
 * 
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class SheetNotFoundException extends RuntimeException{
    
    public SheetNotFoundException(String sheetName) {
        super("The sheet with name '" + sheetName + "' couldn't be found");
    }
    
    public SheetNotFoundException(int sheetNumber){
        super("The sheet number (" + sheetNumber + ") couldn't be found");
    }
}
