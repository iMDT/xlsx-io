package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.XlsxDataType;
import static br.com.imdt.xlsx.io.XlsxDataType.EVENT_HEADER;
import static br.com.imdt.xlsx.io.XlsxDataType.EVEN_FOOTER;
import static br.com.imdt.xlsx.io.XlsxDataType.FIRST_FOOTER;
import static br.com.imdt.xlsx.io.XlsxDataType.FIRST_HEADER;
import static br.com.imdt.xlsx.io.XlsxDataType.ODD_FOOTER;
import static br.com.imdt.xlsx.io.XlsxDataType.ODD_HEADER;
import static br.com.imdt.xlsx.io.XlsxDataType.NUMBER;
import br.com.imdt.xlsx.io.XlsxDataTypeService;

/**
 * 
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxDataTypeServiceImpl implements XlsxDataTypeService {
    
    @Override
    public XlsxDataType getByCellType(String cellType) {
        for(XlsxDataType type : XlsxDataType.values()){
            if(type.getCellType().equalsIgnoreCase(cellType)){
                return type;
            }
        }
        return NUMBER;
    }

    @Override
    public boolean isHeaderOrFooter(String elementType) {
        return ODD_HEADER.getCellType().equalsIgnoreCase(elementType) || EVENT_HEADER.getCellType().equalsIgnoreCase(elementType)
                || FIRST_HEADER.getCellType().equalsIgnoreCase(elementType) || FIRST_FOOTER.getCellType().equalsIgnoreCase(elementType)
                || ODD_FOOTER.getCellType().equalsIgnoreCase(elementType) || EVEN_FOOTER.getCellType().equalsIgnoreCase(elementType);
    }

    @Override
    public boolean isTextElement(String textType, boolean isIsOpen) {
        if (textType == null) {
            return false;
        } else if (textType.isEmpty()) {
            return false;
        }
        return "v".equals(textType) || "inlineStr".equals(textType) || "t".equals(textType) && isIsOpen;
    }
    
}
