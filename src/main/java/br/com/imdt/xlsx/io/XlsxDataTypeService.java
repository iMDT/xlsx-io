package br.com.imdt.xlsx.io;

import br.com.imdt.xlsx.io.impl.ContentHandlerImpl;
import java.util.List;

/**
 * Service interface used in {@link ContentHandlerImpl}
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public interface XlsxDataTypeService{

    /**
     * Return the {@link XlsxDataType} of the this cell type
     *
     * @param cellType
     * @return The object representation of this cell type
     */
    public XlsxDataType getByCellType(String cellType);

    /**
     * Checks if the element type is an header or footer
     *
     * @param elementType
     * @return True if the element is footer or header, false otherwise.
     */
    public boolean isHeaderOrFooter(String elementType);

    /**
     * Checks if the element type is an text element.
     *
     * @param textType
     * @param isIsOpen
     * @return True if the element is text like, false otherwise.
     */
    public boolean isTextElement(String textType, boolean isIsOpen);

    /**
     * Check if the row is empty.
     *
     * @param list
     * @return True if the collection is empty
     */
    public boolean isRowEmpty(List<String> list);

}
