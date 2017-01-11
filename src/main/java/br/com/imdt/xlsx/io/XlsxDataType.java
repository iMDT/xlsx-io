package br.com.imdt.xlsx.io;

/**
 * The cell type indicated by the XML.
 *
 * @author imdt-klaus
 * @see
 * <a href="https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.cell.aspx">Document
 * XLSX Format</a>
 */
public enum XlsxDataType {
    BOOL("b"),
    ERROR("e"),
    FORMULA("str"),
    INLINESTR("inlineStr"),
    SSTINDEX("s"),
    NUMBER("n"),
    SHEETDATA(""),
    ODD_HEADER("oddHeader"),
    EVENT_HEADER("evenHeader"),
    FIRST_HEADER("firstHeader"),
    FIRST_FOOTER("firstFooter"),
    ODD_FOOTER("oddFooter"),
    EVEN_FOOTER("evenFooter"),
    ROW("row"),
    STYLE("s"),
    INLINE_STRING("inlineStr"),
    REFERENCE("r"),
    DATATYPE("t"),
    FORMULA_FIELD("f"),
    INLINE_STRING_OUTER_TAG("is"),
    CELL_VALUE("v"),
    CELL("c");

    private final String cellType;

    private XlsxDataType(String cellType) {
        this.cellType = cellType;
    }

    public String getCellType() {
        return cellType;
    }

    /**
     * Return the {@link XlsxDataType} of the this cell type
     *
     * @param cellType
     * @return The object representation of this cell type
     */
    public static XlsxDataType getByCellType(String cellType) {
        for(XlsxDataType type : XlsxDataType.values()){
            if(type.getCellType().equals(cellType)){
                return type;
            }
        }
        return SHEETDATA;
    }

    /**
     * Checks if the element type is an header or footer
     *
     * @param elementType
     * @return The object representation of this cell type
     */
    public static boolean isHeaderOrFooter(String elementType) {
        return ODD_HEADER.getCellType().equals(elementType) || EVENT_HEADER.getCellType().equals(elementType)
                || FIRST_HEADER.getCellType().equals(elementType) || FIRST_FOOTER.getCellType().equals(elementType)
                || ODD_FOOTER.getCellType().equals(elementType) || EVEN_FOOTER.getCellType().equals(elementType);
    }

    public static boolean isTextElement(String textType, boolean isIsOpen) {
        if (textType == null) {
            return false;
        } else if (textType.isEmpty()) {
            return false;
        }
        return "v".equals(textType) || "inlineStr".equals(textType) || "t".equals(textType) && isIsOpen;
    }
}
