package br.com.imdt.xlsx.io;

/**
 * The cell type indicated by the XML.
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 * @see
 * <a href="https://msdn.microsoft.com/en-us/library/office/documentformat.openxml.spreadsheet.cell.aspx">Document
 * XLSX Format</a>
 */
public enum XlsxDataType {
    BOOL("b"),
    ERROR("e"),
    FORMULA("str"),
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
}
