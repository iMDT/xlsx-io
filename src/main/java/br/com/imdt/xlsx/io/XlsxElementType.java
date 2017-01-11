package br.com.imdt.xlsx.io;

/**
 * The cell element type indicated by the XML.
 *
 * @author imdt-klaus
 */
public enum XlsxElementType {
    ODD_HEADER("oddHeader"),
    EVENT_HEADER("evenHeader"),
    FIRST_HEADER("firstHeader"),
    FIRST_FOOTER("firstFooter"),
    ODD_FOOTER("oddFooter"),
    EVEN_FOOTER("evenFooter"),
    ROW("row");

    private final String elementType;

    private XlsxElementType(String elementType) {
        this.elementType = elementType;
    }

    public String getElementType() {
        return elementType;
    }

    /**
     * Checks if the element type is an header or footer
     * 
     * @param elementType
     * @return The object representation of this cell type
     */
    public static boolean checkIfIsHeaderOrFooter(String elementType) {
       return ODD_HEADER.getElementType().equals(elementType) || EVENT_HEADER.getElementType().equals(elementType)
                || FIRST_HEADER.getElementType().equals(elementType) || FIRST_FOOTER.getElementType().equals(elementType)
                || ODD_FOOTER.getElementType().equals(elementType) || EVEN_FOOTER.getElementType().equals(elementType) ;
    }
    
       private boolean isTextType(String textType, boolean isIsOpen){
        if(textType == null){
            return false;
        }else if(textType.isEmpty()){
            return false;
        }
        return "v".equals(textType) || "inlineStr".equals(textType) || "t".equals(textType) && isIsOpen;
    }
}
