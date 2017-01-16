package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.XlsxDataType;
import br.com.imdt.xlsx.io.XlsxDataTypeService;
import java.util.ArrayList;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Before;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxDataTypeServiceImplTest {

    private XlsxDataTypeService dataTypeService;

    public XlsxDataTypeServiceImplTest() {
    }

    @Before
    public void setUp() {
        dataTypeService = new XlsxDataTypeServiceImpl();
    }

    @Test
    public void testIsTextElement() {
        assertTrue(dataTypeService.isTextElement("v", true));
        assertTrue(dataTypeService.isTextElement("inlineStr", true));
        assertTrue(dataTypeService.isTextElement("t", true));

        assertTrue(dataTypeService.isTextElement("v", false));
        assertTrue(dataTypeService.isTextElement("inlineStr", false));
        assertTrue(dataTypeService.isTextElement("t", true));

        assertFalse(dataTypeService.isTextElement("", true));
        assertFalse(dataTypeService.isTextElement("", false));
        assertFalse(dataTypeService.isTextElement("t", false));

        assertFalse(dataTypeService.isTextElement(null, true));
        assertFalse(dataTypeService.isTextElement(null, false));

    }

    @Test
    public void testIsFooterOrHeader() {
        assertTrue(dataTypeService.isHeaderOrFooter("oddheader"));
        assertTrue(dataTypeService.isHeaderOrFooter("evenheader"));
        assertTrue(dataTypeService.isHeaderOrFooter("firstheader"));
        assertTrue(dataTypeService.isHeaderOrFooter("firstfooter"));
        assertTrue(dataTypeService.isHeaderOrFooter("oddfooter"));
        assertTrue(dataTypeService.isHeaderOrFooter("evenfooter"));

        assertTrue(dataTypeService.isHeaderOrFooter("oDdHEader"));
        assertTrue(dataTypeService.isHeaderOrFooter("evenheader"));
        assertTrue(dataTypeService.isHeaderOrFooter("firSTHeader"));
        assertTrue(dataTypeService.isHeaderOrFooter("FIrstFooter"));
        assertTrue(dataTypeService.isHeaderOrFooter("ODDFooter"));
        assertTrue(dataTypeService.isHeaderOrFooter("evenFOOTER"));
    }

    @Test
    public void testFactory() {
        for (XlsxDataType type : XlsxDataType.values()) {
            assertEquals(type, dataTypeService.getByCellType(type.getCellType()));
        }
    }

    @Test
    public void testIsEmptyRow() {
        ArrayList<String> notEmptyList = new ArrayList<String>(10);

        notEmptyList.add(null);
        notEmptyList.add(" ");
        notEmptyList.add("123");
        notEmptyList.add("das");
        notEmptyList.add("dasdas");
        notEmptyList.add("fsa4d5f");
        notEmptyList.add("");
        notEmptyList.add("fsdf45s");
        notEmptyList.add("dsf45das");
        notEmptyList.add("5sa45sd");
        
        ArrayList<String> emptyList = new ArrayList<String>(10);

        assertFalse(dataTypeService.isRowEmpty(notEmptyList));
        notEmptyList.clear();
        
        assertTrue(dataTypeService.isRowEmpty(emptyList));
        
        emptyList.add(null);
        emptyList.add(null);
        emptyList.add("");
        emptyList.add(" ");
        emptyList.add("    ");
        assertTrue(dataTypeService.isRowEmpty(emptyList));
    }
}
