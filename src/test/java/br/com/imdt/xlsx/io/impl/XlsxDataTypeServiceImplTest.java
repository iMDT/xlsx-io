package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.XlsxDataType;
import br.com.imdt.xlsx.io.XlsxDataTypeService;
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
}
