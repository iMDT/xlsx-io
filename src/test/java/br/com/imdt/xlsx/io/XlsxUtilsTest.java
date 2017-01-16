package br.com.imdt.xlsx.io;

import static org.hamcrest.CoreMatchers.is;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Rule;
import org.junit.rules.ExpectedException;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class XlsxUtilsTest {

    public XlsxUtilsTest() {
    }

    @Rule
    public ExpectedException thrown = ExpectedException.none();

    @Test
    public void shouldReturn0() {
        int result = XlsxUtils.nameToColumn("A");
        assertEquals("should be 0, returned " + result, 0, result);
    }

    @Test
    public void shouldReturn26() {
        int result = XlsxUtils.nameToColumn("AA");
        assertEquals("should be 26, returned " + result, 26, result);
    }

    @Test
    public void shouldReturn28() {
        int result = XlsxUtils.nameToColumn("AC");
        assertEquals("should be 28, returned " + result, 28, result);
    }

    @Test
    public void testNotFirstColumn() {
        assertFalse(XlsxUtils.isFirstColumn("CC1"));
        assertFalse(XlsxUtils.isFirstColumn("AA1"));
    }

    @Test
    public void testIsFirstColumn() {
        assertTrue(XlsxUtils.isFirstColumn("A1"));
        assertTrue(XlsxUtils.isFirstColumn("A14654"));
        assertTrue(XlsxUtils.isFirstColumn("C1"));
        assertTrue(XlsxUtils.isFirstColumn("C14654"));
        assertTrue(XlsxUtils.isFirstColumn("Z1"));
        assertTrue(XlsxUtils.isFirstColumn("D14654"));
        assertTrue(XlsxUtils.isFirstColumn("G1"));
        assertTrue(XlsxUtils.isFirstColumn("H14654"));
    }
    
    @Test
    public void testIncorrectColumnNumber(){

        thrown.expect(IllegalArgumentException.class);
        
        XlsxUtils.getRowNumber("CC1");
    }
    
    @Test
    public void testColumnNumber(){
        assertEquals(22,XlsxUtils.getRowNumber("A22"));
        assertEquals(1,XlsxUtils.getRowNumber("A1"));
        assertEquals(0,XlsxUtils.getRowNumber("A0"));
        assertEquals(123456,XlsxUtils.getRowNumber("A123456"));
        
        assertEquals(22,XlsxUtils.getRowNumber("D22"));
        assertEquals(1,XlsxUtils.getRowNumber("D1"));
        assertEquals(0,XlsxUtils.getRowNumber("T0"));
        assertEquals(123456,XlsxUtils.getRowNumber("G123456"));
        
        assertEquals(22,XlsxUtils.getRowNumber("Z22"));
        assertEquals(1,XlsxUtils.getRowNumber("L1"));
        assertEquals(0,XlsxUtils.getRowNumber("R0"));
        assertEquals(123456,XlsxUtils.getRowNumber("O123456"));
    }
}
