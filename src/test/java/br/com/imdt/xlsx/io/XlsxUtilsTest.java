package br.com.imdt.xlsx.io;

import org.junit.Test;
import static org.junit.Assert.*;

/**
 *
 * @author imdt-klaus
 */
public class XlsxUtilsTest {
    
    public XlsxUtilsTest() {
    }

    @Test
    public void shouldReturn0() {
        int result = XlsxUtils.nameToColumn("A");
        assertEquals("should be 0, returned " + result , 0 ,result);
    }
    
    @Test
    public void shouldReturn26() {
        int result = XlsxUtils.nameToColumn("AA");
        assertEquals("should be 26, returned " + result , 26 ,result);
    }
    
    @Test
    public void shouldReturn28() {
        int result = XlsxUtils.nameToColumn("AC");
        assertEquals("should be 28, returned " + result , 28 ,result);
    }
}
