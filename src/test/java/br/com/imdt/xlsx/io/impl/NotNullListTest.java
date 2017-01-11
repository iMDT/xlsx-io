package br.com.imdt.xlsx.io.impl;

import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Before;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class NotNullListTest {
    
    private NotNullList notNullList;
    
    public NotNullListTest() {
    }
    
    @Before
    public void prepareList(){
        notNullList = new NotNullList();
    }
    
    @Test
    public void testNotNullOnEmptyIndex() {
        
        notNullList.add("asd");
        
        assertTrue(notNullList.get(0).contentEquals("asd"));
        assertTrue(notNullList.get(2).contentEquals(""));
    }
    
}
