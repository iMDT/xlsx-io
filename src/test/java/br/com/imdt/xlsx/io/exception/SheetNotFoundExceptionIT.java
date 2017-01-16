/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package br.com.imdt.xlsx.io.exception;

import static org.hamcrest.CoreMatchers.is;
import org.junit.Test;
import org.junit.Rule;
import org.junit.rules.ExpectedException;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class SheetNotFoundExceptionIT {
    
    @Rule
    public ExpectedException thrown = ExpectedException.none();
    
    @Test
    public void testCorrectMessageIsThrowWithNumber() {
        thrown.expect(SheetNotFoundException.class);
        thrown.expectMessage(is("The sheet number (20) couldn't be found"));
        
        throw new SheetNotFoundException(20);
    }
    
    @Test
    public void testCorrectMessageIsThrowWithName() {
        thrown.expect(SheetNotFoundException.class);
        thrown.expectMessage(is("The sheet with name 'Sheet 23' couldn't be found"));
        
        throw new SheetNotFoundException("Sheet Name");
    }
    
}
