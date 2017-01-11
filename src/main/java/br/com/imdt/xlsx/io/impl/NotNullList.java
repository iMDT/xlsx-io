package br.com.imdt.xlsx.io.impl;

import java.util.ArrayList;

/**
 * Prevent the ArrayList from sending null values or throwing {@link ArrayIndexOutOfBoundsException}
 * 
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class NotNullList extends ArrayList<String> {

    @Override
    public String get(int index) {
        if (index >= this.size()) {
            return "";
        } else {
            return super.get(index);
        }
    }
}
