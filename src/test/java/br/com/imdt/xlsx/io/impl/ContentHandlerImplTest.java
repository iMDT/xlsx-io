package br.com.imdt.xlsx.io.impl;

import br.com.imdt.xlsx.io.DefaultCallback;
import br.com.imdt.xlsx.io.Streamer;
import br.com.imdt.xlsx.io.XlsxMetadataTest;
import br.com.imdt.xlsx.io.XlsxStreamer;
import br.com.imdt.xlsx.io.exception.SheetNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import static org.hamcrest.CoreMatchers.is;
import org.junit.Test;
import static org.junit.Assert.*;
import org.junit.Before;
import org.junit.Rule;
import org.junit.rules.ExpectedException;
import org.xml.sax.SAXException;

/**
 *
 * @author <a href="github.com/klauswk">Klaus Klein</a>
 */
public class ContentHandlerImplTest {

    @Rule
    public ExpectedException thrown = ExpectedException.none();

    private Streamer streamer;

    public ContentHandlerImplTest() {
        try {
            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {

                }
            });
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @Before
    public void setUp() {

    }

    @Test
    public void testWithRawValues() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(23);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(rawValues.get(0));
                    allRowValues.add(rawValues.get(1));
                    allRowValues.add(rawValues.get(2));
                }
            });

            ((XlsxStreamer) streamer).setIgnoreEmptyRow(false);

            streamer.stream();

            for (String s : allRowValues) {
                System.out.println(s);
            }
            final String[] expectedResult = new String[]{"ASD", "DD", "", "DDD", "", "ASD", "", "FAS", "TAS", "123", "DAS", "AAA", "42715", "23.05", "50", "d65as4", "das4d6", "dasdsa", "", "", "", "6a6dsa5hfghfg", "tertew", "gdfg"};

            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public String concatResult(String[] expectedResult) {
        StringBuilder b = new StringBuilder();
        for (String s : expectedResult) {
            b.append(s);
            b.append(",");
        }
        b.setLength(b.length() - 1);
        return b.toString();
    }

    @Test
    public void testWithRawValuesIgnoringMissingRows() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(23);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(rawValues.get(0));
                    allRowValues.add(rawValues.get(1));
                    allRowValues.add(rawValues.get(2));
                }
            });

            streamer.stream();

            for (String s : allRowValues) {
                System.out.println(s);
            }

            final String[] expectedResult = new String[]{"ASD", "DD", "", "DDD", "", "ASD", "", "FAS", "TAS", "123", "DAS", "AAA", "42715", "23.05", "50", "d65as4", "das4d6", "dasdsa", "6a6dsa5hfghfg", "tertew", "gdfg"};

            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @Test
    public void testWithFormattedValues() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });
            ((XlsxStreamer) streamer).setIgnoreEmptyRow(false);

            streamer.stream();
            final String[] expectedResult = new String[]{"\"ASD\"", "\"DD\"","", "\"DDD\"", "", "\"ASD\"", "", "\"FAS\"", "\"TAS\"", "123", "\"DAS\"", "\"AAA\"", "11/12/16", "23.05", "R$ 50", "\"d65as4\"", "\"das4d6\"", "\"dasdsa\"", "", "", "", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    @Test
    public void testWithFormattedValuesIgnoringEmptyRow() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });

            streamer.stream();
            final String[] expectedResult = new String[]{"\"ASD\"", "\"DD\"","", "\"DDD\"", "", "\"ASD\"", "", "\"FAS\"", "\"TAS\"", "123", "\"DAS\"", "\"AAA\"", "11/12/16", "23.05", "R$ 50", "\"d65as4\"", "\"das4d6\"", "\"dasdsa\"", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
       
    @Test
    public void testStreamByIndex() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });

            streamer.streamSheetByIndex(1);
            final String[] expectedResult = new String[]{"\"d65as4\"","\"das4d6\"", "\"dasdsa\"", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    
    @Test
    public void testStreamByIndexWithEmptyRows() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });
            ((XlsxStreamer) streamer).setIgnoreEmptyRow(false);

            streamer.streamSheetByIndex(1);
            final String[] expectedResult = new String[]{"\"d65as4\"","\"das4d6\"", "\"dasdsa\"","","","", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
       
    @Test
    public void testStreamBySheetName() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });

            streamer.streamSheetByName("Sheet2");
            final String[] expectedResult = new String[]{"\"d65as4\"","\"das4d6\"", "\"dasdsa\"", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
     
    @Test
    public void testStreamBySheetNameWithEmptyRows() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });
            ((XlsxStreamer) streamer).setIgnoreEmptyRow(false);

            streamer.streamSheetByName("Sheet2");
            final String[] expectedResult = new String[]{"\"d65as4\"","\"das4d6\"", "\"dasdsa\"","","","", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};

            for (String s : allRowValues) {
                System.out.println(s);
            }
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
        
    @Test
    public void testDoubleStream() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });

            streamer.streamSheetByIndex(0);
            streamer.streamSheetByName("Sheet2");
            final String[] expectedResult = new String[]{"\"ASD\"", "\"DD\"","", "\"DDD\"", "", "\"ASD\"", "", "\"FAS\"", "\"TAS\"", "123", "\"DAS\"", "\"AAA\"", "11/12/16", "23.05", "R$ 50", "\"d65as4\"", "\"das4d6\"", "\"dasdsa\"", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
          
    @Test
    public void testDoubleStreamWithEmptyRows() {
        try {
            final ArrayList<String> allRowValues = new ArrayList<String>(8);

            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {
                    allRowValues.add(formattedValues.get(0));
                    allRowValues.add(formattedValues.get(1));
                    allRowValues.add(formattedValues.get(2));
                }
            });

            ((XlsxStreamer) streamer).setIgnoreEmptyRow(false);
            
            streamer.streamSheetByIndex(0);
            streamer.streamSheetByName("Sheet2");
            final String[] expectedResult = new String[]{"\"ASD\"", "\"DD\"","", "\"DDD\"", "", "\"ASD\"", "", "\"FAS\"", "\"TAS\"", "123", "\"DAS\"", "\"AAA\"", "11/12/16", "23.05", "R$ 50", "\"d65as4\"", "\"das4d6\"", "\"dasdsa\"","","","", "\"6a6dsa5hfghfg\"", "\"tertew\"", "\"gdfg\""};
            
            assertArrayEquals("Expected to contain: " + concatResult(expectedResult), expectedResult, allRowValues.toArray());
        } catch (InvalidFormatException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (SAXException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (OpenXML4JException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    

    @Test
    public void testFetchByExistenceName() {
        try {
            streamer = new XlsxStreamer(ClassLoader.getSystemResourceAsStream("TestFile2.xlsx"), new DefaultCallback() {
                public void onRow(Long sheetNumber, Long rowNum, ArrayList<String> rawValues, ArrayList<String> formattedValues) {

                }
            });

            streamer.streamSheetByIndex(0);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Parser Error");
        }
    }

    @Test
    public void testFetchByInexistenceName() {

        thrown.expect(SheetNotFoundException.class);
        thrown.expectMessage(is("The sheet with name 'Sheet 23' couldn't be found"));

        try {
            streamer.streamSheetByName("Sheet 23");
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }

    @Test
    public void testFetchByNullName() {

        thrown.expect(IllegalArgumentException.class);
        thrown.expectMessage(is("SheetName can't be null!"));

        try {
            streamer.streamSheetByName(null);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (SheetNotFoundException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @Test
    public void testFetchByEmptyName() {

        thrown.expect(IllegalArgumentException.class);
        thrown.expectMessage(is("SheetName can't be empty!"));

        try {
            streamer.streamSheetByName("");
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (SheetNotFoundException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    /*@Test
    public void testFetchByExistenceIndex() {
        try {
            assertNotNull("Expected to not be null", metadata.getSheetByIndex(1));
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }*/
    @Test
    public void testFetchByInexistenceIndex() {

        thrown.expect(SheetNotFoundException.class);
        thrown.expectMessage(is("The sheet number (20) couldn't be found"));

        try {
            streamer.streamSheetByIndex(20);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        }
    }

    @Test
    public void testFetchByInvalidIndex() {

        thrown.expect(IllegalArgumentException.class);
        thrown.expectMessage(is("Index must be higher than -1!"));

        try {
            streamer.streamSheetByIndex(-2);
        } catch (SAXException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);

            fail("Fail to process file");
        } catch (OpenXML4JException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("OpenXML fail");
        } catch (IOException ex) {
            Logger.getLogger(XlsxMetadataTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (ParserConfigurationException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
            fail("Couldn't open the file");
        } catch (SheetNotFoundException ex) {
            Logger.getLogger(ContentHandlerImplTest.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
}
