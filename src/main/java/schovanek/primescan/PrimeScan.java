package schovanek.primescan;

import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.math.BigInteger;

import javax.xml.parsers.ParserConfigurationException;


import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;


public class PrimeScan {
    private static final Logger log = LoggerFactory.getLogger(PrimeScan.class);

    private final OPCPackage xlsxPackage;
    private final int dataColumIdx;
    private final int sheetIdx;
    private final PrintStream output;

    /**
     * Creates a new PrimeScan
     *
     * @param pkg        The XLSX package to process
     * @param output     The PrintStream to output the Primes to
     * @param dataColumIdx  Index of the column that contains the data
     */
    public PrimeScan(OPCPackage pkg, PrintStream output, int dataColumIdx, int sheetIdx) {
        this.xlsxPackage = pkg;
        this.output = output;
        this.dataColumIdx = dataColumIdx;
        this.sheetIdx = sheetIdx;
    }

    public void processSheet(Styles styles, SharedStrings strings, SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {

        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);

        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException e) {
            throw new RuntimeException("SAX parser configuration error: " + e.getMessage(), e);
        }
    }

    public void process() throws IOException, OpenXML4JException, SAXException {
        ReadOnlySharedStringsTable sharedStrings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        SheetIterator sheetIterator = (SheetIterator) xssfReader.getSheetsData();

        try (InputStream sheet = getSheetAt(sheetIterator, sheetIdx)) {
            processSheet(styles, sharedStrings, new PrimesDataHandler(), sheet);
        }
    }

    private class PrimesDataHandler implements SheetContentsHandler {
        @Override
        public void startRow(int rowNum) {
        }

        @Override
        public void endRow(int rowNum) {}

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            // no need to do anything if we do not have a value or reference
            if (StringUtils.isAnyBlank(formattedValue, cellReference)) {
                return;
            }
            // process only the column we are interested in
            if(new CellReference(cellReference).getCol() != dataColumIdx) {
                return;
            }
            try {
                formattedValue = formattedValue.trim();
                BigInteger number = new BigInteger(formattedValue);
                // skip negative numbers and zeros
                if(number.signum() < 1) {
                    return;
                }
                // write out primes
                if (number.isProbablePrime(32)) {
                    output.println(formattedValue);
                }
            } catch (NumberFormatException e) {
                log.debug("Failed to parse number from cell {}: {}", cellReference, e.getMessage());
            }
        }
    }

    public static InputStream getSheetAt(SheetIterator iterator, int index) {
        if (index < 0) {
            throw new IndexOutOfBoundsException("Sheet index must be >= 0, got: " + index);
        }

        int current = 0;
        while (iterator.hasNext()) {
            InputStream sheet = iterator.next();
            if (current == index) {
                return sheet;
            }
            current++;
        }

        throw new IndexOutOfBoundsException(
                "Requested sheet index " + index + ", but workbook has only " + current + " sheets."
        );
    }
}