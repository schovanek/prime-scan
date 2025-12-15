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
import org.apache.poi.util.IOUtils;
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

import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;

public class POIPrimeScan {
    private static final Logger log = LoggerFactory.getLogger(POIPrimeScan.class);

    private final int dataColumIdx;
    private final int sheetIdx;
    private final PrintStream output;

    /**
     * Creates a new POIPrimeScan
     *
     * @param output        The PrintStream to output the Primes to
     * @param sheetIdx      Index of the sheet that contains the data
     * @param dataColumIdx  Index of the column that contains the data
     */
    public POIPrimeScan(PrintStream output, int dataColumIdx, int sheetIdx) {
        this.output = output;
        this.dataColumIdx = dataColumIdx;
        this.sheetIdx = sheetIdx;
    }

    public void process(InputStream xlsxInputStream) throws IOException, OpenXML4JException, SAXException {
        // workaround for org.apache.poi.util.RecordFormatException
        IOUtils.setByteArrayMaxOverride(150_000_000);

        try (OPCPackage xlsxPackage = OPCPackage.open(xlsxInputStream)){
            procesPackage(xlsxPackage);
        }
    }

    private void procesPackage(OPCPackage xlsxPackage) throws IOException, SAXException, OpenXML4JException {
        ReadOnlySharedStringsTable sharedStrings = new ReadOnlySharedStringsTable(xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        SheetIterator sheetIterator = (SheetIterator) xssfReader.getSheetsData();

        try (InputStream sheet = getSheetAt(sheetIterator, sheetIdx)) {
            PrimesDataHandler sheetHandler = new PrimesDataHandler(output, dataColumIdx);
            processSheet(styles, sharedStrings, sheetHandler, sheet);
            sheetHandler.finish();
        }
    }

    private void processSheet(Styles styles, SharedStrings strings, SheetContentsHandler sheetHandler,
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

    private static class PrimesDataHandler implements SheetContentsHandler {

        // Need a private constant unique instance, not a string literal that can collide with an input value
        private static final String POISON_PILL = new String("__STOP__");

        private final int dataColumnIdx;
        private final PrintStream output;
        private final BlockingQueue<String> queue = new LinkedBlockingQueue<>();
        private final Thread worker;

        public PrimesDataHandler(PrintStream output, int dataColumnIdx) {
            this.output = output;
            this.dataColumnIdx = dataColumnIdx;
            this.worker = new Thread(this::consumeLoop, "prime-check-worker");
            this.worker.start();
        }

        @Override
        public void startRow(int rowNum) {
        }

        @Override
        public void endRow(int rowNum) {
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            // skip processing when either the value or the cell reference is missing
            if (StringUtils.isAnyBlank(formattedValue, cellReference)) {
                return;
            }
            // skip processing when the cell reference does not match the requested column
            if (new CellReference(cellReference).getCol() != dataColumnIdx) {
                return;
            }

            try {
                queue.put(formattedValue);
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                throw new RuntimeException(e);
            }
        }

        private void consumeLoop() {
            try {
                while (true) {
                    String value = queue.take().trim();

                    // IMPORTANT:
                    // This comparison intentionally uses identity (==), NOT String.equals(),
                    // because no real cell value can reference this exact object.
                    // Do NOT replace with equals().
                    if (value == POISON_PILL) {
                        break;
                    }
                    processValue(value);
                }
            } catch (InterruptedException e) {
                log.debug("Interrupted while waiting for queue item.");
                Thread.currentThread().interrupt();
            }
        }

        private void processValue(String formattedValue) {
            try {
                BigInteger number = new BigInteger(formattedValue);
                // skip negative numbers and zeros
                if (number.signum() < 1) {
                    return;
                }
                // write out primes
                if (number.isProbablePrime(32)) {
                    output.println(formattedValue);
                }
            } catch (NumberFormatException e) {
                // no cellReference, only formattedValue as requested
                log.debug("Failed to parse number: {}", formattedValue);
            }
        }

        /**
         * Call when the document parsing is finished.
         * Waits for the worker to finish processing the queue.
         */
        public void finish() {
            try {
                // signal worker to stop
                queue.put(POISON_PILL);
                worker.join();
            } catch (InterruptedException e) {
                log.debug("Interrupted while waiting for worker to finish.");
                Thread.currentThread().interrupt();
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