package schovanek.primescan;

import org.apache.commons.lang3.StringUtils;
import org.dhatim.fastexcel.reader.ReadableWorkbook;
import org.dhatim.fastexcel.reader.Row;
import org.dhatim.fastexcel.reader.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.math.BigInteger;
import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;
import java.util.stream.Stream;

public final class FastPrimeScan {

    private static final Logger log = LoggerFactory.getLogger(FastPrimeScan.class);

    private static final int PRIME_CERTAINTY = 32;
    private static final String WORKER_NAME = "prime-check-worker";

    // Need a private constant unique instance, not a string literal that can collide with an input value
    private static final String POISON_PILL = new String("__STOP__");

    private final int dataColumnIndex;
    private final int sheetIdx;
    private final BufferedWriter output;

    public FastPrimeScan(PrintStream output, int dataColumnIndex, int sheetIdx) {
        this.dataColumnIndex = dataColumnIndex;
        this.sheetIdx = sheetIdx;
        this.output = new BufferedWriter(new OutputStreamWriter(output));
    }

    public void process(InputStream xlsxInputStream) throws IOException {
        final BlockingQueue<String> queue = new LinkedBlockingQueue<>();
        Thread worker = startWorkerThread(queue);

        try (ReadableWorkbook wb = new ReadableWorkbook(xlsxInputStream)) {
            Sheet sheet = wb.getSheet(sheetIdx).orElseThrow(() -> new IndexOutOfBoundsException(
                    "Requested sheet index " + sheetIdx + " is out of bounds."
            ));

            try (Stream<Row> rows = sheet.openStream()) {
                enqueueNonBlankCellValues(rows, queue);
            }
        } finally {
            stopWorkerThreadAndJoin(queue, worker);
            output.flush();
        }
    }

    private void enqueueNonBlankCellValues(Stream<Row> rows, BlockingQueue<String> queue) {
        rows.forEach(row -> {
            String value = row.getCellText(dataColumnIndex);
            if (StringUtils.isNotBlank(value)) {
                try {
                    queue.put(value);
                } catch (InterruptedException e) {
                    log.debug("Interrupted while adding item to queue.");
                    Thread.currentThread().interrupt();
                    throw new RuntimeException(e);
                }
            }
        });
    }

    private void stopWorkerThreadAndJoin(BlockingQueue<String> queue, Thread worker) {
        try {
            queue.put(POISON_PILL);
            worker.join();
        } catch (InterruptedException e) {
            log.debug("Interrupted while waiting for worker to finish.");
            Thread.currentThread().interrupt();
        }
    }

    private Thread startWorkerThread(BlockingQueue<String> queue) {
        Runnable primeCheckTask = () -> {
            try {
                while (true) {
                    String rawValue = queue.take();

                    // IMPORTANT:
                    // This comparison intentionally uses identity (==), NOT String.equals(),
                    // because no real cell value can reference this exact object.
                    // Do NOT replace with equals().
                    if (rawValue == POISON_PILL) {
                        break;
                    }

                    try {
                        processCellValue(rawValue.trim());
                    } catch (IOException e) {
                        throw new RuntimeException(e);
                    }
                }
            } catch (InterruptedException e) {
                Thread.currentThread().interrupt();
                log.debug("Interrupted while waiting for queue item.");
            }
        };

        Thread worker = new Thread(primeCheckTask, WORKER_NAME);
        worker.start();
        return worker;
    }

    private void processCellValue(String formattedValue) throws IOException {
        try {
            BigInteger number = new BigInteger(formattedValue);

            // skip negative numbers and zeros
            if (number.signum() < 1) {
                return;
            }

            if (number.isProbablePrime(PRIME_CERTAINTY)) {
                output.write(formattedValue);
                output.newLine();
            }
        } catch (NumberFormatException e) {
            log.debug("Failed to parse number: {}", formattedValue);
        }
    }
}
