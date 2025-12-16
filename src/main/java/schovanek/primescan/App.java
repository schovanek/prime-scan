package schovanek.primescan;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.core.config.Configurator;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.InputStream;
import java.nio.file.Files;

public class App {

    private static final Logger log = LoggerFactory.getLogger(App.class);
    private static final String APP_LOGGER_NAME = "schovanek.primescan";

    private static final String DEBUG_FLAG = "--debug";
    private static final String USAGE = "usage: java -jar primescan-2.0.jar <xlsx file> [--debug]";

    private static final int DATA_COLUMN_INDEX = 1;
    private static final int SHEET_INDEX = 0;

    public static void main(String[] args) {
        CliArgs cli = parseArgs(args);

        File xlsxFile = new File(cli.xlsxPath());
        if (!xlsxFile.isFile()) {
            fail("File not found or not a file: " + xlsxFile.getPath());
        }

        if (cli.debugEnabled()) {
            Configurator.setLevel(APP_LOGGER_NAME, Level.DEBUG);
        }

        try {
            try (InputStream xlsxInputStream = Files.newInputStream(xlsxFile.toPath())) {
                FastPrimeScan primeScan = new FastPrimeScan(System.out, DATA_COLUMN_INDEX, SHEET_INDEX);
                primeScan.process(xlsxInputStream);
            }
        } catch (Exception e) {
            log.debug("Failed to process file: {}", xlsxFile.getPath(), e);
            fail("Failed to process file: " + xlsxFile.getPath() + ": " + e.getMessage());
        }
    }

    private record CliArgs(String xlsxPath, boolean debugEnabled) {}

    private static CliArgs parseArgs(String[] args) {
        if (args.length < 1) {
            fail(USAGE);
        }
        boolean debugEnabled = args.length > 1 && DEBUG_FLAG.equals(args[1]);
        return new CliArgs(args[0], debugEnabled);
    }

    private static void fail(String message) {
        System.err.println(message);
        System.exit(1);
    }
}
