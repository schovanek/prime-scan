package schovanek.primescan;

import java.io.File;

import org.apache.logging.log4j.Level;
import org.apache.logging.log4j.core.config.Configurator;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class App {
    private static final Logger log = LoggerFactory.getLogger(App.class);
    private static final String APP_LOGGER_NAME = "schovanek.primescan";

    public static void main(String[] args) {

        if (args.length < 1) {
            System.err.println(
                    "usage: java -jar primescan-1.0.jar <xlsx file> [--debug]");
            System.exit(1);
        }

        File xlsxFile = new File(args[0]);
        if (!xlsxFile.isFile()) {
            System.err.println("File not found or not a file: " + xlsxFile.getPath());
            System.exit(1);
        }

        if (args.length > 1 && args[1].equals("--debug")) {
            Configurator.setLevel(APP_LOGGER_NAME, Level.DEBUG);
        }

        int dataColumIdx = 1;
        int sheetIdx = 0;
        // workaround for org.apache.poi.util.RecordFormatException
        IOUtils.setByteArrayMaxOverride(150_000_000);
        try (OPCPackage pkg = OPCPackage.open(xlsxFile.getPath(), PackageAccess.READ)){
            PrimeScan primeScan = new PrimeScan(pkg, System.out, dataColumIdx, sheetIdx);
            primeScan.process();
        } catch (Exception e) {
            System.err.println("Failed to process file: " + xlsxFile.getPath() + ": " + e.getMessage());
            log.error("Failed to process file: {}", xlsxFile.getPath(), e);
            System.exit(1);
        }
    }
}
