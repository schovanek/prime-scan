package schovanek.primescan;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.io.PrintStream;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;

class FastPrimeScanTest {

    public static final int DATA_COLUMN_IDX = 1;
    public static final int SHEET_IDX = 0;

    @Test
    void whenLargeXlsxThenOutputHasExpectedLength() {
        String testFile = "/data/990K64BitNumbers.xlsx";
        ByteArrayOutputStream arrayOutputStream = new ByteArrayOutputStream();
        PrintStream output = new PrintStream(arrayOutputStream);

        try {
            InputStream xlsxInputStream = getClass().getResourceAsStream(testFile);
            FastPrimeScan primeScan = new FastPrimeScan(output, DATA_COLUMN_IDX, SHEET_IDX);
            primeScan.process(xlsxInputStream);
        } catch (Exception e) {
            Assertions.fail("Failed to process file: " + testFile, e);
        }

        List<String> result = arrayOutputStream.toString().lines().toList();
        assertEquals(22730, result.size());
        assertEquals("9223372036854775783", result.getFirst());
        assertEquals("9223372036853785847", result.getLast());
    }

    @ParameterizedTest
    @MethodSource("provideXlsxFilePathsAndExpectedOutputs")
    void whenValidXlsxThenExpectedPrimesFound(String testFile, String expectedOutput) {
        ByteArrayOutputStream arrayOutputStream = new ByteArrayOutputStream();
        PrintStream output = new PrintStream(arrayOutputStream);

        try {
            InputStream xlsxInputStream = getClass().getResourceAsStream(testFile);
            FastPrimeScan primeScan = new FastPrimeScan(output, DATA_COLUMN_IDX, SHEET_IDX);
            primeScan.process(xlsxInputStream);
        } catch (Exception e) {
            Assertions.fail("Failed to process file: " + testFile, e);
        }
        String result = arrayOutputStream.toString();

        assertEquals(expectedOutput, result);
    }

    static Stream<Arguments> provideXlsxFilePathsAndExpectedOutputs() {
        return Stream.of(
                Arguments.of("/data/vzorek_dat.xlsx", """
                        5645657
                        15619
                        1234187
                        211
                        7
                        9788677
                        23311
                        54881
                        2147483647
                        """),
                Arguments.of("/data/corner_cases.xlsx", """
                        2
                        3
                        3
                        7
                        340282366920938463463374607431768211297
                        17
                        23
                        """),
                // poison pill in Data column should not stop processing
                Arguments.of("/data/poison_pill.xlsx", """
                        7
                        13
                        """)
        );
    }
}