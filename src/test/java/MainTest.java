import org.example.Main;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Base64;
import java.util.List;
import java.util.stream.Collectors;

public class MainTest {
    private static final String OUTPUT_RESULT_FILE_PATH = "./src/test/outputs/result.xlsx";
    private static final String INPUT_TEMPLATE_FILE_PATH = "storage/templates/サンプル請求書.xlsx";
    private static final String INPUT_DATA_FILE_PATH = "./src/test/json/input.json";
    PrintStream old;
    @Test
    void test() throws Exception {
        Path file = Paths.get(INPUT_DATA_FILE_PATH);
        List<String> text = Files.readAllLines(file); // UTF-8
        String json = String.join("", text).replaceAll("\"", "\\\"");
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        setSystemOut(baos);
        String[] param = new String[] {INPUT_TEMPLATE_FILE_PATH, json};
        Main.main(param);
        System.setOut(old);
        String base64String = baos.toString();
        byte[] decoded = Base64.getDecoder().decode(base64String.replaceAll("\\n", "").replaceAll("\\r", ""));
        FileOutputStream fos = new FileOutputStream(OUTPUT_RESULT_FILE_PATH);
        fos.write(decoded);
        fos.close();
    }
    void setSystemOut(ByteArrayOutputStream baos) {
        PrintStream ps = new PrintStream(baos);
        old = System.out;
        System.setOut(ps);
    }
}
