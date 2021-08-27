import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public abstract class KeyboardReader {
    public static String askForKeyboardInteraction(String message) throws IOException {
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));

        System.out.println(message);

        return reader.readLine();
    }
}
