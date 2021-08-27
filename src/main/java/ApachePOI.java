import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;

public class ApachePOI {

    public static void main(String[] args) {

        try {
            BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));

            System.out.println("APACHE_POI");
            displayMenu();

            String choice = null;

            do {
                choice = reader.readLine();

                switch (choice) {
                    case "1":
                        ApachePOIReader apachePOIReader = new ApachePOIReader();
                        apachePOIReader.read();
                        break;
                    case "2":
                        ApachePOIWriter apachePOIWriter = new ApachePOIWriter();
                        apachePOIWriter.write();
                        break;
                    default:
                        System.out.println("Incorrect, veuillez sélectionner une option ci dessous");
                        displayMenu();
                }
            } while (!choice.equals("1") || !choice.equals("2"));
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    public static void displayMenu() {
        System.out.println("1 - Lire un fichier excel");
        System.out.println("2 - Écrire un fichier excel");
    }
}
