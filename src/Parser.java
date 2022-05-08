import java.io.File;
import java.io.FileNotFoundException;
import java.net.URISyntaxException;
import java.util.Scanner;

public class Parser {
    public static void main(String[] args) {
        LoadFile loadFile = new LoadFile();
            loadFile.loadFile("./src/dane_GPS.txt");
    }
}
