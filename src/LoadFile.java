import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.util.Scanner;
public class LoadFile  {
public void loadFile(String filePath){
  String line;
  ParserLine parserLine = new ParserLine();
    try {
        Scanner scanner = new Scanner(new FileReader(filePath));
        while (scanner.hasNextLine()){
            line = scanner.nextLine().trim();
            line = line.replace("*",",");
            parserLine.parseSplitLine(line);
        }
        zakonczoneParsowanie();
    } catch (FileNotFoundException e) {
        e.printStackTrace();
    }
}

public void zakonczoneParsowanie(){
    System.out.println("Zakonczono parsowanie");
}
}
