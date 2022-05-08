import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedList;
import java.util.List;

public class ParserLine {
    String[] lista;
    String[] splitOutput;
    String[] naglowki = {"Odbiornik", "Data", "Czas(UTC)", "Odchylenie magnetyczne Ziemi",
            "Wysokość geoid", "Sposób określenia pozycji", "Jakość pomiaru",
            "Suma kontrolna", "Szerokość geogr.", "Długość geogr.", "Prędkość (w węzłach)",
            "Kąt przemieszczania się", "Wysokość n.p.m.", "Precyzja horyzontalna",
            "Precyzja wertykalna", "Precyzja (ogólnie)", "Liczba śledzonych satelitów",
            "Numery satelitów do pozycjonowania", "Status urządzenia", "Czas ustalania pozycji",
            "Liczba widocznych satelitów", "ID satelity (numer)", "Wyniesienie nad równik (w stopniach)",
            "Azymut satelity (w stopniach)", "Stosunek sygnał/szum (SNR) satelity"
    };

    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet("Parsowane dane");
    Row row;
    int indLinii = 1;
    int rowCount = 0;
    int columnCount = 0;

    public String[] splitLine(String line) {
        lista = line.split(",");
        return lista;
    }

    public void parseSplitLine(String line) {
        splitOutput = splitLine(line);

        if (sprawdzenieLinii(line) == true) {
            if (splitOutput[0].equals("$GPRMC")) {
                odbiornikRMC(line);
            } else if (splitOutput[0].equals("$GPGGA")) {
                odbiornikGPGGA(line);
            } else if (splitOutput[0].equals("$GPGSA")) {
                odbiornikGPGSA(line);
            } else if (splitOutput[0].equals("$GPGSV")) {
                if (splitOutput[2].equals(String.valueOf(indLinii))) {
                    odbiornikGPGSV(line);
                    indLinii++;
                    if (indLinii > Integer.parseInt(splitOutput[1])) {
                        indLinii = 1;
                    }
                } else {
                    row = sheet.createRow(rowCount);
                    Cell cell = row.createCell(0);
                    CellStyle cellStyle = workbook.createCellStyle();
                    cell.setCellValue("Brak opisu poprzedniej satelity");
                    sheet.autoSizeColumn(columnCount);
                    cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(cellStyle);
                    rowCount += 2;
                    try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                        workbook.write(outputStream);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }

            } else if (splitOutput[0].equals("$GPGLL")) {
                odbiornikGPGLL(line);
            } else if (splitOutput[0].equals("$GPVTG")) {
                odbiornikGPVTG(line);
            } else {
                row = sheet.createRow(rowCount);
                Cell cell = row.createCell(0);
                cell.setCellValue("Dane nie sa obslugiwane: " + line);
                sheet.autoSizeColumn(columnCount);
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cell.setCellStyle(cellStyle);
                rowCount += 2;
                try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                    workbook.write(outputStream);
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        } else {
            row = sheet.createRow(rowCount);
            Cell cell = row.createCell(0);
            CellStyle cellStyle = workbook.createCellStyle();
            cell.setCellValue("Blad linii: "+ line);
            sheet.autoSizeColumn(columnCount);
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
            rowCount += 2;
            try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                workbook.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public boolean sprawdzenieLinii(String line) {
        if (line.startsWith("$") && line.length() <= 80 && line.contains(",")) {
            return true;
        }
        return false;
    }

    public void odbiornikRMC(String line) {
        List<String> str = new LinkedList<>();
        Cell cell;
        CellStyle cellStyle = workbook.createCellStyle();
        int[] kolumny = {0, 2, 18, 8, 9, 10, 11, 1, 3, 7};
        int ikol = 0;
        String napis;
        String sprawdzenie;

        if (!lista[lista.length - 2].equals("A") && !(lista[lista.length - 2].equals("V"))) {
            str.add(lista[0].replace("$", " ").trim());
            napis = lista[1].charAt(0) + "" + lista[1].charAt(1) + ":" + lista[1].charAt(2) + "" + lista[1].charAt(3) + ":" + lista[1].charAt(4) + "" + lista[1].charAt(5);
            str.add(napis);
            if (lista[2].equals("A")) {
                str.add("aktywny");
            } else if (lista[2].equals("V")) {
                str.add("nieaktywny");
            }
            napis = lista[3].substring(0, lista[3].indexOf(".") - 2) + "°" + lista[3].substring(lista[3].indexOf(".") - 2) + "' " + lista[4];
            str.add(napis);
            napis = lista[5].substring(0, lista[5].indexOf(".") - 2) + "°" + lista[5].substring(lista[5].indexOf(".") - 2) + "' " + lista[6];
            str.add(napis);
            str.add(lista[7]);
            str.add(lista[8]);
            napis = lista[9].charAt(0) + "" + lista[9].charAt(1) + "." + lista[9].charAt(2) + "" + lista[9].charAt(3) + "." + lista[9].charAt(4) + "" + lista[9].charAt(5) + "";
            str.add(napis);
            napis = lista[10].charAt(0) + "" + lista[10].charAt(1) + "" + lista[10].charAt(2) + "" + lista[10].charAt(3) + "" + lista[10].charAt(4) + "°" + " " + lista[11];
            sprawdzenie = napis.charAt(0) + "";
            if (sprawdzenie.equals("0")) {
                napis = napis.replace("0", "");
                napis = napis.trim();
            }
            str.add(napis);
            str.add(lista[12]);

            row = sheet.createRow(rowCount);
            for (String lin : naglowki) {
                cell = row.createCell(columnCount++);
                cell.setCellValue(lin);
                sheet.autoSizeColumn(columnCount);
                cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cell.setCellStyle(cellStyle);
            }
            row = sheet.createRow(++rowCount);
            columnCount = 0;
            for (String lin : str) {
                cell = row.createCell(kolumny[ikol]);
                cell.setCellValue(lin);
                sheet.autoSizeColumn(columnCount);
                ikol++;
            }
            rowCount += 2;
            columnCount = 0;
            try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                workbook.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            row = sheet.createRow(rowCount);
            cell = row.createCell(0);
            cellStyle = workbook.createCellStyle();
            cell.setCellValue("Blad linii: "+ line);
            sheet.autoSizeColumn(columnCount);
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
            rowCount += 2;
            try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                workbook.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }

    public void odbiornikGPGGA(String line) {
        List<String> str = new LinkedList<>();
        int[] kolumny = {0, 2, 8, 9, 6, 16, 13, 12, 4, 7};
        String napis;

        str.add(lista[0].replace("$", " ").trim());
        napis = lista[1].charAt(0) + "" + lista[1].charAt(1) + ":" + lista[1].charAt(2) + "" + lista[1].charAt(3) + ":" + lista[1].charAt(4) + "" + lista[1].charAt(5);
        str.add(napis);
        napis = lista[2].substring(0, lista[2].indexOf(".") - 2) + "°" + lista[2].substring(lista[2].indexOf(".") - 2) + "' " + lista[3];
        str.add(napis);
        napis = lista[4].substring(0, lista[4].indexOf(".") - 2) + "°" + lista[4].substring(lista[4].indexOf(".") - 2) + "' " + lista[5];
        str.add(napis);
        str.add(lista[6]);
        napis = lista[7].charAt(0) + "";
        if (napis.equals("0")) {
            str.add(lista[7].charAt(1) + "");
        } else {
            str.add(lista[7]);
        }
        str.add(lista[8]);
        napis = lista[9] + "" + lista[10].toLowerCase();
        str.add(napis);
        napis = lista[11] + "" + lista[12].toLowerCase();
        str.add(napis);
        str.add(lista[lista.length - 1]);

        row = sheet.createRow(rowCount);
        CellStyle cellStyle = workbook.createCellStyle();
        for (String lin : naglowki) {
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
        }
        row = sheet.createRow(++rowCount);
        columnCount = 0;
        int ikol = 0;
        for (String lin : str) {
            Cell cell = row.createCell(kolumny[ikol]);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            ikol++;
        }
        rowCount += 2;
        columnCount = 0;
        try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void odbiornikGPGSA(String line) {
        List<String> str = new LinkedList<>();
        int[] kolumny = {0, 5, 15, 13, 14, 7};
        String napis = "";

        row = sheet.createRow(rowCount);
        for (String lin : naglowki) {
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
        }
        row = sheet.createRow(++rowCount);
        columnCount = 0;

        str.add(lista[0].replace("$", " ").trim());
        String sprawdzenie;
        sprawdzenie = lista[1] + "";
        if (sprawdzenie.equals("A")) {
            napis = "automatyczny, ";
        } else if (sprawdzenie.equals("M")) {
            napis = "manualny, ";
        }
        sprawdzenie = lista[2] + "";
        if (sprawdzenie.equals("1")) {
            napis += "brak ustalonej pozycji";
        } else if (sprawdzenie.equals("2")) {
            napis += "2D";
        } else if (sprawdzenie.equals("3")) {
            napis += "3D";
        }
        str.add(napis);
        int i = 4;
        String ce = lista[3];
        while (i < 15) {
            if (!lista[i].equals("")) {
                Cell cell = row.createCell(17);
                cell.setCellValue(ce + ", " + lista[i]);
                try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                    workbook.write(outputStream);
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                ce = cell.getStringCellValue();
            }
            i++;
        }
        str.add(lista[15]);
        str.add(lista[16]);
        str.add(lista[17]);
        str.add(lista[18]);

        int ikol = 0;
        for (String lin : str) {
            Cell cell = row.createCell(kolumny[ikol]);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            ikol++;
        }
        rowCount += 2;
        columnCount = 0;
        try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void odbiornikGPGSV(String line) {
        List<String> str = new LinkedList<>();
        int[] kolumny = {0, -1, -1, 20, 7};
        String napis = "";

        row = sheet.createRow(rowCount);
        for (String lin : naglowki) {
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
        }
        row = sheet.createRow(++rowCount);
        columnCount = 0;

        str.add(lista[0].replace("$", " ").trim());
        str.add(lista[1]);
        str.add(lista[2]);
        napis = lista[3].charAt(0) + "";
        if (napis.equals("0")) {
            str.add(lista[3].charAt(1) + "");
        } else {
            str.add(lista[3]);
        }
        int i = 4;
        int[] kolumnySatelity = {21, 22, 23, 24, 21, 22, 23, 24, 21, 22, 23, 24, 21, 22, 23, 24};
        int ikolSat = 0;
        String ce = "";
        int zmiennapom = 1;
        while (i < lista.length - 1) {
            if (!lista[i].equals("")) {
                if (zmiennapom <= 4) {
                    Cell cell = row.createCell(kolumnySatelity[ikolSat]);
                    cell.setCellValue(lista[i]);
                    try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                        workbook.write(outputStream);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    zmiennapom++;
                } else {
                    Cell cell = row.getCell(kolumnySatelity[ikolSat]);
                    ce = cell.getStringCellValue();
                    cell.setCellValue(ce + ", " + lista[i]);

                    try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                        workbook.write(outputStream);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                    zmiennapom++;
                }
                ikolSat++;
            }
            i++;
        }
        str.add(lista[lista.length - 1]);

        int ikol = 0;
        for (String lin : str) {
            if (kolumny[ikol] != -1) {
                Cell cell = row.createCell(kolumny[ikol]);
                cell.setCellValue(lin);
                sheet.autoSizeColumn(columnCount);
            }
            ikol++;
        }
        rowCount += 2;
        columnCount = 0;
        try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void odbiornikGPGLL(String line) {
        List<String> str = new LinkedList<>();
        int[] kolumny = {0, 8, 9, 19, 18, 7};
        String napis = "";

        if ((!line.contains("V") && line.indexOf("A") == line.lastIndexOf("A")) || (!line.contains("A") && line.indexOf("V") == line.lastIndexOf("V"))) {
            str.add(lista[0].replace("$", " ").trim());

            napis = lista[1].substring(0, lista[1].indexOf(".") - 2) + "°" + lista[1].substring(lista[1].indexOf(".") - 2) + "' " + lista[2];
            str.add(napis);
            napis = lista[3].substring(0, lista[3].indexOf(".") - 2) + "°" + lista[3].substring(lista[3].indexOf(".") - 2) + "' " + lista[4];
            str.add(napis);
            napis = lista[5].charAt(0) + "" + lista[5].charAt(1) + ":" + lista[5].charAt(2) + "" + lista[5].charAt(3) + ":" + lista[5].charAt(4) + "" + lista[5].charAt(5);
            str.add(napis);
            if (lista[6].equals("A")) {
                str.add("aktywny");
            } else if (lista[6].equals("V")) {
                str.add("nieaktywny");
            }
            str.add(lista[lista.length - 1]);

            row = sheet.createRow(rowCount);
            for (String lin : naglowki) {
                Cell cell = row.createCell(columnCount++);
                cell.setCellValue(lin);
                sheet.autoSizeColumn(columnCount);
                CellStyle cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cell.setCellStyle(cellStyle);
            }
            row = sheet.createRow(++rowCount);
            columnCount = 0;
            int ikol = 0;
            for (String lin : str) {
                Cell cell = row.createCell(kolumny[ikol]);
                cell.setCellValue(lin);
                sheet.autoSizeColumn(columnCount);
                ikol++;
            }
            rowCount += 2;
            columnCount = 0;
            try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                workbook.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            row = sheet.createRow(rowCount);
            Cell cell = row.createCell(0);
            CellStyle cellStyle = workbook.createCellStyle();
            cell.setCellValue("Blad linii: "+ line);
            sheet.autoSizeColumn(columnCount);
            cellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
            rowCount += 2;
            try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
                workbook.write(outputStream);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    public void odbiornikGPVTG(String line) {
        List<String> str = new LinkedList<>();
        int[] kolumny = {0, -1, -1, 10, -1, 7};
        String napis = "";

        str.add(lista[0].replace("$", " ").trim());
        napis = lista[1] + ", " + lista[2];
        str.add(napis);
        napis = lista[3] + ", " + lista[4];
        str.add(napis);
        napis = lista[5] + ", " + lista[6];
        str.add(napis);
        napis = lista[7] + ", " + lista[8];
        str.add(napis);
        str.add(lista[lista.length - 1]);
        row = sheet.createRow(rowCount);
        for (String lin : naglowki) {
            Cell cell = row.createCell(columnCount++);
            cell.setCellValue(lin);
            sheet.autoSizeColumn(columnCount);
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cell.setCellStyle(cellStyle);
        }
        row = sheet.createRow(++rowCount);
        columnCount = 0;
        int ikol = 0;
        for (String lin : str) {
            if (kolumny[ikol] != -1) {
                Cell cell = row.createCell(kolumny[ikol]);
                cell.setCellValue(lin);
            }
            sheet.autoSizeColumn(columnCount);
            ikol++;
        }
        ikol = 0;
        int iPod = 0;
        String[] podpisy = {"Ścieżka poruszania się (w stopniach):", "Ścieżka poruszania się na podstawie " +
                "danych magnetycznych (w stopniach):", "Prędkość (w km/h):"};
        for (String lin : str) {
            if (kolumny[ikol] == -1) {
                row = sheet.createRow(++rowCount);
                Cell cell = row.createCell(0);
                cell.setCellValue(podpisy[iPod] + " " + lin);
                iPod++;
            }
            sheet.autoSizeColumn(columnCount);
            ikol++;
        }
        rowCount += 2;
        columnCount = 0;
        try (FileOutputStream outputStream = new FileOutputStream("Odbiorniki.xlsx")) {
            workbook.write(outputStream);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}