package com.company;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Iterator;



public class Excel {

    private static File file;
    static private Workbook wb; // книга
    private static FileInputStream inputStream;
    static private Sheet sheet;
    private static int excelLines;
    static private RWSetting v = new RWSetting();
    static  private CreateFile createFile = new CreateFile();

    public static void addToExcell(String flexor, String log) throws IOException, InterruptedException {
        file = new File("Data.xls");

        // Read XSL fil
        if (file.exists()) {
            inputStream = new FileInputStream(file);
            wb = new HSSFWorkbook(inputStream);
            // Get first sheet from the workbook
            sheet = wb.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            excelLines = 0;
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                // Get iterator to all cells of current row
                Iterator<Cell> cellIterator = row.cellIterator();
                excelLines++;
            }
            //v.WriteValues("TotalExcelLines", excelLines);
            inputStream.close();
            String s = flexor.replaceAll("[;]+", ";");
            String s1 = s.replaceAll("\\s+", "");
            int rowNum = excelLines;
            String[] parts = s1.split(";");
            String patientNum = parts[3];
            if (parts[0].equals("\u0002{I")) {
                System.out.println("Info message was ignored");
                // break;
            } else {
                for (int j = 4; j < parts.length - 1; j++) {
                    switch (parts[j]) {
                        case "400":
                            writeParts(rowNum, patientNum, parts[j], "Білірубін загальний", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "401":
                            writeParts(rowNum, patientNum, parts[j], "Білірубін прямий", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "402":
                            writeParts(rowNum, patientNum, parts[j], "АСТ", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "403":
                            writeParts(rowNum, patientNum, parts[j], "АЛТ", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "404":
                            writeParts(rowNum, patientNum, parts[j], "Лужна фосфатаза", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "406":
                            writeParts(rowNum, patientNum, parts[j], "ГТП", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "407":
                            writeParts(rowNum, patientNum, parts[j], "Білок", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "408":
                            String alb = "Альбумін";
                            writeParts(rowNum, patientNum, parts[j], alb, parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "409":
                            writeParts(rowNum, patientNum, parts[j], "Сечовина", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "410":
                            writeParts(rowNum, patientNum, parts[j], "Креатинін", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "411":
                            writeParts(rowNum, patientNum, parts[j], "Сечова кислота", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "412":
                            writeParts(rowNum, patientNum, parts[j], "Холестерин", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "413":
                            writeParts(rowNum, patientNum, parts[j], "Тригліцериди", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "414":
                            writeParts(rowNum, patientNum, parts[j], "Ліпопротеїди фракціонно", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "415":
                            writeParts(rowNum, patientNum, parts[j], "Глюкоза сироватки крові", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "416":
                            writeParts(rowNum, patientNum, parts[j], "Альфа-амілаза", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "418":
                            writeParts(rowNum, patientNum, parts[j], "Фосфор неорганічний", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "419":
                            writeParts(rowNum, patientNum, parts[j], "Залізо", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "420":
                            writeParts(rowNum, patientNum, parts[j], "Кадьцій", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "422":
                            writeParts(rowNum, patientNum, parts[j], "Магній", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "424":
                            writeParts(rowNum, patientNum, parts[j], "ЛДГ", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "425":
                            writeParts(rowNum, patientNum, parts[j], "Холінестераза", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "433":
                            writeParts(rowNum, patientNum, parts[j], "Глюкозотолерантний тест", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "434":
                            writeParts(rowNum, patientNum, parts[j], "СК", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "435":
                            writeParts(rowNum, patientNum, parts[j], "Ліпаза", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "439":
                            writeParts(rowNum, patientNum, parts[j], "Трансферин", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;
                        case "440":
                            writeParts(rowNum, patientNum, parts[j], "ОЖСС", parts[j + 1]);
                            j++;
                            rowNum++;
                            break;

                    }

                }
            }

            try {
                FileOutputStream fos = new FileOutputStream(file);
                wb.write(fos);
                fos.close();



            } catch (Exception ex) {
                System.out.println(ex.toString());
                System.out.println("New Data can't be written!");

            }


        } else {
           createFile.createExcel();
        }
    }

    private static void writeParts(int rowNum, String patientNum, String analysis, String analysisName, String result) {
        Row row = sheet.createRow(rowNum);
        Cell patientCell = row.createCell(0);

        int patientID = Integer.parseInt(patientNum);
        patientCell.setCellValue(patientID);
        Cell analysisIdCell = row.createCell(1);

        int analysisID = Integer.parseInt(analysis);
        analysisIdCell.setCellValue(analysisID);
        Cell analysisNameCell = row.createCell(2);
        analysisNameCell.setCellValue(analysisName);
        Cell resultCell = row.createCell(3);

        try {
            String res = result.replaceAll(",", ".");
            double resultNum = Double.parseDouble(res);
            resultCell.setCellValue(resultNum);
        } catch (Exception ex) {
        }
        Cell dateCell = row.createCell(4);
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("dd.MM.yyyy");
        LocalDate localDate = LocalDate.now();
        dateCell.setCellValue(dtf.format(localDate));

    }

}
