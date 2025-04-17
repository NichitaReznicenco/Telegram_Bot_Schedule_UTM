package org.yourcompany.yourproject;

import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

public class JavaSmt {

    public static void main(String[] args) {
        System.out.println("Начинаем конвертацию PDF → Excel...");

        String inputFile = "Anul_II_2024_Semestrul_IV.pdf";
        String outputFile = "Anul_II_2024.xlsx"; // Итоговый файл

        try {
            // 1. Конвертация PDF → Excel (во временный файл)
            PdfDocument pdf = new PdfDocument();
            pdf.loadFromFile(inputFile);
            File tempFile = File.createTempFile("converted_", ".xlsx");
            pdf.saveToFile(tempFile.getAbsolutePath(), FileFormat.XLSX);
            pdf.close();
            System.out.println("PDF сконвертирован во временный файл.");

            // 2. Открытие Excel и обработка
            FileInputStream file = new FileInputStream(tempFile);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheetAt(0);

            // ❗ Удаляем все объединённые области перед объединением вручную
            while (sheet.getNumMergedRegions() > 0) {
                sheet.removeMergedRegion(0);
            }

            // Объединяем заголовки групп (разные строки + переносы)
            mergeGroupHeaders(sheet);

            file.close();

            // 3. Сохраняем результат
            FileOutputStream outFile = new FileOutputStream(outputFile);
            workbook.write(outFile);
            workbook.close();
            outFile.close();

            System.out.println("Файл успешно очищен и сохранён как: " + outputFile);

        } catch (Exception e) {
            System.err.println("Произошла ошибка при обработке файла:");
            e.printStackTrace();
        }
    }

    /**
     * Объединяет заголовки групп вида CR- + 231 из разных строк и из одной ячейки с переносом строки.
     */
    private static void mergeGroupHeaders(Sheet sheet) {
        // Часть 1: объединяем значения из двух строк
        for (int i = 0; i < sheet.getLastRowNum(); i++) {
            Row row1 = sheet.getRow(i);
            Row row2 = sheet.getRow(i + 1);

            if (row1 == null || row2 == null) continue;

            for (int j = 0; j < row1.getLastCellNum(); j++) {
                Cell cell1 = row1.getCell(j);
                Cell cell2 = row2.getCell(j);

                if (cell1 == null || cell2 == null) continue;

                if (cell1.getCellType() == CellType.STRING && cell2.getCellType() == CellType.STRING) {
                    String val1 = cell1.getStringCellValue().trim();
                    String val2 = cell2.getStringCellValue().trim();

                    if (val1.endsWith("-") && val2.matches("\\d+")) {
                        cell1.setCellValue(val1 + val2);
                        cell2.setCellValue("");
                    }
                }
            }
        }

        // Часть 2: объединяем значения внутри одной ячейки (перенос строки или пробел)
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String val = cell.getStringCellValue().trim();

                    if (val.contains("\n")) {
                        String[] parts = val.split("\n");
                        if (parts.length == 2 && parts[0].trim().endsWith("-") && parts[1].trim().matches("\\d+")) {
                            cell.setCellValue(parts[0].trim() + parts[1].trim());
                        }
                    } else if (val.matches("[A-Z]{2,3}-\\s*\\d+")) {
                        // Обрабатываем случаи вроде "SI- 231"
                        cell.setCellValue(val.replaceAll("\\s+", ""));
                    }
                }
            }
        }
    }
}







//    /**
//     * Удаляет строки до указанного индекса и пересчитывает объединённые ячейки.
//     */
//    private static void removeHeaderRows(Sheet sheet, int lastHeaderRowIndex) {
//        // Удаляем строки
//        for (int i = 0; i <= lastHeaderRowIndex; i++) {
//            Row row = sheet.getRow(i);
//            if (row != null) sheet.removeRow(row);
//        }
//
//        // Удаляем все merged regions (чтобы не было конфликтов)
//        for (int i = sheet.getNumMergedRegions() - 1; i >= 0; i--) {
//            sheet.removeMergedRegion(i);
//        }
//
//        // Сдвигаем оставшиеся строки вверх
//        sheet.shiftRows(lastHeaderRowIndex + 1, sheet.getLastRowNum(), -(lastHeaderRowIndex + 1));
//    }