package org.yourcompany.yourproject;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.pdfbox.text.TextPosition;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class PIOaPDF {

    private static final float TOLERANCE = 5f; // Порог сгруппированных координат

    public static void main(String[] args) {
        String pdfFile = "Anul_II_2024_Semestrul_IV.pdf";
        String excelFile = "TableStructured.xlsx";

        try (PDDocument document = PDDocument.load(new File(pdfFile));
             Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet = workbook.createSheet("Таблица");

            List<TextElement> elements = new ArrayList<>();

            // Извлечение текста с координатами
            PDFTextStripper stripper = new PDFTextStripper() {
                @Override
                protected void writeString(String string, List<TextPosition> textPositions) {
                    for (TextPosition tp : textPositions) {
                        elements.add(new TextElement(
                                Math.round(tp.getXDirAdj()),
                                Math.round(tp.getYDirAdj()),
                                tp.getUnicode()
                        ));
                    }
                }
            };
            stripper.getText(document);

            // Кластеризация X и Y
            Map<Integer, Float> xClusters = clusterCoordinates(elements.stream().map(e -> e.x).toList());
            Map<Integer, Float> yClusters = clusterCoordinates(elements.stream().map(e -> e.y).toList());

            // Построение таблицы: row → col → text
            Map<Integer, Map<Integer, StringBuilder>> table = new TreeMap<>();
            for (TextElement el : elements) {
                int row = findClosestCluster(yClusters, el.y);
                int col = findClosestCluster(xClusters, el.x);

                table.computeIfAbsent(row, r -> new TreeMap<>())
                        .computeIfAbsent(col, c -> new StringBuilder())
                        .append(el.text);
            }

            // Запись в Excel
            for (Map.Entry<Integer, Map<Integer, StringBuilder>> rowEntry : table.entrySet()) {
                Row row = sheet.createRow(rowEntry.getKey());
                for (Map.Entry<Integer, StringBuilder> colEntry : rowEntry.getValue().entrySet()) {
                    Cell cell = row.createCell(colEntry.getKey());
                    cell.setCellValue(colEntry.getValue().toString().trim());
                    CellStyle style = workbook.createCellStyle();
                    style.setWrapText(true);
                    cell.setCellStyle(style);
                }
            }

            // Автоширина
            for (int i = 0; i < 50; i++) {
                sheet.autoSizeColumn(i);
            }

            try (FileOutputStream out = new FileOutputStream(excelFile)) {
                workbook.write(out);
            }

            System.out.println("✅ Таблица из PDF успешно перенесена в Excel: " + excelFile);

        } catch (IOException e) {
            System.err.println("Ошибка при обработке PDF:");
            e.printStackTrace();
        }
    }

    // Текст с координатами
    static class TextElement {
        float x, y;
        String text;
        TextElement(float x, float y, String text) {
            this.x = x;
            this.y = y;
            this.text = text;
        }
    }

    // Кластеризация координат
    private static Map<Integer, Float> clusterCoordinates(List<Float> coords) {
        List<Float> sorted = new ArrayList<>(new HashSet<>(coords));
        Collections.sort(sorted);

        Map<Integer, Float> clusterMap = new TreeMap<>();
        int clusterId = 0;
        Float last = null;

        for (Float c : sorted) {
            if (last == null || Math.abs(c - last) > TOLERANCE) {
                clusterMap.put(clusterId++, c);
                last = c;
            }
        }
        return clusterMap;
    }

    // Поиск ближайшего кластера
    private static int findClosestCluster(Map<Integer, Float> clusters, float coord) {
        return clusters.entrySet().stream()
                .min(Comparator.comparingDouble(e -> Math.abs(e.getValue() - coord)))
                .map(Map.Entry::getKey)
                .orElse(0);
    }
}
