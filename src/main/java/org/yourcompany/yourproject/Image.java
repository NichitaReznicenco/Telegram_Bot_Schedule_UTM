package org.yourcompany.yourproject;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;

public class Image {

    public static void main(String[] args) {
        String inputPdf = "Anul_II_2024_Semestrul_IV.pdf";
        String outputImage = "table.png";
        String outputExcel = "ScheduleWithImage.xlsx";

        try {
            // Шаг 1: Извлечение картинки из PDF
            PDDocument document = PDDocument.load(new File(inputPdf));
            PDFRenderer renderer = new PDFRenderer(document);

            // Получаем изображение первой страницы с высоким качеством
            BufferedImage image = renderer.renderImageWithDPI(0, 300); // 0 — номер страницы
            ImageIO.write(image, "png", new File(outputImage));
            document.close();

            // Шаг 2: Вставка картинки в Excel
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Таблица");

            InputStream is = new FileInputStream(outputImage);
            byte[] imageBytes = IOUtils.toByteArray(is);
            int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
            is.close();

            CreationHelper helper = workbook.getCreationHelper();
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = helper.createClientAnchor();

            // Расположение картинки (ячейка B2)
            anchor.setCol1(1);
            anchor.setRow1(1);

            Picture pict = drawing.createPicture(anchor, pictureIdx);
            pict.resize(); // авторазмер

            // Сохраняем Excel
            FileOutputStream out = new FileOutputStream(outputExcel);
            workbook.write(out);
            out.close();
            workbook.close();

            System.out.println("✅ Таблица из PDF вставлена как изображение в Excel: " + outputExcel);

        } catch (IOException e) {
            System.err.println("Ошибка: " + e.getMessage());
            e.printStackTrace();
        }
    }
}
