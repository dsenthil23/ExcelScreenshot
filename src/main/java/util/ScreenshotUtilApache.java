package util;

import org.apache.poi.ss.usermodel.*;
import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class ScreenshotUtilApache{

public static void excelSSApache(String excelPath, String imagePath) throws Exception {
    try (FileInputStream fis = new FileInputStream(excelPath);
         Workbook workbook = WorkbookFactory.create(fis)) {
        Sheet sheet = workbook.getSheetAt(0);

        int rows = sheet.getLastRowNum() + 1;
        int cols = 0;
        for (Row row : sheet) {
            if (row != null) {
                cols = Math.max(cols, row.getLastCellNum());
            }
        }

        Font font = new Font("Arial", Font.PLAIN, 14);
        BufferedImage tempImg = new BufferedImage(1, 1, BufferedImage.TYPE_INT_RGB);
        Graphics2D tempG2d = tempImg.createGraphics();
        tempG2d.setFont(font);
        FontMetrics metrics = tempG2d.getFontMetrics();

        int maxCellWidth = 0;
        int maxCellHeight = metrics.getHeight() + 10; // padding

        for (int rowIdx = 0; rowIdx < rows; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            for (int colIdx = 0; colIdx < cols; colIdx++) {
                String text = "";
                if (row != null) {
                    Cell cell = row.getCell(colIdx);
                    if (cell != null) {
                        text = cell.toString();
                    }
                }
                int textWidth = metrics.stringWidth(text) + 10; // padding
                if (textWidth > maxCellWidth) {
                    maxCellWidth = textWidth;
                }
            }
        }
        tempG2d.dispose();

        int cellWidth = maxCellWidth;
        int cellHeight = maxCellHeight;
        int imageWidth = cols * cellWidth;
        int imageHeight = rows * cellHeight;

        // Generate timestamp
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        File outFile = new File(imagePath);
        String parent = outFile.getParent();
        String name = outFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = (dotIndex == -1) ? name : name.substring(0, dotIndex);
        String ext = (dotIndex == -1) ? "png" : name.substring(dotIndex + 1);
        String newFileName = baseName + "_" + timestamp + "." + ext;
        String newFilePath = (parent != null) ? parent + File.separator + newFileName : newFileName;

        BufferedImage image = new BufferedImage(imageWidth, imageHeight, BufferedImage.TYPE_INT_RGB);
        Graphics2D g2d = image.createGraphics();

        g2d.setColor(Color.WHITE);
        g2d.fillRect(0, 0, imageWidth, imageHeight);

        g2d.setColor(Color.BLACK);
        g2d.setFont(font);

        for (int rowIdx = 0; rowIdx < rows; rowIdx++) {
            Row row = sheet.getRow(rowIdx);
            for (int colIdx = 0; colIdx < cols; colIdx++) {
                int x = colIdx * cellWidth;
                int y = rowIdx * cellHeight;
                g2d.drawRect(x, y, cellWidth, cellHeight);

                String text = "";
                if (row != null) {
                    Cell cell = row.getCell(colIdx);
                    if (cell != null) {
                        text = cell.toString();
                    }
                }
                g2d.drawString(text, x + 5, y + metrics.getAscent() + 5);
            }
        }

        g2d.dispose();
        ImageIO.write(image, ext, new File(newFilePath));
        System.out.println("Excel rendered as image: " + newFilePath);
    }
}
}
