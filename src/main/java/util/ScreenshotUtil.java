package util;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

public class ScreenshotUtil {

    static String timestamp = new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new java.util.Date());

    public static void captureSS(String excelFilePath, String outputFilePath) throws AWTException, IOException, InterruptedException {
        Desktop.getDesktop().open(new File(excelFilePath));
        Thread.sleep(5000);

        File outFile = new File(outputFilePath);
        String parent = outFile.getParent();
        String name = outFile.getName();
        int dotIndex = name.lastIndexOf('.');
        String baseName = (dotIndex == -1) ? name : name.substring(0, dotIndex);
        String ext = (dotIndex == -1) ? "png" : name.substring(dotIndex + 1);
        String newFileName = baseName + "_" + timestamp + "." + ext;
        String newFilePath = (parent != null) ? parent + File.separator + newFileName : newFileName;
        Robot robot = new Robot();
        Rectangle rectangle = new Rectangle(Toolkit.getDefaultToolkit().getScreenSize());
        BufferedImage screenCapture = robot.createScreenCapture(rectangle);
        ImageIO.write(screenCapture, ext, new File(newFilePath));
        System.out.println("Screenshot Location: " + newFilePath);
        Runtime.getRuntime().exec("taskkill /IM excel.exe /F");
    }


    //jxl library
    public static void renderExcelAsImage(String excelPath, String imagePath) throws Exception {
        Workbook workbook = Workbook.getWorkbook(new File(excelPath));
        Sheet sheet = workbook.getSheet(0);

        int rows = sheet.getRows();
        int cols = sheet.getColumns();
        File outFile = new File(imagePath);
        String parent = outFile.getParent();
        String name = outFile.getName();
        Font font = new Font("Arial", Font.PLAIN, 14);
        BufferedImage tempImg = new BufferedImage(1, 1, BufferedImage.TYPE_INT_RGB);
        Graphics2D tempG2d = tempImg.createGraphics();
        tempG2d.setFont(font);
        FontMetrics metrics = tempG2d.getFontMetrics();

        int maxCellWidth = 0;
        int maxCellHeight = metrics.getHeight() + 10; // padding

        for (int row = 0; row < rows; row++) {
            for (int col = 0; col < cols; col++) {
                Cell cell = sheet.getCell(col, row);
                String text = cell.getContents();
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

        for (int row = 0; row < rows; row++) {
            for (int col = 0; col < cols; col++) {
                int x = col * cellWidth;
                int y = row * cellHeight;
                g2d.drawRect(x, y, cellWidth, cellHeight);

                Cell cell = sheet.getCell(col, row);
                String text = cell.getContents();
                g2d.drawString(text, x + 5, y + metrics.getAscent() + 5);
            }
        }

        g2d.dispose();
        ImageIO.write(image, ext, new File(newFilePath));
        workbook.close();
        System.out.println("Excel rendered as image: " + newFilePath);
    }

    //apache POi



}
