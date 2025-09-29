package util;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;

public class ScreenshotUtil {

    public static void captureSS(String excelFilePath, String outputFilePath) throws AWTException, IOException, InterruptedException {

        Desktop.getDesktop().open(new File(excelFilePath));
        Thread.sleep(5000);
        String timestamp = new java.text.SimpleDateFormat("yyyyMMdd_HHmmss").format(new java.util.Date());
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

}
