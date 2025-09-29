package testng;

import org.testng.annotations.Test;
import util.ScreenshotUtil;

import java.awt.*;
import java.io.File;
import java.io.IOException;

public class ExcelScreenshotTest {

    @Test
    public void TestScreenShot() throws InterruptedException, IOException, AWTException {
        String excelPath = "src/test/resources/TestDataExcel.xlsx";
        String ssFile = "src/test/output/excel_screenshot.png";
        ScreenshotUtil.captureSS(excelPath, ssFile);
    }

}
