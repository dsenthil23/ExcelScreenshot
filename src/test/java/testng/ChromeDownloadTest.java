package testng;

import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.FindBy;
import org.openqa.selenium.support.PageFactory;
import org.testng.annotations.AfterTest;
import org.testng.annotations.Test;

import java.io.File;

public class ChromeDownloadTest {

    WebDriver driver;

    @FindBy(xpath = "//button[text()='Download Excel']")
    private WebElement dlButton;

    @FindBy(id="file-link")
    private WebElement downloadList;


    @Test
    public void chromeTest() throws InterruptedException {

        File file = new File("C:\\Users\\SENTHIL\\Desktop\\Index.html");
        String filePath = file.getAbsolutePath();
        driver = new ChromeDriver();
        driver.get("file:///" + filePath.replace("\\", "/"));
        driver.manage().window().maximize();
        PageFactory.initElements(driver, this);
        dlButton.click();
        driver.navigate().to("chrome://downloads");
        Thread.sleep(2000);
        JavascriptExecutor js = (JavascriptExecutor) driver;

        // Get the first file's title text from shadow DOM
        String title = (String) js.executeScript(
                "return document.querySelector('downloads-manager')" +
                        "  .shadowRoot.querySelector('#downloadsList downloads-item')" +
                        "  .shadowRoot.querySelector('a#file-link').textContent.trim();"
        );

        System.out.println("First downloaded file title: " + title);
    }

    @AfterTest
    public void tearDown()
    {
       driver.quit();
    }
}
