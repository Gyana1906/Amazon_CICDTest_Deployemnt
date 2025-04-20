package AMZEcommerceTest;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

public class AMZONReviewTest {

    WebDriver driver;

    @Test
    public void landonpage() throws InterruptedException, IOException {
        WebDriverManager.chromedriver().setup();
        driver = new ChromeDriver();
        driver.get("https://www.amazon.in/");
        driver.manage().window().maximize();
        Thread.sleep(10000);
        System.out.println("Lauched AMZ App in Chrome browser");
        //Xiaomi 108 cm (43 inches) A Pro 4K
        //Xiaomi 108 cm (43 inches) X Pro QLED Series Smart Google TV L43MA-SIN (Black)
        //Mi TV 43 X Pro
        WebDriverWait wait = new WebDriverWait(driver,2000);

        driver.findElement(By.id("twotabsearchtextbox")).sendKeys("Xiaomi 108 cm (43 inches) A Pro 4K");
        driver.findElement(By.id("nav-search-submit-button")).click();

        wait.until(ExpectedConditions.elementToBeClickable(By.linkText("MI Xiaomi 108 cm (43 inches) A Pro 4K Dolby Vision Smart Google LED TV L43MA-AUIN (Black)"))).click();

        Set<String> windows = driver.getWindowHandles();
        Iterator<String> it = windows.iterator();
        String Parent = it.next();
        String Child = it.next();
        driver.switchTo().window(Child);

        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("window.scrollBy(0,14900);");
        Thread.sleep(1000);

        WebElement seemore = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("//a[contains(text(),'See more reviews')]")));
        seemore.click();
        Thread.sleep(2000);

        driver.findElement(By.id("ap_email_login")).sendKeys("xiaomiqa12@gmail.com");
        driver.findElement(By.id("continue")).click();
        driver.findElement(By.id("ap_password")).sendKeys("Chair@1234");
        driver.findElement(By.id("auth-signin-button")).click();
        Thread.sleep(2000);

        File directory = new File(System.getProperty("user.dir") + File.separator + "Output");
        if (!directory.exists()) {
            directory.mkdir();
        }

        String path = directory + File.separator + "AMZReview43APro.xlsx";
        FileOutputStream fis = new FileOutputStream(path);
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Data");
        Row row = sheet.createRow(0);

        row.createCell(0).setCellValue("Review");
        row.createCell(1).setCellValue("Models");
        row.createCell(2).setCellValue("Ratings");
        row.createCell(3).setCellValue("CustomerID");

        int secondrow10 = 1;

        while (true) {

            List<WebElement> reviewdetails = driver.findElements(By.className("review-text-content"));
            List<WebElement> Modeldetails = driver.findElements(By.cssSelector(".a-size-mini.a-link-normal.a-color-secondary"));
            List<WebElement> Ratingdetails = driver.findElements(By.xpath("//div/h5/a/i/span[@class='a-icon-alt']"));
            WebElement Section = driver.findElement(By.id("cm_cr-review_list"));
            List<WebElement> Custoemrid = Section.findElements(By.className("a-profile-name"));

            for (int j = 0; j <= reviewdetails.size() - 1; j++) {


                Row row10 = sheet.createRow(secondrow10);

                String text = reviewdetails.get(j).getText();
                row10.createCell(0).setCellValue(text);


                String text1 = Modeldetails.get(j).getText();
                row10.createCell(1).setCellValue(text1);


                WebElement text2 = Ratingdetails.get(j);
                String text3 = null;
                if (text2.getText().isEmpty()) {
                    text3 = text2.getAttribute("textContent");
                }
                row10.createCell(2).setCellValue(text3);


//                String text4 = Custoemrid.get(j).getText();
//                row10.createCell(3).setCellValue(text4);
//                System.out.println(text4 +"***"+text4.length());

                HashMap<Integer, String> li = new HashMap<>();
                int p = 0;


                String text4 = Custoemrid.get(j).getText();
                if (text4.length() > 0) {
                    li.put(p, text4);
                    p++;
                }
                for (Map.Entry<Integer,String> li1:li.entrySet()){
                    row10.createCell(3).setCellValue(li1.getValue());

                }
                secondrow10++;
            } //System.out.println(Custoemrid.size());

            // Handle Next Page
            if (driver.findElements(By.xpath("//a[text()='Next page']")).size() > 0) {
                WebElement nextPage = driver.findElement(By.xpath("//a[text()='Next page']"));
                if (nextPage.isEnabled()) {
                    nextPage.click();
                    Thread.sleep(2000);
                }

            } else {
                System.out.println("Next page button not found.");
                break;
            }
        }

        wb.write(fis);
        fis.flush();
        fis.close();
        driver.quit();
    }

}

