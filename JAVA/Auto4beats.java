import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.DayOfWeek;
import java.time.Duration;
import java.time.LocalDate;
import java.util.List;

public class Auto4beats {
    public static void main(String[] args) {
        System.setProperty("webdriver.chrome.driver", "C:\\ProgramData\\chocolatey\\lib\\chromedriver\\tools\\chromedriver-win32\\chromedriver.exe");//give path to chromedriver

        WebDriver driver = new ChromeDriver();

        try {
            String excelFile = "C:\\Users\\ssowm\\Desktop\\f1\\4beats\\src\\test\\java\\4BeatsQ1.xlsx";//give path to excel file
            FileInputStream file = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet(getCurrentDayOfWeek());

            int columnIndex = 3;
            String keyword = null;
            Row row = null;
            Cell cell;
            for (int rowindex = 2; rowindex < 12; rowindex++) {
                row = sheet.getRow(rowindex);
                cell = row.getCell(columnIndex - 1);
                keyword = cell.getStringCellValue();

                driver.get("https://www.google.com");

                WebElement searchInput = driver.findElement(By.name("q"));
                searchInput.sendKeys(keyword);

                WebDriverWait wait = new WebDriverWait(driver, Duration.ofSeconds(10));
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector("ul[role='listbox']")));

                WebElement suggestionsContainer = driver.findElement(By.cssSelector("ul[role='listbox']"));
                List<WebElement> suggestionElements = suggestionsContainer.findElements(By.tagName("li"));

                String longestOption = "";
                String shortestOption = "";

                for (WebElement suggestionElement : suggestionElements) {
                    String suggestion = suggestionElement.getText();
                    if (suggestion.length() > longestOption.length()) {
                        longestOption = suggestion;
                    }
                    if (shortestOption.isEmpty() || suggestion.length() < shortestOption.length()) {
                        shortestOption = suggestion;
                    }
                }
                Cell longestOptionCell = row.createCell(3);
                longestOptionCell.setCellValue(longestOption);

                Cell shortestOptionCell = row.createCell(4);
                shortestOptionCell.setCellValue(shortestOption);

                FileOutputStream fileOut = new FileOutputStream(excelFile);
                workbook.write(fileOut);
                fileOut.close();

            }

            driver.quit();
        }
        catch (IOException e) {
            e.printStackTrace();
        }

    }
    private static String getCurrentDayOfWeek() {
        DayOfWeek dayOfWeek = LocalDate.now().getDayOfWeek();
        return dayOfWeek.name();
    }
}

