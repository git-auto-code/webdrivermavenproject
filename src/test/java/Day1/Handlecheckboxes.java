package Day1;

import io.github.bonigarcia.wdm.WebDriverManager;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

public class Handlecheckboxes {
    public static void main(String[] args) {
        WebDriverManager.chromedriver().setup();  // Automatically manages driver version
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.google.com");
    }
}

