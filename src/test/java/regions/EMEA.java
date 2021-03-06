package regions;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.Test;

import util.Docx;

public class EMEA {
	
	
	@Test
	public void activeEMEAUsers() throws Exception {

		System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir")+"\\src\\test\\resources\\chromedriver.exe");
		WebDriver driver=new ChromeDriver();
		driver.manage().window().maximize();
		driver.get("https://www.amazon.in/");


		WebElement elm = driver.findElement(By.xpath("//div[contains(text(),'Get to Know Us')]"));
		JavascriptExecutor js=(JavascriptExecutor)driver;    	
		js.executeScript("arguments[0].scrollIntoView();", elm);
		
		
		Docx.run.setText("amazon- Get to know us");
		Docx.run.setFontSize(33);
		Docx.captureScreenShot( Docx.run, Docx.out, Docx.dirPath);
		
		driver.quit();
	}

}
