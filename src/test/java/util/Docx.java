package util;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterSuite;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.Test;

public class Docx {
	public static XWPFDocument docx;
	public static XWPFRun run;
	public static FileOutputStream out;
	public static String dirPath;

	@BeforeSuite
	public void initReport() {
		try {
			dirPath = System.getProperty("user.dir")+"\\reports\\" ;

			docx = new XWPFDocument();
			run = docx.createParagraph().createRun();
			Date d = new Date();
			String DateTimeStamp=d.toString().replace(" ", "_").replace(":", "_");
			out = new FileOutputStream(dirPath+"VPN_Performance_Details_"+DateTimeStamp+".docx");  

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	
	
	@AfterSuite
	public void closeReport() {
		try {
			TimeUnit.SECONDS.sleep(1);
			System.out.println("Write to doc file sucessfully...");
			docx.write(out);
			out.flush();
			out.close();
			docx.close();
		}catch(Exception e) {
			e.printStackTrace();
		}
	}

	public static void captureScreenShot( XWPFRun run, FileOutputStream out, String dirPath) throws Exception {
		String screenshot_name = System.currentTimeMillis() + ".png";
		BufferedImage image = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
		File file = new File(dirPath + screenshot_name);
		ImageIO.write(image, "png", file);
		InputStream pic = new FileInputStream(dirPath + screenshot_name);
		run.addBreak();
		run.addPicture(pic, XWPFDocument.PICTURE_TYPE_PNG, screenshot_name, Units.toEMU(500), Units.toEMU(350));
		pic.close();
		file.delete();
	}

}