package createPPT;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import java.util.concurrent.TimeUnit;

import javax.imageio.ImageIO;

import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.Point;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.io.FileHandler;

public class Activity {

	
	public static WebDriver driver;

	public  void logIn () throws IOException{
		
		FileInputStream input = new FileInputStream("C:\\Users\\10638193\\workspace\\HackathonProject\\resources\\config.properties");
		
		Properties config = new Properties();
		config.load(input);
		System.setProperty("webdriver.chrome.driver", config.getProperty("webdriver.chrome.driver"));
		driver = new ChromeDriver();

		driver.manage().window().maximize();

		driver.get("https://insurancer.ap.qlikcloud.com/explore/spaces/all");
		driver.manage().timeouts().implicitlyWait(40, TimeUnit.SECONDS);		
		driver.findElement(By.name("email")).clear();
		driver.findElement(By.name("email")).sendKeys(config.getProperty("username"));

		driver.findElement(By.name("password")).clear();
		driver.findElement(By.name("password")).sendKeys(config.getProperty("password"));;
		//pswd.sendKeys(config.getProperty("password"));
		driver.findElement(By.name("submit")).click();
		
		driver.findElement(By.xpath("//body/div/div/div/div/div/div/div[2]/div[1]/div[2]/div[1]/div[1]")).click();
		
		driver.findElement(By.xpath("//body/div/div/div/div/ul/li[3]")).click();

		// WebDriverWait wait = new WebDriverWait(driver,10);
		// WebElement VARName=
		// wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("name")));

		//driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);		
		
		shootWebElement (driver.findElement(By.xpath("//body/div/div/div/div/div/div/div/ul/li[1]/div[1]")));
		
		
		//shootWebElement(driver.findElement(By.xpath("//div[@id='panel_resizable_0_4']")));

		//driver.quit();
	}
	
	public  void shootWebElement(WebElement element) throws IOException {

		File screen = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);

//		File ft = new File("C:\\Users\\10638193\\workspace\\JavaBasic\\one.png");
//		FileHandler.copy(screen, ft);

		Point p = element.getLocation();

		int width = element.getSize().getWidth();
		int height = element.getSize().getHeight();

		//System.out.println("Class Name: " + element.getSize().getClass());

		BufferedImage img = ImageIO.read(screen);

		BufferedImage dest = img.getSubimage(p.getX(), p.getY(), width, height);

		ImageIO.write(dest, "png", screen);

		File f = new File("D:\\Image\\img.png");

		FileHandler.copy(screen, f);

	}
}
