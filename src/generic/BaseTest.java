package generic;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
public class BaseTest implements IAutoConstant{
	public WebDriver driver;
	
	static{
		System.setProperty(GEKCO_KEY, GEKCO_PATH);
		System.setProperty(CHROME_KEY, CHROME_PATH);
	}
	@BeforeMethod
	public void openApplication(){
		
		driver = new FirefoxDriver();
		String url = Lib.getPropertyValue("URL");
		driver.get(url);
		String ito = Lib.getPropertyValue("ITO");
		int implicitTimePeriod = Integer.parseInt(ito);
		driver.manage().timeouts().implicitlyWait(implicitTimePeriod, TimeUnit.SECONDS);
		
	}
	@AfterMethod
	public void closeApplication(){
		driver.close();
	}

}
