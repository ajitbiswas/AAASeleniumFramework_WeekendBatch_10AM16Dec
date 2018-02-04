package scripts;

import org.testng.annotations.Test;

import generic.BaseTest;
import generic.Lib;
import pompages.LoginPage;

public class TestInvalidLogin extends BaseTest{
	@Test
	public void testInvalidLogin() throws InterruptedException{
		
		LoginPage l = new LoginPage(driver);
		int rowCount = Lib.getRowCount("InvalidLogin");
		for (int i = 1; i <= rowCount; i++) {
			String username = Lib.getCellValue("InvalidLogin", i, 0);
			l.setUsername(username);
			Thread.sleep(2000);
			String password = Lib.getCellValue("InvalidLogin", i, 1);
			l.setPassword(password);
			Thread.sleep(2000);
			l.clickLogin();
		}
		
		
		
		
	}

}
