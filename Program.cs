using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

DotNetEnv.Env.Load("C:\\repos\\AutoApex\\AutoApexImport\\.env");

var userName = Environment.GetEnvironmentVariable("APEX_USERNAME");
var password = Environment.GetEnvironmentVariable("APEX_PASSWORD");

// Load the selenium web driver
IWebDriver driver = new ChromeDriver();

void Login()
{
    try
    {
        //Go to the Apex login page
        driver.Navigate().GoToUrl("https://apexconnectandgo.com");
        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
        //Find the username and password fields and enter the credentials
        var usernameField = driver.FindElement(By.Id("user.login_id"));
        usernameField.SendKeys(userName);
        var passwordField = driver.FindElement(By.Id("user.password"));
        passwordField.SendKeys(password);
        //Login
        passwordField.Submit();
        

    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        throw;
    }
}
Login();
Console.ReadLine();

