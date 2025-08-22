using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;

//Set the license for the library
ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
//Load the env file
DotNetEnv.Env.Load("C:\\repos\\AutoApex\\AutoApexImport\\.env");

var excelPath = Environment.GetEnvironmentVariable("EXCEL_PATH");
var userName = Environment.GetEnvironmentVariable("APEX_USERNAME");
var password = Environment.GetEnvironmentVariable("APEX_PASSWORD");


var options = new ChromeOptions();
options.AddArgument("--log-level=3"); //Only fatal errors displayed in the terminal
options.AddArgument("--silent");

var service = ChromeDriverService.CreateDefaultService();
service.SuppressInitialDiagnosticInformation = true;
service.EnableVerboseLogging = false;

// Load the selenium web driver with options
IWebDriver driver = new ChromeDriver(service, options);

void Login()
{
    try
    {
        //Maximize the window
        driver.Manage().Window.Maximize();
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
        goToProfileManager();
    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        throw;
    }
}

//Method to navigate to the Profile Manager Page
void goToProfileManager()
{
    var profileManager = driver.FindElement(By.CssSelector("#logout_left > a:nth-child(1)"));
    Console.WriteLine("Navigating to the profile manager");
    profileManager.Click();

    var manageUsers = driver.FindElement(By.CssSelector("#pageBody > div.drawers-wrapper > ul > li:nth-child(1) > ul > li:nth-child(1) > a"));
    Console.WriteLine("Clicking on 'Manage Users'");
    manageUsers.Click();
}

/*using (var package = new ExcelPackage(new FileInfo(excelPath)))
{
    var worksheet = package.Workbook.Worksheets[0];
    Console.WriteLine(worksheet.FirstValueCell);
}
*/

Login();
Console.ReadLine();
