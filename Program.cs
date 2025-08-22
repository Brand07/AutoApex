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
options.AddArgument("--log-level=4"); //Only fatal errors displayed in the terminal
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
        GoToProfileManager();
    }
    catch (Exception e)
    {
        Console.WriteLine(e);
        throw;
    }
}

//Method to navigate to the Profile Manager Page
void GoToProfileManager()
{
    var profileManager = driver.FindElement(By.CssSelector("#logout_left > a:nth-child(1)"));
    Console.WriteLine("Navigating to the profile manager");
    profileManager.Click();

    var manageUsers = driver.FindElement(By.CssSelector("#pageBody > div.drawers-wrapper > ul > li:nth-child(1) > ul > li:nth-child(1) > a"));
    Console.WriteLine("Clicking on 'Manage Users'");
    manageUsers.Click();
}

void DoesBadgeExist(string badgeNumber)
{
    try
    {
        var badgeElement = driver.FindElement(By.XPath($"//td[contains(text(), '{badgeNumber}')]"));
        if (badgeElement != null)
        {
            Console.WriteLine($"Badge {badgeNumber} exists.");
        }
    }
    catch (NoSuchElementException)
    {
        Console.WriteLine($"Badge {badgeNumber} does not exist.");
    }
}

void SearchBadge(string badgeNumber)
{
    var searchBox = driver.FindElement(By.Id("searchUsersText"));
    searchBox.Click();
    searchBox.SendKeys(badgeNumber);
    //Click on the search button
    var searchButton = driver.FindElement(By.CssSelector("#searchAddUser2 > button"));
    searchButton.Click();
    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
    //Check if the badge exists
    DoesBadgeExist(badgeNumber);
    //Clear the search box
    searchBox.Clear();
    
    
}

Login();

using (var package = new ExcelPackage(new FileInfo(excelPath)))
{
    var worksheet = package.Workbook.Worksheets[0];
    int rowCount = worksheet.Dimension.Rows;
    int colCount = worksheet.Dimension.Columns;
    
    //Map column names to indices
    var colMap = new Dictionary<string, int>();
    for (int col = 1; col <= colCount; col++)
    {
        var colName = worksheet.Cells[1, col].Text.Trim();
        colMap[colName] = col;
    }

    for (int row = 2; row <= rowCount; row++)
    {
        string firstName = worksheet.Cells[row, colMap["First Name"]].Text;
        string lastName = worksheet.Cells[row, colMap["Last Name"]].Text;
        string employeeId = worksheet.Cells[row, colMap["Badge Number"]].Text;
        int badgeNum = int.Parse(employeeId);
        string department = worksheet.Cells[row, colMap["Department"]].Text;

        Console.WriteLine("Searching for the badge association.");
        SearchBadge(badgeNum.ToString());

    }
}



Console.ReadLine();
