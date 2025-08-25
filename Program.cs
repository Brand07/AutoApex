using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
//Libraries used for the FreshService API
using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;


namespace AutoApexImport
{
    public class FreshServiceTicketCreator
    {
        private readonly HttpClient _httpClient;
        private readonly bool _logTickets;

        public FreshServiceTicketCreator(bool logTickets)
        {
            _logTickets = logTickets;
            var apiKey = Environment.GetEnvironmentVariable("API_KEY");
            var domain = Environment.GetEnvironmentVariable("DOMAIN");
            
            
            _httpClient = new HttpClient();
            var byteArray = Encoding.ASCII.GetBytes($"{apiKey}:X");
            _httpClient.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Basic", Convert.ToBase64String(byteArray));
            _httpClient.BaseAddress = new Uri($"https://{domain}.freshservice.com/api/v2/");

        }

        public async Task<string> CreateTicketAsync(string subject, string description, string service,
            long requesterId)
        {
            if (!_logTickets)
            {
                return "Ticket logging is disabled.";
            }
            var ticketData = new
            {
                priority = 2,
                status = 2, // 2 = Open
                requester_id = requesterId,
                custom_fields = new
                {
                    please_select_the_service = service
                }
            };

            var json = JsonConvert.SerializeObject(ticketData);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _httpClient.PostAsync("tickets", content);
            if (!response.IsSuccessStatusCode)
            {
                var errorBody = await response.Content.ReadAsStringAsync();
                throw new HttpRequestException($"FreshService API error: {response.StatusCode}\n{errorBody}");
            }
            return await response.Content.ReadAsStringAsync();
        }
    }
    
    class Program
    {
        static IWebDriver? _driver;
        static string? _excelPath;
        static string? _userName;
        static string? _password;
        
        //Enable FreshService ticket creation?
        //private bool enableTickets = false; //Change to true to input a ticket for each user edited/added.

        static async Task Main()
        {
            //Set the license for the library
            ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
            //Load the env file
            DotNetEnv.Env.Load("C:\\repos\\AutoApex\\AutoApexImport\\.env");

            _excelPath = Environment.GetEnvironmentVariable("EXCEL_PATH");
            _userName = Environment.GetEnvironmentVariable("APEX_USERNAME");
            _password = Environment.GetEnvironmentVariable("APEX_PASSWORD");

            if (string.IsNullOrWhiteSpace(_excelPath) || string.IsNullOrWhiteSpace(_userName) ||
                string.IsNullOrWhiteSpace(_password))
            {
                Console.WriteLine("One or more required environment variables are missing.");
                return;
            }

            var options = new ChromeOptions();
            options.AddArgument("--log-level=4"); //Only fatal errors displayed in the terminal
            options.AddArgument("--silent");
            // Disable Chrome autofill, password manager, and save prompts
            options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
            options.AddUserProfilePreference("profile.password_manager_enabled", false);
            options.AddUserProfilePreference("credentials_enable_service", false);
            options.AddUserProfilePreference("autofill.profile_enabled", false);
            options.AddUserProfilePreference("autofill.address_enabled", false);
            options.AddUserProfilePreference("autofill.credit_card_enabled", false);

            var webService = ChromeDriverService.CreateDefaultService();
            webService.SuppressInitialDiagnosticInformation = true;
            webService.EnableVerboseLogging = false;
            webService.HideCommandPromptWindow = true;
            webService.LogPath = "NUL"; // Suppress logs on Windows
            

            // Load the selenium web driver with options
            _driver = new ChromeDriver(webService, options);

            Login();

            using var package = new ExcelPackage(new FileInfo(_excelPath));
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

            var ticketCreator = new FreshServiceTicketCreator(false);//Change to 'true' to enable ticket creation

                

            for (int row = 2; row <= rowCount; row++)
            {
                string firstName = worksheet.Cells[row, colMap["First Name"]].Text;
                string lastName = worksheet.Cells[row, colMap["Last Name"]].Text;
                string employeeId = worksheet.Cells[row, colMap["Badge Number"]].Text;
                int badgeNum = int.Parse(employeeId);
                string department = worksheet.Cells[row, colMap["Department"]].Text;

                Console.WriteLine("Searching for the badge association.");
                SearchBadge(firstName, lastName, badgeNum.ToString(), department);

                try
                {
                    string subject = $"{firstName} {lastName} needs access to {department} devices.";
                    string description = $"User's badge number is {badgeNum}";
                    string service = "Apex";
                    long requesterId = long.Parse(Environment.GetEnvironmentVariable("REQUESTER_ID") ?? "0");

                    string response =
                        await ticketCreator.CreateTicketAsync(subject, description, service, requesterId);
                    Console.WriteLine($"Ticket created for {firstName} {lastName}");
                    Console.WriteLine(response);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create ticket for {firstName} {lastName}.\n{ex.Message}");
                    throw;
                }
            }
        }

        static void Login()
        {
            if (_driver == null || _userName == null || _password == null)
                throw new InvalidOperationException("Driver or credentials not initialized.");
            try
            {
                //Maximize the window
                _driver.Manage().Window.Maximize();
                //Go to the Apex login page
                _driver.Navigate().GoToUrl("https://apexconnectandgo.com");
                _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
                
                
                //Find the username and password fields and enter the credentials
                var usernameField = _driver.FindElement(By.Id("user.login_id"));
                usernameField.SendKeys(_userName);
                
                var passwordField = _driver.FindElement(By.Id("user.password"));
                passwordField.SendKeys(_password);
                
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

        static void GoToProfileManager()
        {
            if (_driver == null) throw new InvalidOperationException("Driver not initialized.");
            var profileManager = _driver.FindElement(By.CssSelector("#logout_left > a:nth-child(1)"));
            Console.WriteLine("Navigating to the profile manager");
            profileManager.Click();

            var manageUsers =
                _driver.FindElement(
                    By.CssSelector(
                        "#pageBody > div.drawers-wrapper > ul > li:nth-child(1) > ul > li:nth-child(1) > a"));
            Console.WriteLine("Clicking on 'Manage Users'");
            manageUsers.Click();
        }
        

        static void EditDepartment(string department)
        {
            //Departments and their IDs
            var cycleCount = _driver.FindElement(By.Id("editMembershipCheck2"));
            var materialHandling = _driver.FindElement(By.Id("editMembershipCheck4"));
            var sort = _driver.FindElement(By.Id("editMembershipCheck5"));
            var pick = _driver.FindElement(By.Id("editMembershipCheck6"));
            
            //Uncheck all the boxes if checked
            if (cycleCount.Selected) cycleCount.Click();
            if (materialHandling.Selected) materialHandling.Click();
            if (sort.Selected) sort.Click();
            if (pick.Selected) pick.Click();
            
            
            //Click on the department checkbox based on the passed department
            switch (department)
            {
                case "Cycle Count":
                    cycleCount.Click();
                    break;
                case "Material Handler":
                    materialHandling.Click();
                    break;
                case "Sort":
                    sort.Click();
                    break;
                case "Voice Pick":
                    pick.Click();
                    break;
                default:
                    Console.WriteLine($"Unknown department: {department}");
                    break;
            }
            //Click the save button
            var saveButton = _driver.FindElement(By.CssSelector(
                "body > div:nth-child(21) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button:nth-child(2)"));
            Console.WriteLine("Clicking the save button.");
            saveButton.Click();
            
        }

        static void AddDepartment(string department)
        {
            var cycleCount = _driver.FindElement(By.Id("membershipCheck2"));
            var materialHandling = _driver.FindElement(By.Id("membershipCheck4"));
            var sort = _driver.FindElement(By.Id("membershipCheck5"));
            var pick = _driver.FindElement(By.Id("membershipCheck6"));
            
            //Uncheck all the boxes if checked
            if (cycleCount.Selected) cycleCount.Click();
            if (materialHandling.Selected) materialHandling.Click();
            if (sort.Selected) sort.Click();
            if (pick.Selected) pick.Click();
            
            switch (department)
            {
                case "Cycle Count":
                    cycleCount.Click();
                    break;
                case "Material Handler":
                    materialHandling.Click();
                    break;
                case "Sort":
                    sort.Click();
                    break;
                case "Voice Pick":
                    pick.Click();
                    break;
                default:
                    Console.WriteLine($"Unknown department: {department}");
                    break;
            }
            //Click the add button
            Thread.Sleep(2000);
            var addButton = _driver.FindElement(By.XPath(
                "/html/body/div[22]/div[3]/div/button[2]"));
            addButton.Click();
            Console.WriteLine("Add Button Clicked.");
            
            //Pause for 1 second
            Thread.Sleep(1000);
            
            //Click on the submit button to officially add the user
            var submitButton = _driver.FindElement(By.CssSelector("#apexTblDisplay > tbody > tr > td > button"));
            submitButton.Click();
            Thread.Sleep(1000);
            
            //Click 'OK' on the proceeding dialogue box.
            var confirmButton = _driver.FindElement(By.CssSelector(
                "body > div:nth-child(4) > div.ui-dialog-buttonpane.ui-widget-content.ui-helper-clearfix > div > button"));
            confirmButton.Click();

            Console.WriteLine("User has been added.");
        }
        
        //Function to reformat the badge number 
        static void ReformatBadgeNumber(ref string badgeNumber)
        {
            // If the badge number is 4 digits long, append a zero at the front
            if (badgeNumber.Length == 4)
            {
                badgeNumber = "0" + badgeNumber;
            }
            //If the badge number is 5 digits long, do nothing
            else if (badgeNumber.Length == 5)
            {
                Console.WriteLine("Badge number is already 5 digits long.");
            }
        }

        static void EditProfile(string firstUserName, string lastUserName, string badgeNumber, string department)
        {
            if (_driver == null) throw new InvalidOperationException("Driver not initialized.");
            Console.WriteLine("Editing the profile");
            var firstNameField = _driver.FindElement(By.Id("edit_user.first_name"));
            firstNameField.Clear();
            firstNameField.SendKeys(firstUserName);
            
            var lastNameField = _driver.FindElement(By.Id("edit_user.last_name"));
            lastNameField.Clear();
            lastNameField.SendKeys(lastUserName);
            
            var employeeIdField = _driver.FindElement(By.Id("employeeId"));
            employeeIdField.Clear();
            employeeIdField.SendKeys(badgeNumber);
            
            var badgeNumberField = _driver.FindElement(By.Id("badgeNumber"));
            badgeNumberField.Clear();
            ReformatBadgeNumber(ref badgeNumber);
            badgeNumberField.SendKeys(badgeNumber);
            
            //Click the "User Group Membership" option
            Console.WriteLine("Clicking the user group membership.");
            var groupMembership = _driver.FindElement(By.LinkText("User Group Membership:"));
            groupMembership.Click();
            
            //Edit the department
            EditDepartment(department);
            //Click on the save button
            var saveButton = _driver.FindElement(By.Id("updateUser"));
            saveButton.Click();
        }

        static void AddUser(string firstName, string lastName, string badgeNumber, string department)
        {
            //Click on the "Add a User" button
            var addUserButton = _driver.FindElement(By.LinkText("Add a User"));
            addUserButton.Click();
            Console.WriteLine($"Adding user {firstName} {lastName} with badge number {badgeNumber}. ");

            var firstNameField = _driver.FindElement(By.Id("user.first_name"));
            firstNameField.SendKeys(firstName);

            var lastNameField = _driver.FindElement(By.Id("user.last_name"));
            lastNameField.SendKeys(lastName);

            var employeeIdField = _driver.FindElement(By.Id("addPassport.employee_id"));
            employeeIdField.SendKeys(badgeNumber);

            var badgeNumberField = _driver.FindElement(By.Id("addPassport.user_card_key"));
            //Reformat the badge number and send it
            ReformatBadgeNumber(ref badgeNumber);
            badgeNumberField.SendKeys(badgeNumber);

            var userMembership = _driver.FindElement(By.LinkText("User Group Membership:"));
            userMembership.Click();
            AddDepartment(department);
            //TODO - Make a function to add (not edit) the group permissions. 
            // Elements between editing vs adding have different IDs.

        }

        static void DoesBadgeExist(string firstName, string lastName, string badgeNumber, string department)
        {
            if (_driver == null) throw new InvalidOperationException("Driver not initialized.");
            try
            {
                var badgeElement = _driver.FindElement(By.XPath($"//td[contains(text(), '{badgeNumber}')]"));
                Console.WriteLine($"Badge {badgeNumber} exists.");
                Console.WriteLine("Editing the current badge association.");
                
                var profileLink = badgeElement.FindElement(By.XPath("//*[@id=\"tr0\"]/td[1]/a"));
                profileLink.Click();
                
                EditProfile(firstName, lastName, badgeNumber, department);
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"Badge {badgeNumber} does not exist - adding user to the system.");
                
                AddUser(firstName, lastName, badgeNumber, department);
                
            }
        }

        static void SearchBadge(string firstName, string lastName, string badgeNumber, string department)
        {
            if (_driver == null) throw new InvalidOperationException("Driver not initialized.");
            var searchBox = _driver.FindElement(By.Id("searchUsersText"));
            searchBox.Click();
            searchBox.SendKeys(badgeNumber);
            //Click on the search button
            var searchButton = _driver.FindElement(By.CssSelector("#searchAddUser2 > button"));
            searchButton.Click();
            
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
            //Check if the badge exists
            DoesBadgeExist(firstName, lastName, badgeNumber, department);
            //Clear the search box
            searchBox.Clear();
        }
    }
}
