using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OfficeOpenXml;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;

namespace AutoApexImport
{
    class Program
    {
        static IWebDriver? driver;
        static string? excelPath;
        static string? userName;
        static string? password;

        static void Main()
        {
            //Set the license for the library
            ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization");
            //Load the env file
            DotNetEnv.Env.Load("C:\\repos\\AutoApex\\AutoApexImport\\.env");

            excelPath = Environment.GetEnvironmentVariable("EXCEL_PATH");
            userName = Environment.GetEnvironmentVariable("APEX_USERNAME");
            password = Environment.GetEnvironmentVariable("APEX_PASSWORD");

            if (string.IsNullOrWhiteSpace(excelPath) || string.IsNullOrWhiteSpace(userName) ||
                string.IsNullOrWhiteSpace(password))
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

            var service = ChromeDriverService.CreateDefaultService();
            service.SuppressInitialDiagnosticInformation = true;
            service.EnableVerboseLogging = false;
            service.HideCommandPromptWindow = true;
            service.LogPath = "NUL"; // Suppress logs on Windows
            

            // Load the selenium web driver with options
            driver = new ChromeDriver(service, options);

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
                    SearchBadge(firstName, lastName, badgeNum.ToString(), department);
                }
            }
        }

        static void Login()
        {
            if (driver == null || userName == null || password == null)
                throw new InvalidOperationException("Driver or credentials not initialized.");
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

        static void GoToProfileManager()
        {
            if (driver == null) throw new InvalidOperationException("Driver not initialized.");
            var profileManager = driver.FindElement(By.CssSelector("#logout_left > a:nth-child(1)"));
            Console.WriteLine("Navigating to the profile manager");
            profileManager.Click();

            var manageUsers =
                driver.FindElement(
                    By.CssSelector(
                        "#pageBody > div.drawers-wrapper > ul > li:nth-child(1) > ul > li:nth-child(1) > a"));
            Console.WriteLine("Clicking on 'Manage Users'");
            manageUsers.Click();
        }

        static void WaitForOverlayToDisappear(IWebDriver driver, WebDriverWait wait)
        {
            wait.Until(drv =>
            {
                var overlays = drv.FindElements(By.CssSelector(".ui-widget-overlay.ui-front"));
                return overlays.Count == 0 || overlays.All(o => !o.Displayed);
            });
        }

        static void EditDepartment(string department)
        {
            if (driver == null) throw new InvalidOperationException("Driver not initialized.");
            Console.WriteLine("Editing the department");
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {
                var departmentLink = wait.Until(drv => {
                    var el = drv.FindElement(By.LinkText("User Group Membership:"));
                    return (el.Displayed && el.Enabled) ? el : null;
                });
                departmentLink.Click();

                WaitForOverlayToDisappear(driver, wait);

                // Wait for checkboxes to be present
                wait.Until(drv => drv.FindElements(By.CssSelector("input[type='checkbox']")).Count > 0);
                var allDepartments = driver.FindElements(By.CssSelector("input[type='checkbox']"));
                foreach (var dept in allDepartments)
                {
                    if (dept.Selected && dept.Displayed && dept.Enabled)
                    {
                        WaitForOverlayToDisappear(driver, wait);
                        dept.Click();
                        WaitForOverlayToDisappear(driver, wait);
                    }
                }
                //Select the department based on the input
                switch (department)
                {
                    case "Cycle Count":
                        var cycleCount = wait.Until(drv => {
                            var el = drv.FindElement(By.CssSelector("#editMembershipCheck2"));
                            return (el.Displayed && el.Enabled) ? el : null;
                        });
                        if (!cycleCount.Selected && cycleCount.Displayed && cycleCount.Enabled)
                        {
                            WaitForOverlayToDisappear(driver, wait);
                            cycleCount.Click();
                            WaitForOverlayToDisappear(driver, wait);
                        }
                        break;
                    case "Material Handling":
                        var materialHandling = wait.Until(drv => {
                            var el = drv.FindElement(By.CssSelector("#editMembershipCheck4"));
                            return (el.Displayed && el.Enabled) ? el : null;
                        });
                        if (!materialHandling.Selected && materialHandling.Displayed && materialHandling.Enabled)
                        {
                            WaitForOverlayToDisappear(driver, wait);
                            materialHandling.Click();
                            WaitForOverlayToDisappear(driver, wait);
                        }
                        break;
                    case "Voice Pick":
                        var voicePick = wait.Until(drv => {
                            var el = drv.FindElement(By.CssSelector("#editMembershipCheck6"));
                            return (el.Displayed && el.Enabled) ? el : null;
                        });
                        if (!voicePick.Selected && voicePick.Displayed && voicePick.Enabled)
                        {
                            WaitForOverlayToDisappear(driver, wait);
                            voicePick.Click();
                            WaitForOverlayToDisappear(driver, wait);
                        }
                        break;
                    default:
                        Console.WriteLine($"Department '{department}' not recognized. No changes made.");
                        break;
                }
            }
            catch (WebDriverTimeoutException ex)
            {
                Console.WriteLine($"Timeout waiting for element: {ex.Message}");
            }
            catch (ElementNotInteractableException ex)
            {
                Console.WriteLine($"Element not interactable: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        static void EditProfile(string firstUserName, string lastUserName, string badgeNumber, string department)
        {
            if (driver == null) throw new InvalidOperationException("Driver not initialized.");
            Console.WriteLine("Editing the profile");
            var firstNameField = driver.FindElement(By.Id("edit_user.first_name"));
            firstNameField.Clear();
            firstNameField.SendKeys(firstUserName);
            var lastNameField = driver.FindElement(By.Id("edit_user.last_name"));
            lastNameField.Clear();
            lastNameField.SendKeys(lastUserName);
            var employeeIdField = driver.FindElement(By.Id("employeeId"));
            employeeIdField.Clear();
            employeeIdField.SendKeys(badgeNumber);
            var badgeNumberField = driver.FindElement(By.Id("badgeNumber"));
            badgeNumberField.Clear();
            badgeNumberField.SendKeys(badgeNumber);
            //Edit the department
            EditDepartment(department);
            //Click on the save button
            var saveButton = driver.FindElement(By.Id("updateUser"));
            saveButton.Click();
        }

        static void DoesBadgeExist(string firstName, string lastName, string badgeNumber, string departrment)
        {
            if (driver == null) throw new InvalidOperationException("Driver not initialized.");
            try
            {
                var badgeElement = driver.FindElement(By.XPath($"//td[contains(text(), '{badgeNumber}')]"));
                Console.WriteLine($"Badge {badgeNumber} exists.");
                Console.WriteLine("Editing the current badge association.");
                var profileLink = badgeElement.FindElement(By.XPath("//*[@id=\"tr0\"]/td[1]/a"));
                profileLink.Click();
                //TODO Call method to edit the profile
                EditProfile(firstName, lastName, badgeNumber, departrment);
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine($"Badge {badgeNumber} does not exist.");
                //TODO Call method to create a new profile
            }
        }

        static void SearchBadge(string firstName, string lastName, string badgeNumber, string department)
        {
            if (driver == null) throw new InvalidOperationException("Driver not initialized.");
            var searchBox = driver.FindElement(By.Id("searchUsersText"));
            searchBox.Click();
            searchBox.SendKeys(badgeNumber);
            //Click on the search button
            var searchButton = driver.FindElement(By.CssSelector("#searchAddUser2 > button"));
            searchButton.Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMilliseconds(500);
            //Check if the badge exists
            DoesBadgeExist(firstName, lastName, badgeNumber, department);
            //Clear the search box
            searchBox.Clear();
        }

        // Helper to highlight an element using JavaScript
        static void HighlightElement(IWebElement element)
        {
            if (driver == null) return;
            var js = (IJavaScriptExecutor)driver;
            js.ExecuteScript("arguments[0].style.border='3px solid red'", element);
        }
    }
}
