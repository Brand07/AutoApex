# AutoApex
My Apex Connect &amp; Go program rewritten in C#
The original version written in Python can be found here - https://github.com/Brand07/Python--Apex-Connect-And-Go

# AutoApexImport

`AutoApexImport` is a C# console application designed to automate the process of importing data from an Excel spreadsheet into the Apex Connect & Go web application. It uses Selenium for web browser automation to navigate the Apex application and EPPlus to read data from the Excel file.

## Prerequisites

Before you begin, ensure you have met the following requirements:
*   [.NET 9.0 SDK](https://dotnet.microsoft.com/download/dotnet/9.0) or later.
*   Google Chrome installed on your system.

## Installation

1.  Clone the repository:
    ```bash
    git clone https://github.com/Brand07/AutoApex.git
    ```
2.  Navigate to the project directory:
    ```bash
    cd AutoApexImport
    ```
3.  Restore the .NET packages:
    ```bash
    dotnet restore
    ```

## Configuration

The application uses a `.env` file to manage sensitive information and configuration settings.

1.  Create a file named `.env` in the root of the project (`C:\repos\AutoApex\AutoApexImport\.env`).
2.  Add the following environment variables to the `.env` file:

    ```
    EXCEL_PATH="path\\to\\your\\excel\\file.xlsx"
    APEX_USERNAME="your_apex_username"
    APEX_PASSWORD="your_apex_password"
    API_KEY="your_freshservice_api_key"
    DOMAIN="your_domain"
    REQUESTER_ID="your_freshservice_requester_id"
    ```

    Replace the placeholder values with your actual Excel file path and APEX credentials.

## FreshService Ticket Creation

This application can automatically create tickets in FreshService for each user processed from the Excel file. Ticket creation is controlled by a configuration flag and environment variables.

### Enabling/Disabling Ticket Creation

- Ticket creation is controlled by the `logTickets` flag in the code. Set it to `true` to enable ticket creation, or `false` to disable it.
- When enabled, a FreshService ticket will be created for each user processed.

### Additional Environment Variables

Add the following variables to your `.env` file for FreshService integration:

```
API_KEY="your_freshservice_api_key"
DOMAIN="your_freshservice_domain"  # e.g., 'yourcompany' for yourcompany.freshservice.com
REQUESTER_ID="your_freshservice_requester_id"  # Must be a valid FreshService user ID (as a number)
```

## Usage

To run the application, execute the following command from the project's root directory:

```bash
dotnet run
```

The application will then:
1.  Launch a new Chrome browser instance.
2.  Navigate to the Oracle APEX login page.
3.  Log in with the provided credentials.
4.  Read data from the specified Excel file.
5.  Automate the data import process into the APEX application.
6.  (Optional) Create a FreshService ticket for each user, if ticket logging is enabled.

## License

This project is licensed under the MIT License. See the [LICENSE](./LICENSE) file for details.

**Note:** This project uses EPPlus, which is licensed under Polyform Noncommercial 1.0.0. If you intend to use this project for commercial purposes, you must obtain a commercial license for EPPlus. See the [EPPlus License](https://epplussoftware.com/developers/licenseexception) for more information.
