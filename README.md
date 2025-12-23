# Set-Out Of Office 🏖️

Are you an IT admin frustrated with setting up Out of Office notifications one mailbox at a time? 😤
Want a tool that lets you set Out of Office notifications in bulk across multiple users?
Here it is — your bulk Out of Office automation tool 🚀
Configure automatic replies for multiple mailboxes in one go using PowerShell and Microsoft Graph.

![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-blue)
![License](https://img.shields.io/badge/license-MIT-green)
![Microsoft Graph](https://img.shields.io/badge/Microsoft%20Graph-API-orange)

![GIF](https://media.giphy.com/media/v1.Y2lkPTc5MGI3NjExMGRidXowOHgwaTd5bmtjZmVmeHkxYnY4N2twNGVpcmEyMmRpMGRtMiZlcD12MV9naWZzX3NlYXJjaCZjdD1n/PCvSmNGvDmCJlo5jyG/giphy.gif)

## Features

- ✅ **Bulk Configuration** - Set OOO for multiple mailboxes simultaneously
- 🔒 **Safe Testing** - Tests with the first mailbox before applying to all
- 📅 **Flexible Scheduling** - Custom start and end dates with time zones
- 💬 **Multi-line Messages** - Support for formatted automatic reply messages
- ✨ **Interactive UI** - User-friendly prompts with visual feedback
- 🔍 **Email Validation** - Automatic cleaning and validation of email addresses
- 📊 **Progress Tracking** - Real-time feedback on success/failure status

## Prerequisites

- **PowerShell 5.1** or higher
- **Microsoft 365** tenant with Exchange Online
- **Administrator permissions** for mailbox settings
- **Required PowerShell Modules** (auto-installed by script):
  - Microsoft.Graph.Users.Actions
  - Microsoft.Graph.Authentication
  - ExchangeOnlineManagement

## Installation

### Option 1: Clone the Repository
```powershell
git clone https://github.com/yourusername/Set-OutOfOffice.git
cd Set-OutOfOffice
```

### Option 2: Download Script Directly
Download `Set-OutOfOffice.ps1` and save it to your preferred location.

## Usage

### Running the Script

1. **Open PowerShell** as Administrator (recommended)
2. **Navigate to the script directory**
   ```powershell
   cd C:\Path\To\Set-OutOfOffice
   ```
3. **Run the script**
   ```powershell
   .\Set-OutOfOffice.ps1
   ```

![Script Launch](./assets/01-launch.gif)

### Step-by-Step Guide

#### Step 1: Authentication
The script will automatically connect to Microsoft Graph and prompt for authentication.

![Authentication](./assets/02-authentication.gif)

**Required Permissions:**
- MailboxSettings.ReadWrite
- User.ReadWrite.All

#### Step 2: Enter Mailboxes
Paste your list of email addresses (one per line or separated by spaces/commas). Press **Enter twice** when finished.

```
user1@company.com
user2@company.com
user3@company.com

```

![Mailbox Input](./assets/03-mailbox-input.gif)

**Supported Formats:**
- One email per line
- Comma-separated
- Space-separated
- Mixed formats (script automatically cleans input)

#### Step 3: Set Schedule
Enter start and end dates/times for the Out of Office period.

**Format:** `DD/MM/YYYY HH:MM AM/PM`

**Example:**
```
Start: 23/12/2024 05:00 PM
End:   02/01/2025 09:00 AM
```

![Schedule Configuration](./assets/04-schedule.gif)

#### Step 4: Compose OOO Message
Enter your automatic reply message. Press **Enter twice** when finished.

**Example:**
```
Thank you for your email.

I am currently out of the office and will respond to your message when I return on January 2nd, 2025.

For urgent matters, please contact support@company.com.

Best regards

```

![Message Composition](./assets/05-message.gif)

**Tips:**
- Use blank lines for formatting
- HTML formatting is preserved
- Same message applies to internal and external recipients

#### Step 5: Review and Confirm
The script displays a summary of your configuration:

![Configuration Summary](./assets/06-summary.gif)

- Number of mailboxes
- Schedule details
- Duration calculation
- Message preview

Type **Y** to proceed or **N** to cancel.

#### Step 6: Test Mailbox
The script processes the **first mailbox** as a test:

![Test Mailbox](./assets/07-test.gif)

**Verify** the OOO settings in Outlook for this mailbox, then confirm to proceed with remaining mailboxes.

#### Step 7: Bulk Processing
Once confirmed, the script processes all remaining mailboxes:

![Bulk Processing](./assets/08-processing.gif)

**Progress Indicators:**
- ✓ Green checkmarks for successful configurations
- ✗ Red crosses for failures with error details

## Examples

### Example 1: Single Day Off
```
Start: 24/12/2024 12:00 AM
End:   24/12/2024 11:59 PM
Message: "I am out of office today. I will respond tomorrow."
```

### Example 2: Extended Leave
```
Start: 20/12/2024 05:00 PM
End:   06/01/2025 09:00 AM
Message: "I am on extended leave until January 6th. For urgent matters, contact my manager at manager@company.com."
```

### Example 3: Conference/Training
```
Start: 15/01/2025 08:00 AM
End:   17/01/2025 05:00 PM
Message: "I am attending a conference and may have limited email access. I will respond to your email as soon as possible."
```

## Troubleshooting

### Authentication Errors
**Issue:** Failed to connect to Microsoft Graph

**Solutions:**
- Ensure you have Global Admin or Exchange Admin privileges
- Check your internet connection
- Try disconnecting and reconnecting: `Disconnect-MgGraph`

### Permission Errors
**Issue:** Insufficient privileges to access mailbox settings

**Solutions:**
- Verify you have MailboxSettings.ReadWrite permissions
- Check Azure AD role assignments
- Ensure the target mailboxes are in your administrative scope

### Module Installation Errors
**Issue:** Unable to install required modules

**Solutions:**
- Run PowerShell as Administrator
- Set execution policy: `Set-ExecutionPolicy RemoteSigned -Scope CurrentUser`
- Manually install modules:
  ```powershell
  Install-Module Microsoft.Graph.Users.Actions -Scope CurrentUser
  Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
  Install-Module ExchangeOnlineManagement -Scope CurrentUser
  ```

### Date Format Errors
**Issue:** Invalid date format entered

**Solution:** Use exactly this format: `DD/MM/YYYY HH:MM AM/PM`
- Example: `25/12/2024 09:00 AM` ✓
- Not: `12/25/2024` or `25-12-2024` ✗

## Technical Details

### How It Works

1. **Module Check** - Verifies and installs required PowerShell modules
2. **Authentication** - Connects to Microsoft Graph with appropriate scopes
3. **Input Validation** - Cleans and validates email addresses using regex
4. **Date Parsing** - Converts user input to UTC for API compatibility
5. **Graph API Call** - Uses `Update-MgUserMailboxSetting` cmdlet
6. **Error Handling** - Captures and displays detailed error messages

### API Endpoint
The script uses the Microsoft Graph API endpoint:
```
PATCH /users/{userId}/mailboxSettings
```

### Data Format
Automatic replies are configured in JSON format:
```json
{
  "status": "scheduled",
  "scheduledStartDateTime": {
    "dateTime": "2024-12-23T09:00:00.000Z",
    "timeZone": "UTC"
  },
  "scheduledEndDateTime": {
    "dateTime": "2025-01-02T01:00:00.000Z",
    "timeZone": "UTC"
  },
  "internalReplyMessage": "Your message here",
  "externalReplyMessage": "Your message here"
}
```

## Security Considerations

- Script requires **administrative consent** for Graph API permissions
- **Credentials are never stored** - handled by Microsoft.Graph.Authentication
- **Audit logs** are maintained in Microsoft 365 Admin Center
- Consider using **Azure AD Privileged Identity Management** for just-in-time access

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- Built with [Microsoft Graph PowerShell SDK](https://github.com/microsoftgraph/msgraph-sdk-powershell)
- Inspired by the need for efficient bulk mailbox management

## Support

If you encounter any issues or have questions:
- 🐛 [Open an issue](https://github.com/yourusername/Set-OutOfOffice/issues)
- 💬 [Start a discussion](https://github.com/yourusername/Set-OutOfOffice/discussions)

## Changelog

### Version 1.0.0 (Initial Release)
- Interactive mailbox configuration
- Scheduled OOO with custom date ranges
- Test-first approach for safety
- Multi-line message support
- Comprehensive error handling

---

**Made with ❤️ for IT Administrators**
