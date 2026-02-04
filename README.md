# Outlook Email Contact Analyzer Add-in

A powerful Outlook add-in that scans your emails and creates an organized, contact-based view. This add-in analyzes emails from your inbox and groups them by contact, showing whether each contact appeared in To, CC, or BCC fields.

## Features

- **Email Scanning**: Scans up to 500 emails from your inbox
- **Contact Grouping**: Organizes emails by unique contacts
- **Recipient Type Tracking**: Shows To/CC/BCC breakdown for each contact
- **Search & Filter**: Quickly find specific contacts
- **Statistics Dashboard**: View total contacts, emails, and averages
- **Contact Details**: Click any contact to see their recent emails
- **Modern UI**: Clean, professional interface with smooth interactions

## Files Included

1. **manifest.xml** - Office Add-in manifest file
2. **taskpane.html** - Main user interface
3. **taskpane.js** - Core functionality and email processing logic
4. **commands.html** - Required for command surface integration
5. **README.md** - This file

## Setup Instructions

### Prerequisites

- Microsoft 365 subscription with Outlook Online
- A web server to host the add-in files (HTTPS required)
- Basic knowledge of web hosting

### Deployment Steps

#### Option 1: Deploy to Your Own Web Server

1. **Prepare Your Files**
   - Upload all files (manifest.xml, taskpane.html, taskpane.js, commands.html) to your HTTPS web server
   - Note your server URL (e.g., `https://yourdomain.com`)

2. **Update the Manifest**
   - Open `manifest.xml`
   - Replace ALL instances of `https://yourdomain.com` with your actual domain
   - Generate a new GUID for the `<Id>` field (you can use an online GUID generator)
   - Update `<ProviderName>` with your company name

3. **Create Icon Images** (optional but recommended)
   - Create icons: 16x16, 32x32, 64x64, 80x80, 128x128 pixels
   - Upload to your server at `/assets/icon-*.png`
   - Update icon URLs in manifest if using different paths

4. **Install the Add-in**
   - Go to Outlook on the web (outlook.office.com)
   - Click the settings gear icon → View all Outlook settings
   - Go to Mail → Customize actions → Get add-ins
   - Click "My add-ins" → "Add a custom add-in" → "Add from file"
   - Upload your `manifest.xml` file
   - Click "Install"

#### Option 2: Deploy Using GitHub Pages (Free)

1. **Create a GitHub Repository**
   ```bash
   git init
   git add taskpane.html taskpane.js commands.html
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/yourusername/outlook-contact-analyzer.git
   git push -u origin main
   ```

2. **Enable GitHub Pages**
   - Go to repository Settings → Pages
   - Source: Deploy from branch
   - Branch: main, folder: / (root)
   - Save

3. **Update Manifest**
   - Your GitHub Pages URL will be: `https://yourusername.github.io/outlook-contact-analyzer`
   - Update all URLs in manifest.xml to use this URL
   - Upload the manifest to Outlook as described in Option 1, step 4

#### Option 3: Deploy to Azure Static Web Apps (Recommended for Production)

1. **Create Azure Static Web App**
   - Go to portal.azure.com
   - Create new Static Web App
   - Connect to your GitHub repository
   - Set build preset to "Custom"
   - Deploy

2. **Update Manifest**
   - Use your Azure Static Web App URL (e.g., `https://your-app.azurestaticapps.net`)
   - Update all URLs in manifest.xml
   - Upload to Outlook

## How to Use

1. **Open the Add-in**
   - Open Outlook on the web
   - Click on any email
   - Look for "Contact Analyzer" or "View Contacts" button in the ribbon
   - The task pane will open on the right side

2. **Scan Emails**
   - Click the "Scan Emails" button
   - Wait while the add-in analyzes your inbox (may take 10-30 seconds)
   - Results will display automatically

3. **View Results**
   - See all contacts sorted by email count
   - Use the search box to filter contacts
   - Click any contact to view their email details

4. **Clear Results**
   - Click "Clear Results" to reset and start fresh

## Technical Details

### API Permissions

The add-in requests `ReadWriteMailbox` permission to:
- Read email messages from your mailbox
- Access recipient information (To, CC, BCC fields)
- Read email metadata (subject, date)

### Data Privacy

- All data processing happens in your browser
- No data is sent to external servers
- The add-in only reads email metadata, not content
- Results are temporary and cleared when you close Outlook

### Performance

- Scans up to 500 emails per request (configurable in code)
- Processing time: approximately 10-30 seconds for 500 emails
- Results are cached until cleared or refreshed

### Browser Compatibility

- Works in Outlook on the web (all modern browsers)
- Requires JavaScript enabled
- Best experience in Chrome, Edge, Firefox, Safari

## Customization

### Changing Email Limit

Edit `taskpane.js`, line with `$top=500`:
```javascript
const getMessagesUrl = `${restUrl}/v2.0/me/messages?$top=1000&$select=...`;
```

### Modifying UI Colors

Edit `taskpane.html` CSS section:
```css
.btn-primary {
    background: linear-gradient(135deg, #YOUR_COLOR_1 0%, #YOUR_COLOR_2 100%);
}
```

### Adding Folders

To scan folders other than Inbox, modify the REST URL in `taskpane.js`:
```javascript
// For Sent Items
const getMessagesUrl = `${restUrl}/v2.0/me/mailfolders/sentitems/messages?$top=500...`;

// For All Items
const getMessagesUrl = `${restUrl}/v2.0/me/messages?$top=500...`;
```

## Troubleshooting

### "Failed to get access token"
- Ensure you're using Outlook on the web
- Check that you're signed in to your Microsoft 365 account
- Try refreshing the page

### "Error scanning emails"
- Check browser console for detailed error messages
- Verify your internet connection
- Ensure the add-in has proper permissions

### Add-in doesn't appear
- Verify manifest.xml has correct URLs
- Check that all files are accessible via HTTPS
- Try removing and re-adding the add-in
- Clear browser cache

### No contacts found
- Ensure your inbox has emails
- Check that emails have recipients
- Try increasing the email limit in code

## Development

### Local Testing

1. **Use a local HTTPS server**:
   ```bash
   npx http-server -S -C cert.pem -K key.pem
   ```

2. **Or use Webpack Dev Server**:
   ```bash
   npm install webpack webpack-cli webpack-dev-server --save-dev
   npx webpack serve --https
   ```

3. **Update manifest** to use `https://localhost:8080`

### Debugging

- Open browser Developer Tools (F12)
- Check Console tab for JavaScript errors
- Use Network tab to monitor API calls
- Add `console.log()` statements in taskpane.js

## Future Enhancements

Potential features to add:
- Export contacts to CSV
- Email frequency timeline charts
- Filter by date range
- Folder selection
- Contact grouping (domains, organizations)
- Email thread analysis
- Sentiment analysis integration

## Support

For issues or questions:
1. Check the Troubleshooting section
2. Review Microsoft's Outlook Add-in documentation
3. Check browser console for errors
4. Verify all URLs are correct in manifest.xml

## License

This add-in is provided as-is for educational and commercial use.

## Credits

Built with:
- Office.js API
- Microsoft Graph REST API
- Modern vanilla JavaScript
- CSS3 animations

---

**Note**: This add-in works exclusively with Outlook on the web. Desktop versions of Outlook may require different API approaches or additional configuration.
