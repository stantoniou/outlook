# Quick Start Deployment Guide

## For IT Project Managers & Administrators

### Executive Summary
This Outlook add-in provides contact-based email analysis for Outlook Online users. It scans the inbox, groups emails by contact, and shows recipient type breakdown (To/CC/BCC).

### Deployment Timeline
- **Setup**: 15-30 minutes
- **Testing**: 10 minutes
- **Rollout**: Immediate (per-user) or enterprise-wide

---

## Fastest Deployment Path (GitHub Pages - Free)

### Step 1: Host the Files (5 minutes)

1. Create a new GitHub repository (public or private)
2. Upload these files:
   - `taskpane.html`
   - `taskpane.js`
   - `commands.html`
3. Enable GitHub Pages in Settings → Pages
4. Note your URL: `https://[username].github.io/[repo-name]`

### Step 2: Configure Manifest (5 minutes)

1. Open `manifest.xml`
2. Find & replace ALL instances of `https://yourdomain.com` with your GitHub Pages URL
3. Generate new GUID at `<Id>` field (use https://guidgenerator.com)
4. Update `<ProviderName>` to your organization name

### Step 3: Deploy to Outlook (5 minutes)

**Option A: Individual User**
1. Go to outlook.office.com
2. Settings → View all Outlook settings
3. Mail → Customize actions → Get add-ins
4. My add-ins → Add a custom add-in → Add from file
5. Upload `manifest.xml`

**Option B: Enterprise Deployment (Admin)**
1. Microsoft 365 Admin Center
2. Settings → Integrated apps → Upload custom apps
3. Upload `manifest.xml`
4. Assign to users/groups
5. Deploy

---

## Production Deployment (Azure - Recommended)

### Prerequisites
- Azure subscription
- GitHub account

### Steps

1. **Create Azure Static Web App**
   ```bash
   # Via Azure Portal
   # Resources → Create → Static Web App
   # Connect GitHub repo
   # Build preset: Custom
   # Skip build settings
   ```

2. **Update Manifest**
   - URL format: `https://[app-name].azurestaticapps.net`
   - Replace all URLs in manifest.xml

3. **Deploy via Admin Center** (see above)

### Cost
- GitHub Pages: **Free**
- Azure Static Web Apps: **Free tier** (100GB bandwidth/month)

---

## Security & Compliance

### Permissions Required
- `ReadWriteMailbox` - Required to read email metadata

### Data Handling
- No data leaves the browser
- No external API calls (except Microsoft Graph)
- No persistent storage
- GDPR compliant (no data collection)

### Network Requirements
- HTTPS (SSL/TLS required)
- Outbound HTTPS to Microsoft Graph API
- No inbound connections needed

---

## Testing Checklist

- [ ] Add-in appears in Outlook ribbon
- [ ] Task pane opens successfully
- [ ] "Scan Emails" button works
- [ ] Contacts display correctly
- [ ] Search/filter functions work
- [ ] No console errors in browser DevTools
- [ ] Works in multiple browsers (Chrome, Edge, Firefox)

---

## Rollout Strategy

### Phase 1: Pilot (Week 1)
- Deploy to 5-10 test users
- Collect feedback
- Monitor for issues

### Phase 2: Department (Week 2)
- Deploy to single department
- Provide training/documentation
- Monitor usage

### Phase 3: Enterprise (Week 3+)
- Full rollout
- Ongoing support

---

## Troubleshooting Quick Reference

| Issue | Solution |
|-------|----------|
| Add-in not appearing | Check manifest URLs, verify HTTPS |
| "Access token failed" | User needs to re-authenticate |
| No contacts found | Increase email limit in code |
| Slow performance | Reduce scan limit from 500 to 250 |

---

## Monitoring & Maintenance

### What to Monitor
- User adoption rates
- Error reports
- Performance feedback

### Maintenance Schedule
- Monthly: Check for Office.js updates
- Quarterly: Review Microsoft Graph API changes
- Annually: Security audit

---

## Support Resources

### Microsoft Documentation
- [Outlook Add-ins Overview](https://docs.microsoft.com/en-us/office/dev/add-ins/outlook/)
- [Graph API Reference](https://docs.microsoft.com/en-us/graph/api/overview)

### Internal Support
- Include link to your internal support portal
- Add FAQ document URL
- Provide contact email

---

## Customization for Your Organization

### Branding
1. Update icons (16x16, 32x32, 64x64, 80x80, 128x128)
2. Modify colors in `taskpane.html` CSS
3. Update app name in manifest

### Functionality
- Change email limit (default: 500)
- Add additional folders (Sent Items, Archive)
- Modify UI layout

### See README.md for detailed customization instructions

---

## Compliance & Approval

### Information to Provide to Security Team
- **Data Access**: Email metadata only (recipients, dates, subjects)
- **Data Storage**: Browser memory only, session-based
- **External Dependencies**: Microsoft Graph API only
- **Code Review**: All source code included
- **Hosting**: Your organization's infrastructure

### Approval Checklist
- [ ] Security review completed
- [ ] Privacy impact assessment
- [ ] IT infrastructure approval
- [ ] User acceptance testing
- [ ] Documentation approved

---

## Quick Commands

```bash
# Validate manifest
npx office-addin-manifest validate manifest.xml

# Start local HTTPS server for testing
npx http-server -S -p 8080

# Check manifest syntax
xmllint --noout manifest.xml
```

---

**Questions?** Contact your IT support team or refer to README.md for detailed documentation.
