# Security Policy

## Supported Versions

| Version | Supported          |
| ------- | ------------------ |
| 3.x.x   | :white_check_mark: |
| 2.x.x   | :x:                |
| 1.x.x   | :x:                |

## Reporting a Vulnerability

If you discover a security vulnerability in SMIRT, please report it responsibly:

### 🔒 Private Disclosure
- **DO NOT** create public GitHub issues for security vulnerabilities
- Send details privately to the repository maintainer
- Include steps to reproduce the issue
- Provide impact assessment if possible

### 📧 Contact
- Create a private security advisory on GitHub
- Or contact via repository discussions (mark as sensitive)

### ⏱️ Response Time
- Initial acknowledgment: Within 48 hours
- Status update: Within 7 days
- Resolution timeline: Depends on severity

## Security Considerations

### 🔐 Authentication
- User credentials stored in `users.json` (client-side)
- Session management via LocalStorage (24h expiry)
- **Recommendation**: Use proper authentication service for production

### 📡 Data Transmission
- JSONP communication with Google Apps Script
- HTTPS required for secure data transmission
- No sensitive data in URLs or logs

### 🗄️ Data Storage
- Google Sheets: Business data (non-sensitive)
- Google Drive: Signature images (access controlled)
- LocalStorage: Session data and drafts (client-side)

### 🖊️ Digital Signatures
- Canvas-based signatures (image format)
- Stored as PNG files in Google Drive
- Unique timestamped filenames

## Best Practices

### 🚀 Deployment
- Always use HTTPS in production
- Regularly update Google Apps Script permissions
- Monitor Google Drive storage quotas
- Implement proper backup procedures

### 👥 User Management
- Regular review of user credentials
- Implement proper role-based access
- Monitor user activity logs

### 📊 Data Protection
- Regular backup of Google Sheets data
- Secure handling of signature files
- Compliance with data protection regulations

## Known Limitations

### 🔍 Current Security Limitations
- Client-side authentication (suitable for trusted environments)
- No server-side session validation
- Limited audit logging capabilities

### 🛡️ Mitigation Strategies
- Deploy in secure network environments
- Regular security reviews
- User training on proper usage

---

**Last Updated**: November 15, 2025  
**Version**: 3.0.0