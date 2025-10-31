# 📦 Installation Guide - Email to Excel Agent

## Prerequisites

- Python 3.7 or higher
- pip (Python package manager)
- Internet connection
- Email account (Gmail, Outlook, etc.)

## Step-by-Step Installation

### 1. Verify Python Installation

```bash
python3 --version
```

You should see Python 3.7 or higher. If not, install Python from [python.org](https://www.python.org/downloads/)

### 2. Install Dependencies

Navigate to the project directory and run:

```bash
pip install -r requirements.txt
```

Or install individually:

```bash
pip install streamlit pandas openpyxl google-auth google-auth-oauthlib google-auth-httplib2 google-api-python-client python-dotenv
```

### 3. Verify Installation

Run the verification script:

```bash
python3 verify_structure.py
```

You should see all checks pass with ✅ marks.

### 4. Run the Application

```bash
streamlit run email_to_excel_app.py
```

The application will open in your browser at `http://localhost:8501`

## Platform-Specific Instructions

### Windows

```cmd
# Install Python from python.org
# Open Command Prompt or PowerShell

# Navigate to project directory
cd path\to\project

# Install dependencies
pip install -r requirements.txt

# Run application
streamlit run email_to_excel_app.py
```

### macOS

```bash
# Install Python (if not already installed)
brew install python3

# Navigate to project directory
cd /path/to/project

# Install dependencies
pip3 install -r requirements.txt

# Run application
streamlit run email_to_excel_app.py
```

### Linux

```bash
# Install Python (if not already installed)
sudo apt-get update
sudo apt-get install python3 python3-pip

# Navigate to project directory
cd /path/to/project

# Install dependencies
pip3 install -r requirements.txt

# Run application
streamlit run email_to_excel_app.py
```

## Virtual Environment (Recommended)

Using a virtual environment keeps dependencies isolated:

### Create Virtual Environment

```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
# On Windows:
venv\Scripts\activate
# On macOS/Linux:
source venv/bin/activate

# Install dependencies
pip install -r requirements.txt

# Run application
streamlit run email_to_excel_app.py
```

### Deactivate Virtual Environment

```bash
deactivate
```

## Dependency Details

| Package | Version | Purpose |
|---------|---------|---------|
| streamlit | Latest | Web interface framework |
| pandas | Latest | Data manipulation |
| openpyxl | Latest | Excel file creation |
| google-auth | Latest | Google authentication |
| google-auth-oauthlib | Latest | OAuth2 flow |
| google-auth-httplib2 | Latest | HTTP library for Google APIs |
| google-api-python-client | Latest | Gmail API client |
| python-dotenv | Latest | Environment variable management |

## Troubleshooting Installation

### Issue: "pip: command not found"

**Solution:**
```bash
# Try pip3 instead
pip3 install -r requirements.txt

# Or use python -m pip
python3 -m pip install -r requirements.txt
```

### Issue: "Permission denied"

**Solution:**
```bash
# Install for current user only
pip install --user -r requirements.txt

# Or use sudo (Linux/macOS)
sudo pip3 install -r requirements.txt
```

### Issue: "No module named 'streamlit'"

**Solution:**
```bash
# Ensure you're in the correct environment
# Reinstall dependencies
pip install --upgrade -r requirements.txt
```

### Issue: SSL Certificate Error

**Solution:**
```bash
# Install with trusted host
pip install --trusted-host pypi.org --trusted-host files.pythonhosted.org -r requirements.txt
```

### Issue: Outdated pip

**Solution:**
```bash
# Upgrade pip
pip install --upgrade pip

# Then install dependencies
pip install -r requirements.txt
```

## Verifying Installation

### Check Installed Packages

```bash
pip list
```

Look for:
- streamlit
- pandas
- openpyxl
- google-auth
- google-api-python-client

### Test Import

```bash
python3 -c "import streamlit; import pandas; import openpyxl; print('All imports successful!')"
```

### Run Verification Script

```bash
python3 verify_structure.py
```

All checks should pass.

## Email Provider Setup

### Gmail Setup

1. **Enable 2-Factor Authentication**
   - Go to [Google Account Security](https://myaccount.google.com/security)
   - Enable 2-Step Verification

2. **Generate App Password**
   - Go to [App Passwords](https://myaccount.google.com/apppasswords)
   - Select "Mail" and your device
   - Copy the 16-character password

3. **Enable IMAP**
   - Go to [Gmail Settings](https://mail.google.com/mail/u/0/#settings/fwdandpop)
   - Click "Forwarding and POP/IMAP"
   - Enable IMAP
   - Save changes

### Outlook Setup

1. **Enable IMAP**
   - Go to Outlook Settings
   - Select "Mail" > "Sync email"
   - Enable IMAP

2. **Use Regular Password**
   - No app password needed for Outlook
   - Use your regular account password

### Other Providers

Check your email provider's documentation for:
- IMAP server address
- IMAP port (usually 993)
- SSL/TLS requirements
- App password requirements

## Running in Production

### Using Screen (Linux/macOS)

```bash
# Start screen session
screen -S email-agent

# Run application
streamlit run email_to_excel_app.py

# Detach: Ctrl+A, then D
# Reattach: screen -r email-agent
```

### Using nohup

```bash
nohup streamlit run email_to_excel_app.py &
```

### Using systemd (Linux)

Create `/etc/systemd/system/email-agent.service`:

```ini
[Unit]
Description=Email to Excel Agent
After=network.target

[Service]
Type=simple
User=your-username
WorkingDirectory=/path/to/project
ExecStart=/usr/bin/streamlit run email_to_excel_app.py
Restart=always

[Install]
WantedBy=multi-user.target
```

Enable and start:

```bash
sudo systemctl enable email-agent
sudo systemctl start email-agent
```

## Updating

### Update Dependencies

```bash
pip install --upgrade -r requirements.txt
```

### Update Application

```bash
git pull  # If using git
# Or download latest version
```

## Uninstallation

### Remove Dependencies

```bash
pip uninstall -r requirements.txt -y
```

### Remove Virtual Environment

```bash
rm -rf venv
```

### Remove Application Files

```bash
rm -rf /path/to/project
```

## Getting Help

If you encounter issues:

1. Check this installation guide
2. Review the troubleshooting section
3. Verify Python and pip versions
4. Check internet connection
5. Review error messages carefully

## Next Steps

After successful installation:

1. Read the [Quick Start Guide](QUICKSTART.md)
2. Review the [Full Documentation](README_EMAIL_TO_EXCEL.md)
3. Run the application: `streamlit run email_to_excel_app.py`
4. Connect your email account
5. Start exporting emails!

---

**Installation Complete?** Run `streamlit run email_to_excel_app.py` to get started!
