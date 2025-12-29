# Nameplate Generator

A web-based application for generating professional nameplates from CSV data using PowerPoint templates.

## 📋 Features

- **Easy-to-use web interface** - Upload CSV and PowerPoint template files through a simple drag-and-drop interface
- **Customizable output** - Choose to include Name, Organization, Country, or any combination
- **Automated generation** - Creates individual nameplate slides for each participant
- **Dual-sided nameplates** - Automatically creates top and bottom text (with 180° rotation for table nameplates)
- **Professional formatting** - Auto-adjusts font sizes and maintains consistent styling

## 🚀 Quick Start

### Prerequisites

- Python 3.7 or higher
- Web browser (Chrome, Firefox, Edge, etc.)

### Installation

1. **Install required Python packages:**
   ```bash
   pip install flask flask-cors python-pptx
   ```

2. **Optional - For hidden server with system tray icon:**
   ```bash
   pip install pystray pillow
   ```
   This allows the server to run completely hidden with a small icon in your system tray.

3. **Download the project files:**
   - `server.py` - Backend server
   - `server_tray.pyw` - Hidden server with tray icon (optional)
   - `nameplate.html` - Web interface
   - `start.bat` - Launcher (Windows)
   - `stop.bat` - Stop server (Windows)

### Running the Application

**Option 1: Double-Click Start (Easiest - No Terminal Window)**

1. Double-click `start.bat`
2. The server starts hidden in the background
3. Your browser opens automatically
4. If you installed `pystray` and `pillow`, you'll see a small green icon in your system tray
5. To stop the server:
   - Right-click the tray icon and select "Quit", OR
   - Double-click `stop.bat`

**Option 2: Manual Start (Shows Terminal)**

1. **Start the server:**
   
   Open a terminal/command prompt in the project folder and run:
   ```bash
   python server.py
   ```
   
   You should see:
   ```
   ==================================================
   Nameplate Generator Server
   ==================================================
   
   Server is running on: http://localhost:5000
   
   Keep this window open while using the app.
   Press Ctrl+C to stop the server.
   ==================================================
   ```

2. **Open the web interface:**
   
   Double-click `nameplate.html` to open it in your web browser.

3. **Verify connection:**
   
   The page should show a green banner saying "Server is running" at the top.

## 📝 Usage Guide

### Step 1: Prepare Your Files

**CSV File Requirements:**
- Must have columns in this order: **Name** (Column A), **Organization** (Column B), **Country** (Column C)
- Column headers should be exactly: `Name`, `Organization`, `Country`
- The application will also recognize variations like `name`, `NAME`, `Nama`, `Org`, `organisation`, `ORGANIZATION`, `country`, `COUNTRY`
- Supports comma (`,`), semicolon (`;`), or tab-delimited formats
- Empty cells are allowed (e.g., if someone doesn't have an organization)

**Example CSV:**
```csv
Name,Organization,Country
John Doe,Acme Corp,United States
Jane Smith,Tech Inc,Canada
Ahmad Hassan,Innovation Ltd,Malaysia
Maria Garcia,,Spain
```

**Important Notes:**
- Column A must contain the Name (required)
- Column B should contain the Organization (optional)
- Column C should contain the Country (optional)
- Keep the columns in this A-B-C order for best results

**PowerPoint Template:**
- Must be a `.pptx` file
- Should contain at least one slide with a text box for the nameplate area
- The largest text box on the first slide will be used as the template

### Step 2: Upload Files

1. Click "Click to upload CSV file" and select your CSV file
2. Click "Click to upload PPTX template" and select your PowerPoint template
3. Click "Continue to Options"

### Step 3: Choose What to Include

Select which information to display on the nameplates:
- ✅ **Name** - Always included
- ☐ **Organization** - Optional
- ☐ **Country** - Optional

The preview will show how the nameplate will look based on your selections.

### Step 4: Generate

1. Click the "Generate" button
2. Wait for processing (a spinner will appear)
3. The generated PowerPoint file will automatically download
4. Open the downloaded file to view your nameplates

## 🎨 Output Format

Each nameplate includes:
- **Main nameplate** (bottom): Name in large bold text, organization and country in smaller text below
- **Reversed nameplate** (top): Same content rotated 180° for dual-sided table nameplates

**Example output for "John Doe, Acme Corp, United States":**

```
┌─────────────────────────┐
│                         │
│  setatS detinU - proC emcA  │ (rotated 180°)
│      eoD nhoJ      │
│                         │
├─────────────────────────┤
│                         │
│      John Doe      │
│  Acme Corp - United States  │
│                         │
└─────────────────────────┘
```

## 🔧 Troubleshooting

### Server Status Shows "Offline"

**Problem:** Red banner saying "Server is offline"

**Solutions:**
1. Make sure you've started the server with `python server.py`
2. Check that no other application is using port 5000
3. Verify all required packages are installed
4. Try restarting the server

### Error: "No module named 'flask'"

**Problem:** Missing Python packages

**Solution:**
```bash
pip install flask flask-cors python-pptx
```

### Error: "No participants found in CSV"

**Problem:** CSV file doesn't have a recognizable Name column

**Solutions:**
1. Make sure your CSV has a column named `Name`, `name`, `NAME`, or `Nama`
2. Check that the CSV file isn't empty
3. Verify the CSV file is properly formatted

### Generated File Won't Open

**Problem:** Downloaded PowerPoint file is corrupted

**Solutions:**
1. Check that your template PPTX file is valid and opens in PowerPoint
2. Make sure the server completed processing (check terminal for errors)
3. Try with a simpler template file

### Browser Can't Download File

**Problem:** Generate button works but file doesn't download

**Solutions:**
1. Check your browser's download settings
2. Allow pop-ups for the page if blocked
3. Try a different browser

## 🛑 Stopping the Server

**If using start.bat (hidden mode):**
- Right-click the green system tray icon and select "Quit", OR
- Double-click `stop.bat`

**If running manually in terminal:**
1. Go to the terminal window where the server is running
2. Press `Ctrl+C`
3. The server will shut down

## 💡 Tips

- **Keep it running:** Leave the terminal window with the server open while using the application
- **Multiple generations:** You can generate nameplates multiple times without restarting the server
- **Template design:** For best results, use a simple template with one large text box
- **CSV formatting:** The application auto-detects comma, semicolon, or tab delimiters
- **Large batches:** The application can handle hundreds of participants in one go

## 📁 Project Structure

```
nameplate-generator/
├── server.py           # Backend Flask server
├── server_tray.pyw     # Hidden server with system tray icon (optional)
├── nameplate.html      # Frontend web interface
├── start.bat          # Windows launcher - runs hidden (recommended)
├── stop.bat           # Stop the hidden server
├── start.sh           # Mac/Linux launcher (optional)
├── template.pptx      # Your default template (optional but recommended)
└── README.md          # This file
```

## 💡 Using a Default Template (Recommended)

To avoid uploading a PowerPoint template every time:

1. **Create or choose your template** - Design your nameplate template in PowerPoint
2. **Save it as `template.pptx`** in the same folder as the other files
3. **The app will automatically use it!** - No need to upload a template anymore

When `template.pptx` exists in the folder:
- ✅ The server will show: "Found template.pptx - will use this automatically"
- ✅ The upload template section will be hidden in the web interface
- ✅ You only need to upload your CSV file each time

If you don't have `template.pptx`:
- You'll need to upload both CSV and template files each time
- The app will show the template upload section

## 🔒 Privacy & Security

- All processing happens locally on your computer
- No data is sent to external servers
- Files are temporarily stored during processing and deleted automatically
- Your CSV and PowerPoint files never leave your machine

## 📄 License

This project uses the following open-source libraries:
- Flask (BSD License)
- python-pptx (MIT License)

## 🆘 Support

If you encounter any issues:
1. Check the troubleshooting section above
2. Verify all prerequisites are installed
3. Check the terminal window for error messages
4. Try with a simple test CSV file first

---

**Created with ❤️ for easy nameplate generation**