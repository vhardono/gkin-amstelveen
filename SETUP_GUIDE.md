# Church Bulletin Integration Setup Guide

This guide will help you set up the automated church bulletin (mededelingen) generation system that integrates data from multiple sources.

## Overview

The system integrates data from three sources to create a complete church bulletin:
1. **Dropbox Excel files** - General announcements and notices
2. **Scipio Portal** - Birthday list for church members
3. **GKIN Website** - Preaching roster (preekrooster)

## Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Configure Environment

Copy the environment template and fill in your credentials:

```bash
cp .env.template .env
```

Edit `.env` with your actual values:

```env
# Dropbox Configuration
DROPBOX_ACCESS_TOKEN=your_actual_dropbox_token
DROPBOX_EXCEL_PATH=/Church/Excel

# Scipio Portal Configuration
SCIPIO_BASE_URL=https://your-scipio-domain.com
SCIPIO_USERNAME=your_username
SCIPIO_PASSWORD=your_password
SCIPIO_PIN=your_pin_if_required

# GKIN Website (already configured)
PREEKROSTER_URL=https://gkin.org/main/index.php/nl/preekrooster-2#gsc.tab=0

# Church Information
CHURCH_NAME=Your Church Name
```

### 3. Run the System

#### Preview Data Sources
Check if all sources are accessible:

```bash
python3 main_integration.py preview
```

#### Generate Bulletin for Today
```bash
python3 main_integration.py today
```

#### Generate Bulletin for Next Sunday
```bash
python3 main_integration.py next-sunday
```

## Configuration Details

### Dropbox Setup

1. Go to [Dropbox Developers](https://www.dropbox.com/developers/apps)
2. Create a new app with "Full Dropbox" access
3. Generate an access token
4. Place your Excel files in the configured Dropbox path
5. Set the `DROPBOX_ACCESS_TOKEN` in your `.env` file

### Scipio Portal Setup

1. Get your Scipio portal URL, username, and password
2. Some systems require a PIN code
3. Test login by running the preview command
4. The system will automatically navigate to the birthday list

### GKIN Website

The GKIN preekrooster URL is already configured and should work out of the box.

## Output Files

The system generates two files in the `output/` directory:
- `mededelingen_YYYYMMDD.docx` - Word document
- `mededelingen_YYYYMMDD.pdf` - PDF document

## Troubleshooting

### Common Issues

1. **Dropbox Authentication Error**
   - Ensure your access token is valid and has full Dropbox access
   - Check that the Excel files exist in the specified path

2. **Scipio Login Failed**
   - Verify your username and password are correct
   - Check if a PIN is required
   - Ensure the base URL is correct

3. **No Data Found**
   - Run the preview command to check source availability
   - Check the logs for specific error messages

### Debug Mode

For detailed debugging, check the log output which shows:
- Data collection progress
- Specific errors encountered
- Number of items found from each source

## File Structure

```
windsurf-project/
├── main_integration.py      # Main integration script
├── bulletin_generator.py    # Bulletin generation logic
├── config.py               # Configuration management
├── data_sources/
│   ├── dropbox_reader.py   # Dropbox Excel integration
│   ├── scipio_scraper.py   # Scipio portal scraper
│   └── preekroster_scraper.py # GKIN website scraper
├── output/                 # Generated bulletins
├── templates/              # Bulletin templates (if customized)
├── .env                    # Your configuration (create this)
└── .env.template          # Configuration template
```

## Customization

### Bulletin Template

You can customize the bulletin format by modifying the `bulletin_generator.py` file or by creating custom templates in the `templates/` directory.

### Data Processing

Each data source has its own processing logic in the respective files under `data_sources/`. You can modify these to match your specific data formats.

## Automation

To automate bulletin generation, you can set up a cron job or scheduled task:

```bash
# Generate bulletin every Saturday for Sunday
0 18 * * 6 cd /path/to/windsurf-project && python3 main_integration.py next-sunday
```

## Support

If you encounter issues:
1. Check the log output for specific error messages
2. Run the preview command to verify source connectivity
3. Verify your `.env` configuration
4. Check that all required dependencies are installed
