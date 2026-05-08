# Church Bulletin Generator (Mededelingen)

A Python application that generates church bulletins by combining data from multiple sources:

- Excel files from Dropbox directory
- Preekroster (preaching roster) from website
- Birthday list from Scipio (authenticated website)

## Installation

1. Clone or download this project
2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Copy the configuration file:
```bash
cp .env.example .env
```

4. Edit `.env` with your actual configuration values

## Configuration

Edit the `.env` file with your specific settings:

### Dropbox Integration
- `DROPBOX_ACCESS_TOKEN`: Your Dropbox API access token
- `DROPBOX_EXCEL_PATH`: Path to Excel files in Dropbox (default: `/Church/Excel`)

### Website Integration
- `PREEKROSTER_URL`: URL of the preaching roster page
- `SCIPIO_BASE_URL`: Base URL of the Scipio system
- `SCIPIO_USERNAME`: Username for Scipio login
- `SCIPIO_PASSWORD`: Password for Scipio login

### Output Configuration
- `OUTPUT_DIR`: Directory for generated bulletins (default: `./output`)
- `TEMPLATE_DIR`: Directory for templates (default: `./templates`)
- `CHURCH_NAME`: Name of your church

## Usage

Run the bulletin generator:
```bash
python main.py
```

The application will:
1. Read Excel files from your Dropbox directory
2. Scrape the preaching roster from the configured website
3. Login to Scipio and retrieve the birthday list
4. Generate bulletin in both Word (.docx) and PDF formats
5. Save files to the output directory

## Output Files

Generated bulletins are saved with the format:
- `mededelingen_YYYYMMDD.docx` (Word document)
- `mededelingen_YYYYMMDD.pdf` (PDF document)

## Excel File Structure

The system expects Excel files with columns like:
- `titel` or `title`: Announcement title
- `inhoud` or `content` or `omschrijving`: Announcement content
- `datum` or `date`: Date
- `prioriteit` or `priority`: Priority (1-10)
- `categorie` or `category`: Category

## Website Scraping

### Preekroster
The scraper tries multiple approaches to extract preaching roster data:
- HTML tables with date/speaker columns
- Lists with roster entries
- Div-based layouts

### Scipio Birthday List
The system:
1. Logs into Scipio with provided credentials
2. Searches for birthday pages at common URLs
3. Extracts birthday information from tables, lists, or cards

## Troubleshooting

### Dropbox Issues
- Ensure your Dropbox access token has read permissions
- Verify the Excel file path is correct

### Website Scraping Issues
- Check that URLs are accessible
- For Scipio, verify username and password are correct
- Some websites may require additional authentication steps

### Excel File Issues
- Ensure files are in `.xlsx` or `.xls` format
- Check that column headers match expected names

## Customization

### Adding New Data Sources
1. Create a new scraper class in `data_sources/`
2. Add it to `main.py`
3. Update `bulletin_generator.py` to process the new data

### Customizing Templates
- Modify the `_generate_word_bulletin()` and `_generate_pdf_bulletin()` methods
- Add custom styling and formatting as needed

## Dependencies

- `pandas`: Excel file processing
- `openpyxl`: Excel file reading
- `requests`: HTTP requests
- `beautifulsoup4`: HTML parsing
- `selenium`: Web automation (if needed)
- `python-docx`: Word document generation
- `reportlab`: PDF generation
- `jinja2`: Template engine
- `python-dotenv`: Environment variable management
- `dropbox`: Dropbox API integration
- `lxml`: XML/HTML parser

## License

This project is open source. Feel free to modify and distribute according to your needs.
