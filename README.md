# Web Scraping and Excel Export Script

This Node.js script utilizes Puppeteer for web scraping and ExcelJS for exporting data to Excel. The script is designed to extract product details from a specific website, and save the information in an Excel file.

## Dependencies
- **puppeteer:** Node library for browser automation.
- **fs:** Node.js File System module for file operations.
- **ExcelJS:** Node library for working with Excel files.
- **https:** Node.js HTTP module for making HTTPS requests.

## Functionality
1. **urlToBase64(url):** Converts an image from a URL to a base64-encoded string.
2. **add_row(product_details, row_index):** Adds a row to the Excel file for a given product.

## Usage
1. Install dependencies: `npm install puppeteer exceljs`
2. Run the script: `node app.js`
3. The script will navigate through product pages, collect details, and save the information in `output.xlsx`.

## Important Notes
- The script assumes a specific HTML structure on the target website. Changes in the structure may require script modifications.
- The headless browser is set to visible (`headless: false`). For production, set it to `headless: "new"` for a background operation.
- Additional error handling may be needed based on specific use cases.

Feel free to adapt the script to your requirements and consult the Puppeteer and ExcelJS documentation for further customization.
