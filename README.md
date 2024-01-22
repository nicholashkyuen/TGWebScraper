# TGWebScraper

TGWebScraper is a powerful web scraping tool designed to retrieve shipping information of packages from forwarders' websites. It currently supports DHL, TNT, FedEx, UPS, and SF Express. With just a few simple inputs, you can obtain an Excel file containing key information such as Master Air Waybill, ETA, Freight No., and Port, derived from the forwarders' websites.

## Installation

To use TGWebScraper, follow these steps:

1. Clone the repository to your local machine using the following command:

   ```bash
   git clone https://github.com/nicholashkyuen/TGWebScraper.git
   ```

2. Navigate to the project directory:

   ```bash
   cd TGWebScraper
   ```

3. Install the required dependencies. It is recommended to use a virtual environment:

   ```bash
   python3 -m venv env
   source env/bin/activate  # For Linux/Mac
   env\Scripts\activate  # For Windows
   pip install -r requirements.txt
   ```

## Usage

1. Open the terminal or command prompt and navigate to the project directory.

2. Run the `main.py` script:

   ```bash
   python main.py
   ```

3. The program will launch with a simple user interface that prompts you to enter the forwarder and waybill number.

4. Choose the appropriate forwarder from the dropdown menu (DHL, TNT, FedEx, UPS, or SF Express) and provide the corresponding waybill number in the input field.
   
5. Click the "Run" button to initiate the scraping process.

6. The scraper will retrieve the shipping information from the forwarder's website and generate an Excel file with the extracted data.

7. Once the scraping is complete, you can choose a file in your computer to save the generated Excel file named RPA_Custom_Declaration_YYYYMMDD.xlsx. Note that YYYYMMDD represents the current date in the format year-month-day.

Example: RPA_Custom_Declaration_20240122.xlsx


## Benefits

TGWebScraper offers the following benefits:

- **Time-saving**: By automating the process of retrieving shipping information, TGWebScraper can save you more than half of the time typically spent on preparing documents for customs declaration.

- **User-friendly**: The program is designed to be easy to use. With just a few simple steps, you can obtain the required shipping information in an organized Excel file.

- **Support for multiple forwarders**: TGWebScraper currently supports popular forwarders such as DHL, TNT, FedEx, UPS, and SF Express. This ensures compatibility with a wide range of shipping providers.

## Contributions

Contributions to TGWebScraper are welcome! If you have any suggestions, bug reports, or feature requests, please open an issue on the [GitHub repository](https://github.com/your-username/TGWebScraper/issues).

## Disclaimer

TGWebScraper is a tool developed for personal use. It relies on web scraping techniques to extract data from forwarders' websites, and its functionality may be affected if the structure or layout of these websites changes. Use this tool responsibly and in compliance with the terms and conditions of the forwarders' websites.
