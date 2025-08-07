# Trendyol Scraper

![Trendyol Scraper](https://img.icons8.com/color/48/000000/scraper.png)

## Overview

The **Trendyol Scraper** is a Python-based web scraping tool designed to extract seller information from the Trendyol website. It utilizes Selenium for web automation and provides a user-friendly GUI built with PyQt6. The scraped data is formatted and saved into an Excel file for easy analysis.

## Features

- Scrapes seller information including ratings, review counts, and more.
- User-friendly GUI for selecting files and monitoring progress.
- Data is saved in a well-structured Excel format.

## Installation

To set up the project, follow these steps:

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/TRENDYOL-SCRAPER-GUI.git
   ```

2. Navigate to the project directory:
   ```
   cd TRENDYOL-SCRAPER-GUI
   ```

3. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the GUI application:
   ```
   python scraperGUI.py
   ```

2. Click on "📂 Dosya Seç" to select an Excel file containing seller names.

3. Click on "🚀 Başlat" to start the scraping process.

4. Monitor the progress in the log output and progress bar.

5. Once completed, the scraped data will be saved in a new Excel file named **Trendyol Satıcı Bilgileri (Detaylı).xlsx**.

## Requirements

Make sure you have the following installed:

- Python 3.x
- Selenium
- Pandas
- OpenPyXL
- PyQt6

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Acknowledgments

- [Selenium](https://www.selenium.dev/)
- [PyQt6](https://www.riverbankcomputing.com/software/pyqt/intro)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)