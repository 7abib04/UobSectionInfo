
---

# UOB Course Information Scraper

## Overview

This script automates the process of collecting course information from the University of Bahrain (UOB) website using Selenium. It allows users to input a course code and retrieves the course sections and instructor details, saving the data into an Excel file.

## Features

- Interacts with the UOB website to fetch course information.
- Saves the retrieved data into an Excel file.
- Handles input for the course code.
- Selects the year and semester automatically.
- Provides error handling for unavailable courses and file writing issues.

## Requirements

- Python 3.x
- Selenium
- xlwt
- Chrome WebDriver

## Installation

1. Install the required Python packages:

```bash
pip install selenium xlwt
```

2. Download and install the Chrome WebDriver compatible with your version of Chrome from [here](https://sites.google.com/a/chromium.org/chromedriver/downloads).

## Usage

### 1. Set up the Chrome WebDriver

Ensure the Chrome WebDriver is in your system PATH or provide the executable path in the script.

### 2. Run the Script

Execute the script in your Python environment:

```bash
python script_name.py
```

### 3. Input Course Code

When prompted, enter the course code you want to fetch information for.

```bash
enter the course code: <course_code>
```

### 4. Output

The script will create an Excel file named `sections.xls` containing the course sections and instructor details.

## Example

Here is an example of how to use the script:

1. Run the script:

```bash
python uob_course_scraper.py
```

2. Input the course code when prompted:

```bash
enter the course code: STAT101
```

3. The script will fetch the information and save it to `sections.xls`.

## Code Explanation

The script performs the following steps:

1. **Imports Required Libraries**:
    - `selenium.webdriver` for web automation.
    - `xlwt` for writing data to an Excel file.

2. **Initializes Variables**:
    - `count` to keep track of the row number in the Excel sheet.
    - `sections` to store the fetched sections data.

3. **Prompts User for Course Code**:
    - Prompts the user to input the course code.

4. **Sets Up Excel Workbook**:
    - Creates an Excel workbook and sheet.
    - Defines styles for the header and data cells.

5. **Interacts with UOB Website**:
    - Navigates to the UOB course information page.
    - Selects the year and semester.
    - Inputs the course code and submits the form.
    - Retrieves the course sections and instructor information.

6. **Writes Data to Excel**:
    - Parses the retrieved data and writes it to the Excel sheet.
    - Saves the Excel file as `sections.xls`.

7. **Error Handling**:
    - Handles errors for unavailable courses and file writing issues.

## License

This project is licensed under the MIT License.

---
