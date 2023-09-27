# TNB OCR

TNB OCR is a Python application that extracts data from PDF documents, performs Optical Character Recognition (OCR), and stores the extracted information in an Excel spreadsheet.

## Table of Contents
- [Features](#features)
- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Usage](#usage)
- [Contributing](#contributing)
- [License](#license)

## Features

- Extracts data from PDF files.
- Performs OCR on specific regions of interest within each page of the PDF.
- Stores the extracted data in an Excel spreadsheet with customizable headers.
- Provides options for customizing the sheet name and adding custom text to the Excel file.

## Prerequisites

Before you begin, ensure you have met the following requirements:

- Python 3.7 or later installed on your system.

## Installation

1. Clone this repository to your local machine:

    ```bash
    git clone https://github.com/ZennZhiEn/TNB_OCR.git
    ```

2. Navigate to the project directory:

    ```bash
    cd TNB_OCR
    ```

3. Create a virtual environment (optional but recommended):

    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows, use `venv\Scripts\activate`
    ```

4. Install the required libraries from the `requirements.txt` file:

    ```bash
    pip install -r requirements.txt
    ```

   If you encounter issues during installation, you can install the libraries individually. See [Installing Libraries](#installing-libraries) below.

### Installing Libraries

If you encounter issues during installation, you can install the libraries one by one. For example:

```bash
# Install tkinter (a standard library in Python, no need to install)
# pip install pdf2image
pip install Pillow
pip install opencv-python
pip install numpy
pip install paddleocr
pip install openpyxl
