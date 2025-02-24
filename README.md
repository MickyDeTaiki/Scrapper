# Project Setup Guide

## Prerequisites
- Mac or Linux operating system
- Visual Studio Code installed
- Python 3 installed

## Setup Instructions

### 1. Prepare the Project Directory
Ensure all project files are located in the same directory in Visual Studio Code.

### 2. Create a Virtual Environment
Open a terminal and run the following command to set up a virtual environment:
```sh
python3 -m venv <name_of_virtual_environment>
```

### 3. Activate the Virtual Environment
Run the appropriate command to activate the virtual environment:
```sh
source <name_of_virtual_environment>/bin/activate
```

### 4. Install Dependencies
Use the `requirements.txt` file to install necessary dependencies:
```sh
pip install -r requirements.txt
```

### 5. Download and Setup ChromeDriver
Download the `chromedriver.exe` file and place it in the project directory. You can get it from:
[ChromeDriver Download](https://googlechromelabs.github.io/)

## Running the Project
- Run `main_PLZ.py` to scrape the list of PLZ from the start.
- Run `main.py` to scrape the list of a given PLZ number.

## Troubleshooting
- Ensure Python and pip are installed and accessible.
- Verify that the virtual environment is activated before running `pip install`.
- Check that `chromedriver.exe` is correctly placed and has execution permissions.

For additional help, refer to the official documentation or community forums.

