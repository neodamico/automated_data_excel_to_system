### [ENGLISH]
# Automated Data Entry from Excel to System

This project is a Python-based automation script designed to read data from an Excel spreadsheet and automatically input that data into a specified system. The automation uses the libraries `openpyxl` for handling `.xlsx` files and `pyautogui` for simulating mouse and keyboard actions to interact with the system interface.

### Overview
The automation reads through an Excel spreadsheet named `vendas_de_produtos.xlsx`, iterating over each row of data. It captures the details such as client name, product, quantity, and category, and inputs them into the appropriate fields of the system's interface. After completing the data input for a row, the automation saves and confirms the entry before moving on to the next row.

This script is highly adaptable and can be customized for other types of systems where automated data entry is needed.

### Features
- Excel Data Handling: Reads Excel files using the `openpyxl` library.
- Automated Input: Utilizes `pyautogui` to automate the input of data, replicating user interaction with a system.
- Customizable Workflow: The script can be adjusted for different data sets, field positions, or application use cases.
 
### Requirements
- Python 3.x
- `openpyxl` library
- `pyautogui` library
- Excel file: `vendas_de_produtos.xlsx`

### Installation

1. Clone the repositorie:
```bash
git clone https://github.com/neodamico/automated_data_excel_to_system
``` 
2. Navigate into the project directory:
```
cd automated-data-entry

```
3. Create and activate a virtual environment:
```
python -m venv venv
source venv/bin/activate   # For Linux/Mac
.\venv\Scripts\activate    # For Windows

```
4. Install required libraries:
```
pip install openpyxl pyautogui

```
5. Ensure that your Excel file `(vendas_de_produtos.xlsx)` is in the same directory as app.py.

### Usage

1. Activate your virtual environment (if not already active).
2. Run the script:
```
python app.py

```

Once the script is running, it will begin to read each row from the Excel file and fill in the corresponding fields in the system's interface. The mouse will move automatically to the specific coordinates of the fields.

> [!NOTE]
> Ensure that your system interface is open and in the correct position for the automation to work as expected.

### Stopping the Automation
To stop the automation, you can:

Press Ctrl+alt+del in the terminal to interrupt the script.
If you lose control of your mouse, force-quit the Python process via Task Manager (Windows) or Activity Monitor (Mac).

### Customization
To modify the script for different use cases, adjust the coordinates of the fields in the pyautogui.click() function calls and tailor the data handling to match the structure of your spreadsheet.

### Future Enhancements

- [ ] Implement keyboard shortcuts for user interruption.
- [ ] Add a dynamic way to set the coordinates for the system's input fields.
- [ ] Enhance error handling for more robust data processing.
