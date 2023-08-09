# ![PAA Excel Page Creator],

This Streamlit app allows you to generate Excel front pages based on uploaded data files and a template. You can specify values for overwriting specific cells in the template and then generate a zip file containing the generated Excel files.

## How to Use
Clone or download this repository to your local machine.

Install the required dependencies by running the following command:

pip install streamlit pandas openpyxl pillow

Navigate to the directory where you have saved the code and run the Streamlit app using the following command:

streamlit run app.py

The Streamlit app will open in your web browser.

## App Usage
### Step 1: Upload Data File
Click the "Browse File" button in the sidebar.
Choose an Excel file (`.xlsx` or `.xls`) containing the data you want to use for generating front pages.    

### Step 2: Upload Template
Upload a template Excel file (`.xlsx` or `.xls`) that will be used as the basis for generating front pages.

### Step 3: Enter Values for Overwriting Cells
After uploading the data file and template, you can enter values for overwriting specific cells in the template. For example, you can enter a value for the "Slut dokumentation" cell.

### Step 4: Generate Files
Click the "Generate Files" button.
The app will generate individual Excel files based on the data from the uploaded file and the template. The specified cell values will be overwritten.
The generated files will be saved in a zip file named `generated_files.zip`.

### Step 5: Download Generated Files
Once the files are generated, a download button will appear.
Click the "Download Generated Files" button to download the zip file containing the generated Excel files.

## Notes
If you encounter any issues or errors, please make sure you have provided valid input files and values for overwriting cells.
Make sure to customize the cell mappings in the `cell_mapping` dictionary within the code if you need to adjust cell positions.
