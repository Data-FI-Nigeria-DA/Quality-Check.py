import pandas as pd
import os
from datetime import datetime
import re
import numpy as np

# Path to the directory containing the CSV files
folder_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/FY26Q3_RADET/May 2026/8th May, 2026'

# Path for the separate viral load output file
viral_load_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Cleaned_Viral_Load_Values_tbd.xlsx'

# Path for the unique CD4 values output file
unique_cd4_output_path = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Cleaned_CD4_Values_tbd.xlsx'


# Define output directory for projects
output_base_dir = 'C:/Users/oluwabukola.arowolo/OneDrive - Palladium International, LLC/Documents/DataFi/Project_Export_Quality_Check_all/May 2026/Project_Export_Quality_Check_8thMay'
os.makedirs(output_base_dir, exist_ok=True)

# Define filter date
filter_date = datetime(2025, 1, 1)
#end_date =datetime(2024, 12, 31)

# Combine all CSV files into one DataFrame, specifying 'latin1' encoding
all_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith(('.csv', '.xlsx'))]

# Combine all files into one DataFrame
combined_data = pd.DataFrame()

for file in all_files:
    try:
        if file.endswith('.csv'):
            data = pd.read_csv(file, encoding='latin1', on_bad_lines='skip')
        elif file.endswith('.xlsx'):
            data = pd.read_excel(file, engine='openpyxl')
        else:
            continue
        
        # Add Filename column
        data['Filename'] = os.path.basename(file)
        combined_data = pd.concat([combined_data, data], ignore_index=True)
    except Exception as e:
        print(f"Error processing file {file}: {e}")

if combined_data.empty:
    print("No valid files found or data could not be combined.")
    exit()

# Extract project name from the 'Filename' column
combined_data['ProjectName'] = combined_data['Filename'].str.split('_').str[0]
#combined_data['ProjectName'] = combined_data['IP']

# Convert date columns to datetime format
for col in ['Date of Birth (yyyy-mm-dd)', 'Last Pickup Date (yyyy-mm-dd)', 'ART Start Date (yyyy-mm-dd)', 'Date of Current ART Status', 'Date of Current Viral Load (yyyy-mm-dd)',
             'Confirmed Date of Previous ART Status', 'Date of Current ViralLoad Result Sample (yyyy-mm-dd)', 'Date of Start of Current ART Regimen','Date of Last CD4 Count',
               'Date of TB Screening (yyyy-mm-dd)', 'Date of TB Sample Collection (yyyy-mm-dd)','Date of TPT Start (yyyy-mm-dd)', 
               'Date of Registration', 'Enrollment  Date (yyyy-mm-dd)', 'TPT Completion date (yyyy-mm-dd)', 
               'Date of TB Diagnostic Result Received (yyyy-mm-dd)', 'Date of Last CD4 Count',
               'Date of Viral Load Eligibility Status', 'Date of Viral Load Sample Collection (yyyy-mm-dd)',#'Date of Additional TB Diagnosis Result using XRAY (for client with negative lab results with CAD score of 40 & above)',
               'Date of Start of TB Treatment (yyyy-mm-dd)', 'Date of Completion of TB Treatment (yyyy-mm-dd)',  'Date of commencement of EAC (yyyy-mm-dd)',
               'Date of last EAC Session Completed', 'Date of Extended EAC Completion (yyyy-mm-dd)', 'Date of Repeat Viral Load - Post EAC VL Sample collected (yyyy-mm-dd)',
               'Date of Repeat Viral load result- POST EAC VL', 'Date of devolvement', 'Date of current DSD', 'Date of Return of DSD Client to Facility (yyyy-mm-dd)',
               'Date of Cervical Cancer Screening (yyyy-mm-dd)', 'Date of Precancerous Lesions Treatment (yyyy-mm-dd)', 'Date Biometrics Enrolled (yyyy-mm-dd)', 'Date Biometrics Recapture (yyyy-mm-dd)']:
    
    combined_data[col] = pd.to_datetime(combined_data[col], errors='coerce')

 
quality_issue_counts = []
line_lists = {} 

def clean_cd4(value):
    if pd.isna(value) or str(value).strip() == '':
        return np.nan

    val_str = str(value).strip().lower()

    # Handle HTML encoded less-than
    if '&lt;=' in val_str or '<=' in val_str:
        match = re.search(r'(\\d+)', val_str)
        if match:
            return float(match.group(1)) - 1  # e.g., <=200 → 199
    elif '&lt;' in val_str or '<' in val_str:
        match = re.search(r'(\\d+)', val_str)
        if match:
            return float(match.group(1)) - 1  # e.g., <200 → 199
    elif  '&gt;=' in val_str or '>=' in val_str:
        match = re.search(r'(\\d+)', val_str)
        if match:
            return float(match.group(1)) + 1  # e.g., >=200 → 201
    elif '&gt;' in val_str or '>' in val_str:
        match = re.search(r'(\\d+)', val_str)
        if match:
            return float(match.group(1)) + 1  # e.g., >200 → 201

    try:
        return float(val_str.replace(',', ''))
    except ValueError:
        return value  
    
# Apply the cleaning function to create the new column
combined_data['Cleaned Last CD4 Count'] = combined_data['Last CD4 Count'].apply(clean_cd4)


def cd4_lt(value):
    """
    Checks if a CD4 count is valid according to the following criteria:

    -   The CD4 count is a number (integer or float) less than 200.
    -   The CD4 count is a string that contains the substring "<200".

    Args:
        value: The CD4 count to validate. Can be a string or a number.

    Returns:
        bool: True if the CD4 count is valid, False otherwise.
    """
    if pd.isna(value):
        return False  # Blank is considered *invalid*

    if isinstance(value, (int, float)):
        return value < 200
    elif isinstance(value, str):
        value_str = value.strip().lower()
        if "<200" in value_str:
            return True
        try:
            numeric_value = float(value)
            return numeric_value < 200
        except ValueError:
            return False
    else:
        return False

# Define a function to check if a value is invalid
def is_invalid_cd4(value):
    # Check for blank (NaN)
    if pd.isna(value):
        return False  # Blank is valid
    
    #combined_data['Last CD4 Count'] = pd.to_numeric(combined_data['Last CD4 Count'], errors='coerce')
    #combined_data['Age'] = pd.to_numeric(combined_data['Age'], errors='coerce')
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()

    try:
        numeric_value = float(value_str)
        if 0 <= numeric_value <= 1600:
            return False  # Valid numeric value
    except ValueError:
        pass  # Not a valid number, proceed to pattern matching
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        #r"^(0|[1-9]\d{0,2}|1[0-5]\d{2}|1600)$",
        #r"^\d+$",  # Integer (e.g., 200)
        #r"^\d+\.\d+$",  # Float (e.g., 200.5)
        #r"^(?:1(?:[0-5]\d|\d{2})|\d{1,3})\.\d+$", #float less than or equal 1600
        #r"^[><]=?\d+$",  # Comparisons with numbers (e.g., >=200, <=200)
        r"^(>=200|<200|>200)$",  # Correct regex for >=200 OR <200
        r"(cells|cell/mms)"
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return False  # Value is valid
    
    # Define invalid patterns
    invalid_patterns = [
        r"[a-zA-Z]",  # Contains letters
        r"[+@`\]\^#%_,/\\|\s]",  # Contains specific invalid characters or spaces
        r"cp/ml|cell/pl|cp / ml",  # Contains invalid units
        r"[><]=?\s*\d",  # Improper comparison with numbers
        r"\d\s+\d",  # Numbers with spaces in between
        r"<|>|=",  # Standalone comparison symbols
        r"^\d{1,4}-\d{1,2}-\d{1,4}$",  # Matches date format (e.g., 2024-12-15)
        r"[.]{2,}",  # Multiple periods in a row
        r"^\d*\s+[a-zA-Z]+$",  # Number followed by a letter or word
        r"\s{2,}",  # Multiple spaces
        r"^\d*[+-/*|]$",  # Number followed by invalid operator
        r"^\d*[><=]+$",  # Number followed by comparison alone
        r"^\d+\.$",  # Number ending with a period
        r"[><]=?\d+\s*\d",  # Comparison sign followed by two numbers (e.g., <20, >30)
        r"^=>\d+$", #Comparison such as =>200
        r"-"#,
        #r"(> (?:[2-9]\d{3}|[1][7-9]\d{2}|[1][6][1-9]\d|[1][6][0][1-9]))"
    ]
    
    # Check if the value matches any invalid patterns
    for pattern in invalid_patterns:
        if re.search(pattern, value_str):
            return True  # Value is invalid
    
    # Default case: If it doesn't match any valid pattern and doesn't match specific invalid ones
    return True


def is_valid_cd4(value):
    # Check for blank (NaN)
    if pd.isna(value):
        return False  # Blank is valid
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        r"^(0|[1-9]\d{0,2}|1[0-5]\d{2}|1600)$", # integer less than or equal 1600
        #r"^\d+$",  # Integer (e.g., 200)
        #r"^\d+\.\d+$",  # Float (e.g., 200.5)
        r"^(?:1(?:[0-5]\d|\d{2})|\d{1,3})\.\d+$", #float less than or equal 1600
        #r"^[><]=?\d+$",  # Comparisons with numbers (e.g., >=200, <=200)
        r"^(>=200|<200|>200)$",  # Correct regex for >=200 OR <200
        r"(cells|cell/mms)"

    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return True  # Value is valid
        
    return True


# Clean viral load column
def clean_viral_load(value):
    if pd.isna(value) or str(value).strip() == '':
        return np.nan

    val_str = str(value).strip().lower()

    # Rule 1: Convert undetected-like values to 0
    if any(keyword in val_str for keyword in ['undetected', 'tnd', 'not detected', 'notdetected', 'nd', 'target not detected', 'not dectected', 'Not detected', 'NotDetected', 'Titer min','Titermin','TARGET0']):
        return 0

    # Rule 2: Convert <20, <30, < 20detected etc. to 0
    if re.match(r'^<\s?\d+', val_str) or '<' in val_str:
        return 0

    # Rule 3: Return plain integers
    if val_str.isdigit():
        return int(val_str)

    # Rule 4: Remove 'copies/mL' and extract number (supports float and comma)
    # match = re.match(r'^(-?[\d,]+(?:\.\d+)?)(?:\s*copies/ml)?$', val_str.replace(' ', ''))
    match = re.match(r'^(-?[\d,]+(?:\.\d+)?)(?:\s*(?:copies/ml|cp/ml|copies|Copies))$', val_str.replace(' ', ''))

    if match:
        number_str = match.group(1).replace(',', '')
        try:
            return float(number_str)
        except ValueError:
            return value

    # Rule 5: Preserve negative numbers
    if re.match(r'^-\d+(\.\d+)?$', val_str):
        try:
            return float(val_str)
        except ValueError:
            return value

    # Rule 6: Preserve specific text values
    if any(keyword in val_str for keyword in ['repeatsamplecollection', 'failedtwice', 'failedthreetimes', 'invalid', 'failed','INVALID','InvalidClot.Repeatcollection','repeat sample collection']):
        return value

    # Default: return original value
    return value

# Apply the cleaning function to create the new column
combined_data['Cleaned Current Viral Load (c/ml)'] = combined_data['Current Viral Load (c/ml)'].apply(clean_viral_load)



# Define a function to check if a value is invalid
def is_invalid_viralload(value):
     

    # Check for blank (NaN)
    if pd.isna(value):
        return True #False  # Blank is valid
    
    # Convert the value to a string for pattern matching
    value_str = str(value).strip()
    
    # Define valid patterns (expandable for additional valid values)
    valid_patterns = [
        r"^\d+$",  # Integer (e.g., 200)
        r"^\d+\.\d+$",  # Float (e.g., 200.5)
        r"^[><]?\d+$",  # Comparisons with numbers (e.g., >=200)
        r"^[<]?\d+$",  # Comparisons with numbers (e.g., <200)
        r"^<\d{2}$", #comparisons with numbers such as <20, <40
        r"(cp/ml|cell/pl|cell/mms|cp / ml|mL|Copies/mL|Copies|copies|NOT DETECT|FAILED|failed|Failed|Invalid|NotDetected|Not Detected|Not detected|cp/ml|Detected|NOT DETECTED|Target Not Detected|NOT DECTECTED|detected|undected|CP/ML|TND|TNF|ND|not det|CP|nd|NOT/D|tnd|nd|< \d+Copies/mL|TargetNotDetected)",  # Contains invalid units
        r"(< 20Copies/mL|< 30Copies/mL| Copies| copies)" , 
        r"(< \d+Copies/mL)",
        r"(< \d+detected)",
        r"^< \d+$",
        r"(<?\s?\d+Copies/mL)",
        r"[<>]?\s?\d{1,3}(?:,\d{3})*Copies/mL",
        r"< 20",
        r"< 30detected",
        r"<20",
        r"failed",
        r"\d{1,3}(,\d{3})+",
        r"Failed,TwiceREPEATSAMPLECOLLECTION|InvalidFailedtwice,REPEATSAMPLECOLLECTION|Invalid,FailedTwiceREPEATSAMPLECOLLECTION|Failedthreetimes,REPEATSAMPLECOLLECTION|Invalid,TwiceREPEATSAMPLECOLLECTION|Failedtwice,REPEATSAMPLECOLLECTION|Invalid,REPEATSAMPLECOLLECTION|Failed,REPEATSAMPLECOLLECTION|InvalidTwiceREPEATSAMPLECOLLECTION|FailedthreetimesREPEATSAMPLECOLLECTION|FailedTwiceREPEATSAMPLECOLLECTION|InvalidFailedtwiceREPEATSAMPLECOLLECTION|REPEAT SAMPLE COLLECTION|Invalid format|Invalidformat|Titer min|Titermin"
    ]
    
    # Check if the value matches any valid pattern
    for pattern in valid_patterns:
        if re.fullmatch(pattern, value_str):
            return True  #False  # Value is valid
    
    # Define invalid patterns
    invalid_patterns = [
        r"[a-zA-Z]",  # Contains letters
        r"[+@`\]\^#%_/\\|\s]",  # Contains specific invalid characters or spaces
        #r"cp/ml|cell/pl|cell/mms|cp / ml",  # Contains invalid units
        #r"[><]=?\s*\d",  # Improper comparison with numbers
        r"\d\s+\d",  # Numbers with spaces in between
        r"<|>|=|-",  # Standalone comparison symbols
        r"^\d{1,4}-\d{1,2}-\d{1,4}$",  # Matches date format (e.g., 2024-12-15)
        r"[.]{2,}",  # Multiple periods in a row
        r"^\d*\s+[a-zA-Z]+$",  # Number followed by a letter or word
        r"\s{2,}",  # Multiple spaces
        r"^\d*[+-/*|]$",  # Number followed by invalid operator
        r"^\d*[><=]+$",  # Number followed by comparison alone
        r"^\d+\.$",  # Number ending with a period
        #r"[><]=?\d+\s*\d",  # Comparison sign followed by two numbers (e.g., <20, >30)
        r"^[><]=?\d+$", # Comparisons with numbers (e.g., >=200, <=200)
        r"^=>\d+$", #Comparison such as =>200
        r"-", #comparison if it contains minus,
        r"\d+,\d{1,2}(?!\d)", #commas in the wrong place
        r"O",
        r"o",
        r"Abdulrashid Usman"

    ]
    
    # Check if the value matches any invalid patterns
    for pattern in invalid_patterns:
        if re.search(pattern, value_str):
            return False  #True  # Value is invalid
    
    # Default case: If it doesn't match any valid pattern and doesn't match specific invalid ones
    return True

# Convert Last CD4 Count and Current Viral Load to integers
# combined_data['Last CD4 Count'] = pd.to_numeric(combined_data['Last CD4 Count'], errors='coerce')
# combined_data['Current Viral Load (c/ml)'] = pd.to_numeric(combined_data['Current Viral Load (c/ml)'], errors='coerce')

# Iterate through projects
# Iterate through each unique project name
for project_name in combined_data['ProjectName'].unique():
    project_data = combined_data[combined_data['ProjectName'] == project_name].copy()

    # Create directory for the current project
    project_dir = os.path.join(output_base_dir, project_name)
    os.makedirs(project_dir, exist_ok=True)

    project_issues = {}
    all_line_lists_data = []

    # Condition 1: Blank Date Of Birth
    blank_DateOfBirth_CheckwithHI = project_data[project_data['Date of Birth (yyyy-mm-dd)'].isna()]
    project_issues['Blank Date Of Birth_CheckwithHI'] = blank_DateOfBirth_CheckwithHI.shape[0]
    blank_DateOfBirth_CheckwithHI['QualityIssue'] = 'Blank Date Of Birth_CheckwithHI'
    all_line_lists_data.append(blank_DateOfBirth_CheckwithHI)

    # Condition 2: Blank Age
    blank_age_CheckwithHI = project_data[project_data['Age'].isna()]
    project_issues['Blank Age_CheckwithHI'] = blank_age_CheckwithHI.shape[0]
    blank_age_CheckwithHI['QualityIssue'] = 'Blank Age_CheckwithHI'
    all_line_lists_data.append(blank_age_CheckwithHI)

    # Condition 3: Blank Last Pickup Date
    blank_LastPickUpDate_CheckwithHI = project_data[project_data['Last Pickup Date (yyyy-mm-dd)'].isna()]
    project_issues['Blank Last Pickup Date_CheckwithHI'] = blank_LastPickUpDate_CheckwithHI.shape[0]
    blank_LastPickUpDate_CheckwithHI['QualityIssue'] = 'Blank Last Pickup Date_CheckwithHI'
    all_line_lists_data.append(blank_LastPickUpDate_CheckwithHI)

    # Condition 4: LastPickUpDate less than ART start Date
    LastPickUpDate_less_than_ARTStartDate_CheckwithHI = project_data[
        (project_data['Last Pickup Date (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
        (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['LastPickUpDate_less_than_ARTStartDate_CheckwithHI'] = LastPickUpDate_less_than_ARTStartDate_CheckwithHI.shape[0]
    LastPickUpDate_less_than_ARTStartDate_CheckwithHI['QualityIssue'] = 'LastPickUpDate less than ART start Date_CheckwithHI'
    all_line_lists_data.append(LastPickUpDate_less_than_ARTStartDate_CheckwithHI)

    # Condition 5: Blank ART Start Date
    blank_ARTStartDate_CheckwithHI = project_data[
        (project_data['ART Start Date (yyyy-mm-dd)'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Blank ART Start Date_CheckwithHI'] = blank_ARTStartDate_CheckwithHI.shape[0]
    blank_ARTStartDate_CheckwithHI['QualityIssue'] = 'Blank ART Start Date_CheckwithHI'
    all_line_lists_data.append(blank_ARTStartDate_CheckwithHI)

    # Condition 6: Gender =' Male' and Pregnancy status = ' Pregnant' or 'Breastfeeding'
    Male_with_pregnancystatus = project_data[
        (project_data['Sex'].isin(['Male'])) & (project_data['Pregnancy Status'].isin(['Pregnant', 'Breastfeeding']))]
    project_issues['Male_with_pregnancystatus'] = Male_with_pregnancystatus.shape[0]
    Male_with_pregnancystatus['QualityIssue'] = 'Gender = Male and Pregnancy status = Pregnant/Breastfeeding'
    all_line_lists_data.append(Male_with_pregnancystatus)

    # Condition 7: Date_of start of Current Regimen less than ART start Date
    Date_start_regimen_less_than_ARTStartDate = project_data[
        (project_data['Date of Start of Current ART Regimen'] < project_data['ART Start Date (yyyy-mm-dd)']) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date_start_regimen_less_than_ARTStartDate'] = Date_start_regimen_less_than_ARTStartDate.shape[0]
    Date_start_regimen_less_than_ARTStartDate['QualityIssue'] = 'Date of start of Current Regimen less than ART start Date'
    all_line_lists_data.append(Date_start_regimen_less_than_ARTStartDate)

    # Condition 8: Blank Month of Refill with Active status
    ActiveStatus_with_blankMonthofRefill_CheckwithHI = project_data[
        (project_data['Months of ARV Refill'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['ActiveStatus_with_blankMonthofRefill_CheckwithHI'] = ActiveStatus_with_blankMonthofRefill_CheckwithHI.shape[0]
    ActiveStatus_with_blankMonthofRefill_CheckwithHI['QualityIssue'] = 'Blank Month of Refill with Active status_CheckwithHI'
    all_line_lists_data.append(ActiveStatus_with_blankMonthofRefill_CheckwithHI)

    # Condition 9: Date_of Enrollment less than Date of registration
    Date_enroll_less_than_Date_Registration = project_data[
        (project_data['Enrollment  Date (yyyy-mm-dd)'] < project_data['Date of Registration']) &
        (~(project_data['Care Entry Point'].isin(['Transfer-in','Transited']))) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date_enroll_less_than_Date_Registration'] = Date_enroll_less_than_Date_Registration.shape[0]
    Date_enroll_less_than_Date_Registration['QualityIssue'] = 'Date of Enrollment less than Date of registration'
    all_line_lists_data.append(Date_enroll_less_than_Date_Registration)

    # Condition 10: Invalid CD4 values
    cd4_invalid = project_data[
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Last CD4 Count'] >= '2018-01-01') &
        (project_data['Cleaned Last CD4 Count'].apply(is_invalid_cd4))]
    project_issues['Invalid CD4'] = cd4_invalid.shape[0]
    cd4_invalid['QualityIssue'] = 'Invalid CD4 values'
    all_line_lists_data.append(cd4_invalid)

    # Condition 11: CD4 Count for less than 5yrs
    CD4_Count_for_less_than_five_years = project_data[
        (project_data['Age'] < 5) & (project_data['Cleaned Last CD4 Count'].apply(is_valid_cd4)) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) & 
        (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['CD4_Count_for_less_than_five_years'] = CD4_Count_for_less_than_five_years.shape[0]
    CD4_Count_for_less_than_five_years['QualityIssue'] = 'CD4 Count for less than 5yrs'
    all_line_lists_data.append(CD4_Count_for_less_than_five_years)

    # Condition 12: Current ART Status "Died" but blank Cause of Death
    died_with_blank_causeOfDeath_CheckwithHI = project_data[
        (project_data['Current ART Status'] == 'Died') & project_data['Cause of Death'].isna() &
        (project_data['Previous ART Status'].isin(['Active', 'Active Restart'])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Current ART Status'] >= filter_date)]
    project_issues['Died with Blank Cause of Death'] = died_with_blank_causeOfDeath_CheckwithHI.shape[0]
    died_with_blank_causeOfDeath_CheckwithHI['QualityIssue'] = 'Died with Blank Cause of Death'
    all_line_lists_data.append(died_with_blank_causeOfDeath_CheckwithHI)

    # Condition 13: Blank Care Entry Point for ART Start Date greater than or equal to October 2024
    lamis_cut_off_date = datetime(2024, 1,1)
    blank_careEntryPoint = project_data[
        (project_data['ART Start Date (yyyy-mm-dd)'] >= lamis_cut_off_date) &
        (project_data['Care Entry Point'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Blank Care Entry Point'] = blank_careEntryPoint.shape[0]
    blank_careEntryPoint['QualityIssue'] = 'Blank Care Entry Point'
    all_line_lists_data.append(blank_careEntryPoint)

    # Condition 14: Non-integer and Invalid Current Viral Load
    viralload_invalid = project_data[
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) &
        (~project_data['Cleaned Current Viral Load (c/ml)'].apply(is_invalid_viralload))]
    project_issues['Invalid viralload'] = viralload_invalid.shape[0]
    viralload_invalid['QualityIssue'] = 'Invalid viralload'
    all_line_lists_data.append(viralload_invalid)

    # Condition 15: ART Start Date less than DateOfBirth
    ARTStartDate_Date_less_than_DateOfBirth = project_data[
        (project_data['ART Start Date (yyyy-mm-dd)'] < project_data['Date of Birth (yyyy-mm-dd)']) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['ARTStartDate_Date_less_than_DateOfBirth'] = ARTStartDate_Date_less_than_DateOfBirth.shape[0]
    ARTStartDate_Date_less_than_DateOfBirth['QualityIssue'] = 'ART Start Date less than DateOfBirth'
    all_line_lists_data.append(ARTStartDate_Date_less_than_DateOfBirth)

    # Condition 16: LastPickUpDate less than DateOfBirth
    LastPickUpDate_less_than_DateOfBirth = project_data[
        (project_data['Last Pickup Date (yyyy-mm-dd)'] < project_data['Date of Birth (yyyy-mm-dd)']) & 
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['LastPickUpDate_less_than_DateOfBirth'] = LastPickUpDate_less_than_DateOfBirth.shape[0]
    LastPickUpDate_less_than_DateOfBirth['QualityIssue'] = 'LastPickUpDate less than DateOfBirth'
    all_line_lists_data.append(LastPickUpDate_less_than_DateOfBirth)


    # Condition 17: Current ART Status Date < ART Start Date
    ART_Status_Date_less_than_ARTStartDate = project_data[
        (project_data['Date of Current ART Status'] < project_data['ART Start Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', ''])) &
          (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Current ART Status Date less than ART Start Date'] = ART_Status_Date_less_than_ARTStartDate.shape[0]
    ART_Status_Date_less_than_ARTStartDate['QualityIssue'] = 'Current ART Status Date less than ART Start Date'
    all_line_lists_data.append(ART_Status_Date_less_than_ARTStartDate)

    # Condition 18: Date of Last CD4 < ART Start Date
    # lastCD4CountDate_less_than_ARTStartDate = project_data[
    #     (project_data['Date of Last CD4 Count'] < project_data['ART Start Date (yyyy-mm-dd)']) & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    # project_issues['Date of Last CD4 less than ART Start Date'] = lastCD4CountDate_less_than_ARTStartDate.shape[0]
    # lastCD4CountDate_less_than_ARTStartDate['QualityIssue'] = 'Date of Last CD4 less than ART Start Date'
    # all_line_lists_data.append(lastCD4CountDate_less_than_ARTStartDate)

    # Condition 19: Current ART Status Date < Last Pickup Date
    ART_Status_Date_less_than_LastPickUpDate_CheckwithHI = project_data[
        (project_data['Date of Current ART Status'] < project_data['Last Pickup Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['Current ART Status Date less than Last Pickup Date_CheckwithHI'] = ART_Status_Date_less_than_LastPickUpDate_CheckwithHI.shape[0]
    ART_Status_Date_less_than_LastPickUpDate_CheckwithHI['QualityIssue'] = 'Current ART Status Date less than Last Pickup Date_CheckwithHI'
    all_line_lists_data.append(ART_Status_Date_less_than_LastPickUpDate_CheckwithHI)

    # Condition 20: Date of Previous ART_Status > Date of Current ART Status
    PreviousARTStatusDate_greater_than_CurrentARTStatusDate_CheckwithHI = project_data[
        (project_data['Confirmed Date of Previous ART Status'] > project_data['Date of Current ART Status']) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))
          & (project_data['Client Verification Outcome'].isin(['valid', '']))]
    project_issues['Date of Previous ART Status greater than Date of Current ART Status_CheckwithHI'] = PreviousARTStatusDate_greater_than_CurrentARTStatusDate_CheckwithHI.shape[0]
    PreviousARTStatusDate_greater_than_CurrentARTStatusDate_CheckwithHI['QualityIssue'] = 'Date of Previous ART Status greater than Date of Current ART Status_CheckwithHI'
    all_line_lists_data.append(PreviousARTStatusDate_greater_than_CurrentARTStatusDate_CheckwithHI)

    # Condition 21: Date of Current Viral Load (yyyy-mm-dd) without Current Viral Load (c/ml)
    
    non_numeric_texts = [
        'failed', 'repeatsamplecollection', 'failedtwice', 'failedthreetimes', 'invalid',
        'Failed,TwiceREPEATSAMPLECOLLECTION', 'InvalidFailedtwice,REPEATSAMPLECOLLECTION',
        'Invalid,FailedTwiceREPEATSAMPLECOLLECTION', 'Failedthreetimes,REPEATSAMPLECOLLECTION',
        'Invalid,TwiceREPEATSAMPLECOLLECTION', 'FailedTwice,RepeatSampleCollection',
        'Invalid,REPEATSAMPLECOLLECTION', 'REPEATSAMPLECOLLECTION', 'InvalidTwiceREPEATSAMPLECOLLECTION', 
        'FailedthreetimesREPEATSAMPLECOLLECTION', 'FailedTwiceREPEATSAMPLECOLLECTION', 'InvalidFailedtwiceREPEATSAMPLECOLLECTION',
        'REPEAT SAMPLE COLLECTION', 'Invalid format', 'Invalidformat'
    ]

    # Filter out rows where Cleaned VL is NaN and not in the list of non-numeric text values
    CurrentVLDate_without_CurrentVL = project_data[
        (project_data['Date of Current Viral Load (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Cleaned Current Viral Load (c/ml)'] == 'blank replaced with NA') & #.isna()) &
        #(~project_data['other current viral load (c/ml)'].astype(str).str.lower().isin([txt.lower() for txt in non_numeric_texts])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))
    ]


    # Condition 22: Date of Viral Load Sample Collection (yyyy-mm-dd) < ART start Date
    VLsamplecollectionDate_less_than_ARTStartDate = project_data[
        (project_data['Date of Viral Load Sample Collection (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date of Viral Load Sample Collection less than ART Start Date'] = VLsamplecollectionDate_less_than_ARTStartDate.shape[0]
    VLsamplecollectionDate_less_than_ARTStartDate['QualityIssue'] = 'Date of Viral Load Sample Collection less than ART Start Date'
    all_line_lists_data.append(VLsamplecollectionDate_less_than_ARTStartDate)

    # Condition 23: Date of Current ViralLoad Result Sample (yyyy-mm-dd) < ART start Date
    CurrentVLResultSampleDate_less_than_ARTStartDate = project_data[
        (project_data['Date of Current ViralLoad Result Sample (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date of Current ViralLoad Result Sample less than ART Start Date'] = CurrentVLResultSampleDate_less_than_ARTStartDate.shape[0]
    CurrentVLResultSampleDate_less_than_ARTStartDate['QualityIssue'] = 'Date of Current ViralLoad Result Sample less than ART Start Date'
    all_line_lists_data.append(CurrentVLResultSampleDate_less_than_ARTStartDate)

    # Condition 24: Date of Current Viral Load (yyyy-mm-dd) < ART start Date
    CurrentVLDate_less_than_ARTStartDate = project_data[
        (project_data['Date of Current Viral Load (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
          (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Date of Current Viral Load less than ART Start Date'] = CurrentVLDate_less_than_ARTStartDate.shape[0]
    CurrentVLDate_less_than_ARTStartDate['QualityIssue'] = 'Date of Current Viral Load less than ART Start Date'
    all_line_lists_data.append(CurrentVLDate_less_than_ARTStartDate)


    # Condition 25: Wrong ART Enrollment setting
    WrongARTEnrollmentSetting = project_data[
        (~project_data['ART Enrollment Setting'].isin(['Facility', 'Community'])) & (~project_data['ART Enrollment Setting'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['WrongARTEnrollmentSetting'] = WrongARTEnrollmentSetting.shape[0]
    WrongARTEnrollmentSetting['QualityIssue'] = 'Wrong ART Enrollment Setting'
    all_line_lists_data.append(WrongARTEnrollmentSetting)

    # Condition 26: Active Restart without previous current ART Status
    TX_RTT_without_PreviousARTStatus_CheckwithHI = project_data[
        (project_data['Current ART Status'].isin(['Active Restart'])) &
        (project_data['Previous ART Status'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of Current ART Status'] >= filter_date)
    ]
    project_issues['TX_RTT_without_PreviousARTStatus_CheckwithHI'] = TX_RTT_without_PreviousARTStatus_CheckwithHI.shape[0]
    TX_RTT_without_PreviousARTStatus_CheckwithHI['QualityIssue'] = 'Active Restart without previous current ART Status_CheckwithHI'
    all_line_lists_data.append(TX_RTT_without_PreviousARTStatus_CheckwithHI)

    # Condition 27: Date of TB Screening < ART Start Date
    #project_data['Last_CD4_Count'] = project_data['Last CD4 Count'].copy()
    #project_data['Last_CD4_Count'] = pd.to_numeric(project_data['Last_CD4_Count'], errors='coerce')
    TBScreeningDate_less_than_ARTStartDate = project_data[
        (project_data['Date of TB Screening (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
        (~(project_data['Date of Last CD4 Count'] < project_data['ART Start Date (yyyy-mm-dd)'])) &
        (~project_data['Cleaned Last CD4 Count'].apply(cd4_lt)) & 
        #(~project_data['Last_CD4_Count'] < 200) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['TBScreeningDate_less_than_ARTStartDate'] = TBScreeningDate_less_than_ARTStartDate.shape[0]
    TBScreeningDate_less_than_ARTStartDate['QualityIssue'] = 'Date of TB Screening less than ART Start Date'
    all_line_lists_data.append(TBScreeningDate_less_than_ARTStartDate)

    # Condition 28: Date of TB Screening with blank TB Screening Type
    TBScreeningDate_without_TBScreeningType = project_data[
        (~project_data['Date of TB Screening (yyyy-mm-dd)'].isna()) & (project_data['TB Screening Type'].isna()) &
        (project_data['TB status'] != 'Currently on TB treatment') &
        #(~(project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] < project_data['Date of TB Screening (yyyy-mm-dd)'])) &
        #(~(project_data['Date of Start of TB Treatment (yyyy-mm-dd)'] < project_data['Date of TB Screening (yyyy-mm-dd)'])) &
        (~project_data['Cleaned Last CD4 Count'].apply(cd4_lt)) & 
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (~(project_data['TB status'].isin(['Currently on TB treatment']))) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['TBScreeningDate_without_TBScreeningType'] = TBScreeningDate_without_TBScreeningType.shape[0]
    TBScreeningDate_without_TBScreeningType['QualityIssue'] = 'Date of TB Screening with blank TB Screening Type'
    all_line_lists_data.append(TBScreeningDate_without_TBScreeningType)

    # Condition 29: Date of TB Screening with blank TB Status
    TBScreeningDate_without_TBStatus = project_data[
        (~project_data['Date of TB Screening (yyyy-mm-dd)'].isna()) & (project_data['TB status'].isna()) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['TBScreeningDate_without_TBStatus'] = TBScreeningDate_without_TBStatus.shape[0]
    TBScreeningDate_without_TBStatus['QualityIssue'] = 'Date of TB Screening with blank TB Status'
    all_line_lists_data.append(TBScreeningDate_without_TBStatus)

    # Condition 30: Date of TB Sample collection for Negative TB Status
    NegativeTBStatus_with_SampleCollectionDate = project_data[
        (~(project_data['TB status'].isin(['Presumptive TB', 'Currently on TB treatment', 'Confirmed TB', '']) | project_data['TB status'].isna())) &
        (~project_data['Date of TB Sample Collection (yyyy-mm-dd)'].isna()) &
        (~project_data['Cleaned Last CD4 Count'].apply(cd4_lt)) & 
        (~project_data['Clinical Staging at Last Visit'].isin(['STAGE III', 'STAGE IV'])) &
        (~(project_data['Date of TB Sample Collection (yyyy-mm-dd)'] < project_data['Date of TB Screening (yyyy-mm-dd)'])) &
        (~(project_data['Date of TB Screening (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)'])) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
        (~(project_data['TB Diagnostic Test Type'].isin(['TB-LAM']))) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['NegativeTBStatus_with_SampleCollectionDate'] = NegativeTBStatus_with_SampleCollectionDate.shape[0]
    NegativeTBStatus_with_SampleCollectionDate['QualityIssue'] = 'Date of TB Sample collection for Negative TB Status'
    all_line_lists_data.append(NegativeTBStatus_with_SampleCollectionDate)

    # Condition 31: Date of TB Sample Collection with blank TB Diagnostic Test Type
    TBSampleCollection_without_TBDiagnosticType = project_data[
        (~project_data['Date of TB Sample Collection (yyyy-mm-dd)'].isna()) & (project_data['TB Diagnostic Test Type'].isna()) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['TBSampleCollection_without_TBDiagnosticType'] = TBSampleCollection_without_TBDiagnosticType.shape[0]
    TBSampleCollection_without_TBDiagnosticType['QualityIssue'] = 'Date of TB Sample Collection with blank TB Diagnostic Test Type'
    all_line_lists_data.append(TBSampleCollection_without_TBDiagnosticType)

    # Condition 32: Wrong format of TB Diagnostic Test Type
    Wrongformat_TBDiagnosticTestType = project_data[
        (~(project_data['TB Diagnostic Test Type'].isin(['Gene Xpert', 'TB-LAM','TB LAM','TrueNAT', 'LF-LAM', 'LF LAM','TB-LAMP','Chest X-ray', 'AFB Smear Microscopy','TB LAMP', 'COBAS', 'Cobas']) | project_data['TB Diagnostic Test Type'].isna())) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['Wrongformat_TBDiagnosticTestType'] = Wrongformat_TBDiagnosticTestType.shape[0]
    Wrongformat_TBDiagnosticTestType['QualityIssue'] = 'Wrong format of TB Diagnostic Test Type'
    all_line_lists_data.append(Wrongformat_TBDiagnosticTestType)

    # Condition 33: Date of TB Sample Collection < ART start Date
    # TBSampleCollectionDate_less_than_ARTStartDate = project_data[
    #     (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['TBSampleCollectionDate_less_than_ARTStartDate'] = TBSampleCollectionDate_less_than_ARTStartDate.shape[0]
    # TBSampleCollectionDate_less_than_ARTStartDate['QualityIssue'] = 'Date of TB Sample Collection less than ART start Date'
    # all_line_lists_data.append(TBSampleCollectionDate_less_than_ARTStartDate)

    # Condition 34: Date of TB Diagnostic Result Received < ART start Date
    # TBDiagnosticResultDate_less_than_ARTStartDate = project_data[
    #     (project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['TBDiagnosticResultDate_less_than_ARTStartDate'] = TBDiagnosticResultDate_less_than_ARTStartDate.shape[0]
    # TBDiagnosticResultDate_less_than_ARTStartDate['QualityIssue'] = 'Date of TB Diagnostic Result Received less than ART start Date'
    # all_line_lists_data.append(TBDiagnosticResultDate_less_than_ARTStartDate)


    # Condition 35: Wrong format of TB Diagnostic Result
    WrongFormatOfTBDiagnosticResult_CheckwithHI = project_data[
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= filter_date) &
        (~(project_data['TB Diagnostic Result'].isin(['MTb Not Detected','Invalid','Incomplete','Negative','Error','NotDetected','X-ray not suggestive','Not suggestive for TB','Not Detected','AFB Negative','Not detected','MTB not detected','MTB Not Detected ','-MTB NOT DETECTED','MTBD',
                                                     'MTB detected RIF resistance not detected','X-ray suggestive', 'Positive', 'MTB detected RIF resistance detected', 'MTB Detected (Rifampicin not Resistance)',
                                                     'MTB DETECTED', 'MTB DETECTED RIF Resistance Indeterminate','MTB detected RR','MTB detected RR not detected','MTB detected RR detected','MTB detected RR inderterminate','MTB Detected (Rifampicin Resistance Detected)','AFB Positive','Suggestive for TB', 'MTB TRACE DETECTED RIF INDETERMINATE', 'MTB trace RIF resistance indeterminate']) | project_data['TB Diagnostic Result'].isna())) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['WrongFormatOfTBDiagnosticResult_CheckwithHI'] = WrongFormatOfTBDiagnosticResult_CheckwithHI.shape[0]
    WrongFormatOfTBDiagnosticResult_CheckwithHI['QualityIssue'] = 'Wrong Format Of TB Diagnostic Result_CheckwithHI'
    all_line_lists_data.append(WrongFormatOfTBDiagnosticResult_CheckwithHI)



    # Condition 36: Positive TB Diagnostic Result with blank Date of start of TB Treatment
    # PositiveTBDiagnosticResult_without_DateStartofTBTreatment = project_data[
    #     (project_data['TB Diagnostic Result'].isin(['MTB detected RIF resistance not detected','MTB detected RR','MTB detected RR not detected','MTB detected RR detected','MTB detected RR inderterminate','MTB Detected (Rifampicin Resistance Detected)','X-ray suggestive','AFB Positive','Suggestive for TB','MTBD', 'Positive', 'MTB detected RIF resistance detected', 'MTB detected RR','MTB Detected (Rifampicin not Resistance)',
    #                                                 'MTB DETECTED', 'MTB DETECTED RIF Resistance Indeterminate', 'MTB TRACE DETECTED RIF INDETERMINATE', 'MTB trace RIF resistance indeterminate'])) &
    #     (project_data['Date of Start of TB Treatment (yyyy-mm-dd)'].isna()) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['PositiveTBDiagnosticResult_without_DateStartofTBTreatment'] = PositiveTBDiagnosticResult_without_DateStartofTBTreatment.shape[0]
    # PositiveTBDiagnosticResult_without_DateStartofTBTreatment['QualityIssue'] = 'Positive TB Diagnostic Result with blank Date of start of TB Treatment'
    # all_line_lists_data.append(PositiveTBDiagnosticResult_without_DateStartofTBTreatment)

    # Condition 37:Date of start of TB Treatment where TB Diagnostic Result = 'Negative' or contains 'MTB not Detected
    DateStartofTBTreatment_where_TBDiagResult_is_neg = project_data[
        (project_data['TB Diagnostic Result'].isin(['NotDetected','X-ray not suggestive','Not suggestive for TB','Not Detected','AFB Negative','Not detected','MTb Not Detected','Negative','Not detected','MTB not detected','MTB Not Detected ','-MTB NOT DETECTED'])) &
        (~project_data['Date of Start of TB Treatment (yyyy-mm-dd)'].isna()) &
        (project_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= filter_date) &
        (~(project_data['Date of Start of TB Treatment (yyyy-mm-dd)'] < project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'])) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) ]
    project_issues['DateStartofTBTreatment_where_TBDiagResult_is_neg '] = DateStartofTBTreatment_where_TBDiagResult_is_neg .shape[0]
    DateStartofTBTreatment_where_TBDiagResult_is_neg ['QualityIssue'] = 'Date of start of TB Treatment where TB Diagnostic Result = Negative or contains MTB not Detected'
    all_line_lists_data.append(DateStartofTBTreatment_where_TBDiagResult_is_neg )


    # Condition 38:Date of start of TB Treatment with blank TB Type (new, relapsed etc)
    DateStartofTBTreatment_without_TBType = project_data[
        (~project_data['Date of Start of TB Treatment (yyyy-mm-dd)'].isna()) &
        (project_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= filter_date) &
        (project_data['TB Type (new, relapsed etc)'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['DateStartofTBTreatment_without_TBType'] = DateStartofTBTreatment_without_TBType.shape[0]
    DateStartofTBTreatment_without_TBType['QualityIssue'] = 'Date of start of TB Treatment with blank TB Type (new, relapsed etc)'
    all_line_lists_data.append(DateStartofTBTreatment_without_TBType)


    # Condition 39: Date of Completion of TB Treatment < Date of Start of TB Treatment
    # DateCompletionofTBTreatment_less_than_DateStartofTBTreatment = project_data[
    #     (project_data['Date of Completion of TB Treatment (yyyy-mm-dd)'] < project_data['Date of Start of TB Treatment (yyyy-mm-dd)']) &
    #     (~project_data['Date of Start of TB Treatment (yyyy-mm-dd)'].isna()) &
    #     (project_data['Date of Start of TB Treatment (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['DateCompletionofTBTreatment_less_than_DateStartofTBTreatment'] = DateCompletionofTBTreatment_less_than_DateStartofTBTreatment.shape[0]
    # DateCompletionofTBTreatment_less_than_DateStartofTBTreatment['QualityIssue'] = 'Date of Completion of TB Treatment less than Date of Start of TB Treatment'
    # all_line_lists_data.append(DateCompletionofTBTreatment_less_than_DateStartofTBTreatment)


    # Condition 40:Date of Completion of TB Treatment with blank TB Treatment Outcome
    DateCompletionofTBTreatment_without_TBTreatmentOutcome = project_data[
        (~project_data['Date of Completion of TB Treatment (yyyy-mm-dd)'].isna()) &
        (project_data['TB Treatment Outcome'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['DateCompletionofTBTreatment_without_TBTreatmentOutcome'] = DateCompletionofTBTreatment_without_TBTreatmentOutcome.shape[0]
    DateCompletionofTBTreatment_without_TBTreatmentOutcome['QualityIssue'] = 'Date of Completion of TB Treatment with blank TB Treatment Outcome'
    all_line_lists_data.append(DateCompletionofTBTreatment_without_TBTreatmentOutcome)

    # Condition 41:Date of TPT Start where TB Status is positive.
    # DateTPTStart_where_TBStatus_is_positive_CheckwithHI = project_data[
    #     (((project_data['TB status'].isin(['Presumptive TB'])) & (project_data['TB Diagnostic Result'].isin(['MTB detected RIF resistance not detected', 'Positive', 'MTB detected RIF resistance detected', 'MTB Detected (Rifampicin not Resistance)', 'MTB DETECTED', 'MTB DETECTED RIF Resistance Indeterminate', 'MTB TRACE DETECTED RIF INDETERMINATE', 'MTB trace RIF resistance indeterminate']))) |
    #      (project_data['TB status'].isin(['Currently on TB treatment', 'Confirmed TB']))) &
    #     (~project_data['Date of TPT Start (yyyy-mm-dd)'].isna()) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Date of TPT Start (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= filter_date) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['DateTPTStart_where_TBStatus_is_positive_CheckwithHI'] = DateTPTStart_where_TBStatus_is_positive_CheckwithHI.shape[0]
    # DateTPTStart_where_TBStatus_is_positive_CheckwithHI['QualityIssue'] = 'Date of TPT Start where TB Status is positive'
    # all_line_lists_data.append(DateTPTStart_where_TBStatus_is_positive_CheckwithHI)


    # Condition 42:Date of TPT Start with blank TPT Type
    DateTPTStart_without_TPTtype = project_data[
        (project_data['Date of TPT Start (yyyy-mm-dd)'] >= filter_date) &
        (~project_data['Date of TPT Start (yyyy-mm-dd)'].isna()) &
        (project_data['TPT Type'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['DateTPTStart_without_TPTtype'] = DateTPTStart_without_TPTtype.shape[0]
    DateTPTStart_without_TPTtype['QualityIssue'] = 'Date of TPT Start with blank TPT Type'
    all_line_lists_data.append(DateTPTStart_without_TPTtype)


    # Condition 43: Date of TPT Completion < Date of TPT Start
    DateTPTCompletion_less_than_DateTPTStart_CheckwithHI = project_data[
        (project_data['TPT Completion date (yyyy-mm-dd)'] < project_data['Date of TPT Start (yyyy-mm-dd)']) &
        (project_data['TPT Completion date (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Date of TB Screening (yyyy-mm-dd)'] >= filter_date) &
        (~project_data['Date of TPT Start (yyyy-mm-dd)'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['DateTPTCompletion_less_than_DateTPTStart_CheckwithHI'] = DateTPTCompletion_less_than_DateTPTStart_CheckwithHI.shape[0]
    DateTPTCompletion_less_than_DateTPTStart_CheckwithHI['QualityIssue'] = 'Date of TPT Completion less than Date of TPT Start_CheckwithHI'
    all_line_lists_data.append(DateTPTCompletion_less_than_DateTPTStart_CheckwithHI)


    # Condition 44:Date of TPT Completion with blank TPT Completion status
    DateTPTCompletion_without_CompletionStatus = project_data[
        (~project_data['TPT Completion date (yyyy-mm-dd)'].isna()) &
        (project_data['TPT Completion status'].isna()) &
        (project_data['TPT Completion date (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['DateTPTCompletion_without_CompletionStatus'] = DateTPTCompletion_without_CompletionStatus.shape[0]
    DateTPTCompletion_without_CompletionStatus['QualityIssue'] = 'DateTPTCompletion_without_CompletionStatus'
    all_line_lists_data.append(DateTPTCompletion_without_CompletionStatus)
    
    
    # 'TB Diagnostic Result' to lowercase for case-insensitive matching
    project_data['TB Diagnostic Result'] = project_data['TB Diagnostic Result'].str.lower().fillna('')

    pos_neg_pattern = r'pos|neg|\+|\-|\_|\+ve|\-ve|nag|p0s|nrg|pso|ng|ned'
    chest_x_ray_pos_neg_pattern = r'not sugestive|suggestive|mbt detectected|mt detected|mt not detected|mtb n0t detected|dectected|mtb not detectd|not detected|detected|dtected|detectted|detectd|dedected|detect|mtbd|deteted|dectected|\-mtb|suggestive for tb|mtb not detected'
    mtb_pos_neg_pattern = r'mbt detectected|mt detected|mt not detected|mtb n0t detected|dectected|mtb not detectd|not detected|detected|dtected|detectted|detectd|dedected|detect|mtbd|deteted|mtb trace|dectected|\-mtb|error|incomplete|invalid|mtb not detected|mtb detected (rifampicin resistance detected)|mtb trace rif resistance indeterminate|not detected'

    tb_exclusion_pattern = r'(?:tb\s+(?:positive|negative|pos|neg|ned))'
    afb_exclusion_pattern = r'(?:afb\s+(?:positive|negative|pos|neg|ned))'

    pos_pattern = r'pos|\+|\+ve|p0s|pso'
    chest_x_ray_pos_pattern = r'^(suggestive|x-ray suggestive|mtb detected|mbt detectected|mt detected|detected|mtbdetect|mtb detectected|mtb detectted|detectted|mtb detectd|detectd|mtb dectected|dectected|mtb dedected|dedected|mtb dtected|dtected|mtd detected|mtb detectted|detectted|mtbd|mtb trace|detect|ptb detect|dedected|ptb suspect|deteted|\+mtb)'
    mtb_pos_pattern = r'^(mtb detected|mbt detectected|mt detected|detected|mtbdetect|mtb detectected|mtb detectted|detectted|mtb detectd|detectd|mtb dectected|dectected|mtb dedected|dedected|mtb trace|mtb dtected|dtected|mtd detected|mtb detectted|detectted|mtbd|mtb trace|detect|ptb detect|dedected|ptb suspect|deteted|\+mtb)'

    tb_pos_exclusion_pattern = r'(?:TB\s+(?:positive|pos|p0s))'
    afb_pos_exclusion_pattern = r'(?:AFB\s+(?:positive|pos|p0s))'


    
    End_of_quarter = pd.to_datetime('2026-03-31')
    six_months_ago = pd.to_datetime('2025-10-01')


    tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('Gene Xpert', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(mtb_pos_neg_pattern, na=False)))
    )
    ]
    project_issues['tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_Xpert_with_result_mappingissues)





    tx_tb_d_TB_Test_Type_TruNAT_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('TrueNAT', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(mtb_pos_neg_pattern, na=False)))
    )
    ]
    project_issues['tx_tb_d_TB_Test_TruNAT_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_TruNAT_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_TruNAT_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_TruNAT_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_TruNAT_with_result_mappingissues)




    tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('Chest X-ray', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(chest_x_ray_pos_neg_pattern, na=False)))  
    )
    ]
    project_issues['tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_Xray_with_result_mappingissues)



    tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('AFB Smear Microscopy', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) #&
    #(project_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False))
    ) 
    )
    ]
    project_issues['tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_AFB_with_result_mappingissues)







    tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('Clinical evaluation only', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)) &
    (project_data['TB Diagnostic Result'].str.contains(tb_exclusion_pattern, na=False)) &
    (project_data['TB Diagnostic Result'].str.contains(afb_exclusion_pattern, na=False))) 
    )
    ]
    project_issues['tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_Clinical_with_result_mappingissues)









    tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues = project_data[
   ((project_data['Date of TB Screening (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Screening (yyyy-mm-dd)'] <= End_of_quarter)) & 
   ((project_data['Client Verification Outcome'].isin(['valid', 'valid ', ' valid', 'Valid'])) |
     (project_data['Client Verification Outcome'].isna()) | # Catches np.nan
        (project_data['Client Verification Outcome'] == '') | # Catches explicit empty strings
        (project_data['Client Verification Outcome'].str.strip() == '') ) &
   (project_data['Current ART Status'].str.contains('Active')) &
   ((project_data['TB Screening Type'].isin(['CXR', 'Smear', 'Gene Xpert', 'Chest X-ray without CAD', 'Chest X-ray','Chest X-ray with CAD', 'Chest X-Ray with CAD and/or Symptom screening', 'Symptom screen (alone)']))) &
    (project_data['TB status'].str.contains('Presumptive TB|TB Suspected and referred for evaluation|Confirmed TB|Currently on TB treatment')) &
   ((project_data['Date of TB Sample Collection (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Sample Collection (yyyy-mm-dd)'] <= End_of_quarter)) &
   (~project_data['TB Diagnostic Test Type'].isna()) &
   ((project_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] >= six_months_ago) & (combined_data['Date of TB Diagnostic Result Received (yyyy-mm-dd)'] <= End_of_quarter)) &
   (
         
    ((project_data['TB Diagnostic Test Type'].str.contains('TB-LAM', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False)))  |

    ((project_data['TB Diagnostic Test Type'].str.contains('LF-LAM|TB LAMP', na=False)) &
    ~(project_data['TB Diagnostic Result'].str.contains(pos_neg_pattern, na=False))) 
    )
    ]
    project_issues['tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues'] = tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues.shape[0]
    tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues['QualityIssue'] = 'tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues'
    all_line_lists_data.append(tx_tb_d_TB_Test_Type_LAM_with_result_mappingissues)










    # Condition 45: Date of Cervical Cancer Screening less than ART start Date
    # DateCXCA_Screening_less_than_ARTStartDate = project_data[
    #     (project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] < project_data['ART Start Date (yyyy-mm-dd)']) &
    #     (project_data['Client Verification Outcome'].isin(['valid', ''])) &
    #     (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    # project_issues['DateCXCA_Screening_less_than_ARTStartDate'] = DateCXCA_Screening_less_than_ARTStartDate.shape[0]
    # DateCXCA_Screening_less_than_ARTStartDate['QualityIssue'] = 'DateCXCAScreening_less_than_ARTStartDate'
    # all_line_lists_data.append(DateCXCA_Screening_less_than_ARTStartDate)


    # Condition 46:Date of Cervical Cancer Screening with blank Cervical Cancer Screening Type
    CXCAScreening_with_blank_Screentype = project_data[
        (project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= filter_date) &
        (~project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'].isna()) &
        (project_data['Cervical Cancer Screening Type'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['CXCAScreening_with_blank_Screentype'] = CXCAScreening_with_blank_Screentype.shape[0]
    CXCAScreening_with_blank_Screentype['QualityIssue'] = 'CXCAScreening_with_blank_Screentype'
    all_line_lists_data.append(CXCAScreening_with_blank_Screentype)

    # Condition 47:Date of Cervical Cancer Screening with blank Cervical Cancer Screening Method
    CXCAScreening_with_blank_Screenmethod = project_data[
        (project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= filter_date) &
        (~project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'].isna()) &
        (project_data['Cervical Cancer Screening Method'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart'])) ]
    project_issues['CXCAScreening_with_blank_Screenmethod'] = CXCAScreening_with_blank_Screenmethod.shape[0]
    CXCAScreening_with_blank_Screenmethod['QualityIssue'] = 'CXCAScreening_with_blank_Screenmethod'
    all_line_lists_data.append(CXCAScreening_with_blank_Screenmethod)


    # Condition 48:Result of Cervical Cancer Screening = 'positive' with blank Date of Precancerous Lesions Treatment
    PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate = project_data[
        (project_data['Date of Cervical Cancer Screening (yyyy-mm-dd)'] >= filter_date) &
        (project_data['Result of Cervical Cancer Screening'].isin(['Positive'])) &
        (project_data['Date of Precancerous Lesions Treatment (yyyy-mm-dd)'].isna()) &
        (project_data['Client Verification Outcome'].isin(['valid', ''])) &
        (project_data['Current ART Status'].isin(['Active', 'Active Restart']))]
    project_issues['PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate'] = PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate.shape[0]
    PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate['QualityIssue'] = 'PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate'
    all_line_lists_data.append(PositiveCXCAScreeningResult_with_blank_CXCATreatmentdate)

    

    # Create DataFrame for transposed quality issues summary for the current IP
    quality_issues_df_transposed = pd.DataFrame.from_dict(project_issues, orient='index', columns=['Number of Records']).reset_index()
    quality_issues_df_transposed.columns = ['Quality Issue', 'Number of Records']

    # Combine all line list data for the current IP into one DataFrame
    all_line_lists_df = pd.concat(all_line_lists_data, ignore_index=True)

    # Define the output file path for the current IP
    output_file_path = os.path.join(project_dir, f"{project_name}_Quality_Check.xlsx")

    # Save the transposed quality issues and the combined line list to one Excel file with two sheets
    with pd.ExcelWriter(output_file_path, engine='openpyxl') as writer:
        quality_issues_df_transposed.to_excel(writer, sheet_name='Quality Issues Summary', index=False)
        all_line_lists_df.to_excel(writer, sheet_name='Quality Issues Line List', index=False)

    print(f"Quality check for project '{project_name}' saved to: {output_file_path}") 


# Save original and cleaned viral load data to a new file
viral_load_columns = ['Current Viral Load (c/ml)', 'Cleaned Current Viral Load (c/ml)']
if all(col in combined_data.columns for col in viral_load_columns):
    filtered_data = combined_data[combined_data['Date of Current Viral Load (yyyy-mm-dd)'] > filter_date].copy()
    viral_load_df = filtered_data[viral_load_columns].copy()
    distinct_viral_load_df = viral_load_df.drop_duplicates()

    try:
        distinct_viral_load_df.to_excel(viral_load_output_path, index=False)
        print(f"\nSuccessfully saved distinct viral load data to: {viral_load_output_path}")
    except Exception as e:
        print(f"\nError saving viral load data: {e}")
else:
    print("\nWarning: 'Current Viral Load (c/ml)' or 'Cleaned Current Viral Load (c/ml)' column not found. Skipping viral load data export.")


# Save original and cleaned Last cd4 count data to a new file
last_cd4_count_columns = ['Last CD4 Count', 'Cleaned Last CD4 Count']
if all(col in combined_data.columns for col in last_cd4_count_columns):
    filtered_data = combined_data[combined_data['Date of Last CD4 Count'] > filter_date].copy()
    cd4_count_df = filtered_data[last_cd4_count_columns].copy()
    distinct_cd4_count_df = cd4_count_df.drop_duplicates()

    try:
        distinct_cd4_count_df.to_excel(unique_cd4_output_path, index=False)
        print(f"\nSuccessfully saved distinct cd4 count data to: {unique_cd4_output_path}")
    except Exception as e:
        print(f"\nError saving cd4 count data: {e}")
else:
    print("\nWarning: 'Last CD4 Count' or 'Cleaned Last CD4 Count' column not found. Skipping cd4 count data export.")
