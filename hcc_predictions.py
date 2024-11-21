import pandas as pd
import xlwings as xw
import numpy as np
from datetime import datetime, timedelta

#Define the function that filters dataset by procedure code and predicts cost for the next 12 months
def procedurePredictions(sheet_name, new_sheet_name, procedure_code):
    wb = xw.Book.caller()
    sheet = wb.sheets[sheet_name]

    # read hcc_medical worksheet
    medical = sheet['A1'].expand('table').options(pd.DataFrame, header=1, index=False).value

    # rename columns
    medical.rename(columns={'Primary Procedure Description': 'Procedure', 'Service Date': 'ServiceDate', 
                            'Sum Employer Paid Amount (Med)': 'Med_Employer_Paid', 'Member ID': 'Member_ID', 
                            'ICD-10 Diagnosis Description (Primary)': 'ICD-10_Diagnosis', 'Provider Name': 'ProviderName', 
                            'CPT / HCPCS Procedure Code': 'Procedure_Code', 'Included Emergency Room Visit ': 'Is_ER_Visit', 
                            'Included Inpatient Admit': 'Is_Admit', 'Sum Inpatient Days': 'Inpatient_Days', 'Paid Date': 'PaidDate', 
                            'Revenue Code': 'RevCode' }, inplace=True)

    # change employer paid to numeric and dates to dateime
    medical[['ServiceDate', 'PaidDate']] = medical[['ServiceDate', 'PaidDate']].apply(pd.to_datetime)
    medical[['Med_Employer_Paid']] = medical[['Med_Employer_Paid']].apply(pd.to_numeric)
    medical['Procedure_Code'] = medical['Procedure_Code'].astype(str).str.strip().str.replace('.0','', regex=False)

    # filter for procedure
    def process_data(procedure_code):
        if len(medical[medical['Procedure_Code'] == procedure_code].values) > 0:
            procedure = medical[medical['Procedure_Code'] == procedure_code].copy()
            continue_processing(procedure)
        else:
            move_to_next_function()
    
    def move_to_next_function():
        # Logic for the next function
        print('Procedure not found. Moving to the next function.')

    def continue_processing(procedure):
        #Find the last 3 service dates by member ID
        last_three_dates = procedure.groupby('Member_ID')['ServiceDate'].apply(
            lambda x: x.nlargest(3))
        last_three_dates = last_three_dates.reset_index()

        #Define the function and Calculate the average of days between the last 3 service dates
        def calculate_date_differences(dates):
            if len(dates) < 3:
                return None  # Not enough dates to calculate differences
            # Calculate differences in days
            differences = (dates.diff().dt.days[1:]).dropna()  # Skip the first NaT
            return differences.mean()
        
        #Define the variable for finding the average number of days between last 3 injections/service dates
        average_days = last_three_dates.groupby('Member_ID')['ServiceDate'].apply(calculate_date_differences).abs().reset_index(name='Average_Days')

        #Create new column for the predicted number of doses a year
        average_days['procedures_year'] = 365/average_days['Average_Days']

        # Group by 'member_id' and get the earliest date
        earliest_dates = procedure.groupby('Member_ID')['ServiceDate'].min().reset_index()
        latest_dates = procedure.groupby('Member_ID')['ServiceDate'].max().reset_index()

        # Rename the columns for clarity
        earliest_dates.columns = ['Member_ID', 'earliest_date']
        latest_dates.columns = ['Member_ID', 'last_date']
        
        members_procedure = procedure.groupby('Member_ID').agg(
            Procedure_Cost = ('Med_Employer_Paid', 'sum'),
            Procedure_Counts = ('Procedure_Code' , 'count')
        ).reset_index()

        members_procedure['Average_Cost_Per_Procedure'] = members_procedure['Procedure_Cost']/members_procedure['Procedure_Counts']
        members_procedure_df = pd.merge(members_procedure, average_days, on='Member_ID').merge(earliest_dates, on='Member_ID').merge(latest_dates, on='Member_ID')
        
        #calculate the difference between today's date and last reported service date
        today = pd.Timestamp.today()
        last_date = members_procedure_df['last_date']
        difference = (today - last_date)/ np.timedelta64(1, 'D')
        members_procedure_df['Diff_From_Todays_Date_to_Last_Date'] = difference

        #Predict 12 month cost only if there is less than 275 days since the last service date
        members_procedure_df['Predict_12_Months'] = members_procedure_df.apply(
            lambda row: row['Average_Cost_Per_Procedure'] * row['procedures_year'] 
            if row['Diff_From_Todays_Date_to_Last_Date'] < 275 else 0, axis=1
        )

        member_predictions = members_procedure_df

    #Write results to a new sheet
        try: 
            new_sheet = wb.sheets[new_sheet_name]
            new_sheet.delete()  # Delete the existing sheet if it exists
        except:
            pass  # If the sheet does not exist, we will create it

        new_sheet = wb.sheets.add(new_sheet_name)  # Create a new sheet
        print(new_sheet)
        new_sheet.range("A1").options(index=False).value =  member_predictions  # Write DataFrame to the new sheet
        
    
    process_data(procedure_code)

if __name__ == "__main__":
    xw.Book("HCC_Predictions.xlsm").set_mock_caller()  # Set the mock caller for testing

# Predictions for RX data
def drugPredictions(sheet_name, new_sheet_name, drugName):
    wb = xw.Book.caller()
    sheet = wb.sheets[sheet_name]

    # read hcc_medical worksheet
    rx = sheet['A1'].expand('table').options(pd.DataFrame, header=1, index=False).value

    # rename columns
    rx.rename(columns={'Preferred Drug Name (Artemis)': 'DrugName', 'Service Date': 'ServiceDate', 'Sum Employer Paid Amount (Rx)': 'Rx_Employer_Paid', 'Member ID': 'Member_ID', 'NDC Code': 'NDC_Code', 'Provider Name': 'ProviderName', 'Sum Days Supply': 'DaysSupply', 'Specialty Drug Indicator (HCG)': 'Is_SpecialtyDrug', 'Paid Date': 'PaidDate', 'Sum Rx Scripts (HCG)': 'RxScripts'}, inplace=True)

    # change employer paid to numeric and dates to dateime
    rx[['ServiceDate', 'PaidDate']] = rx[['ServiceDate', 'PaidDate']].apply(pd.to_datetime)
    rx[['Rx_Employer_Paid']] = rx[['Rx_Employer_Paid']].apply(pd.to_numeric)

    # filter for drugName
    
    def process_data(drugName):
        if len(rx[rx['DrugName'].str.contains(drugName, case=False, na=False)].values) > 0:
            drug = rx[rx['DrugName'].str.contains(drugName, case=False, na=False)].copy()
            continue_processing(drug)
        else:
            move_to_next_function()
    
    def move_to_next_function():
        # Logic for the next function
        print("Drug name not found. Moving to the next function.")

    def continue_processing(drug):
        #Find the last 3 service dates by member ID
        last_three_dates = drug.groupby('Member_ID')['ServiceDate'].apply(
            lambda x: x.nlargest(3))
        last_three_dates = last_three_dates.reset_index()

        #Define the function and Calculate the average of days between the last 3 service dates
        def calculate_date_differences(dates):
            if len(dates) < 3:
                return None  # Not enough dates to calculate differences
            # Calculate differences in days
            differences = (dates.diff().dt.days[1:]).dropna()  # Skip the first NaT
            return differences.mean()
        
        #Define the variable for finding the average number of days between last 3 injections/service dates
        average_days = last_three_dates.groupby('Member_ID')['ServiceDate'].apply(calculate_date_differences).abs().reset_index(name='Average_Days')

        #Create new column for the predicted number of doses a year
        average_days['doses_year'] = 365/average_days['Average_Days']
        

        # Group by 'member_id' and get the earliest date
        earliest_dates = drug.groupby('Member_ID')['ServiceDate'].min().reset_index()
        latest_dates = drug.groupby('Member_ID')['ServiceDate'].max().reset_index()

        # Rename the columns for clarity
        earliest_dates.columns = ['Member_ID', 'earliest_date']
        latest_dates.columns = ['Member_ID', 'last_date']

        members_drug = drug.groupby('Member_ID').agg(
            Drug_Cost = ('Rx_Employer_Paid', 'sum'),
            Script_Counts = ('RxScripts' , 'count')
        ).reset_index()

        members_drug['Average_Cost_Per_Script'] = members_drug['Drug_Cost']/members_drug['Script_Counts']

        members_drug_df = pd.merge(members_drug, average_days, on='Member_ID').merge(earliest_dates, on='Member_ID').merge(latest_dates, on='Member_ID')
        
        #calculate the difference between today's date and last reported service date
        today = pd.Timestamp.today()
        last_date = members_drug_df['last_date']
        difference = (today - last_date)/ np.timedelta64(1, 'D')
        members_drug_df['Diff_From_Todays_Date_to_Last_Date'] = difference

        #Predict 12 month cost only if there is less than 275 days since the last service date
        members_drug_df['Predict_12_Months'] = members_drug_df.apply(
            lambda row: row['Average_Cost_Per_Script'] * row['doses_year'] 
            if row['Diff_From_Todays_Date_to_Last_Date'] < 275 else 0, axis=1
        )
        
        drug_predictions = members_drug_df

    #Write results to a new sheet
        try: 
            new_sheet = wb.sheets[new_sheet_name]
            new_sheet.delete()  # Delete the existing sheet if it exists
        except:
            pass  # If the sheet does not exist, we will create it

        new_sheet = wb.sheets.add(new_sheet_name)  # Create a new sheet
        print(new_sheet)
        new_sheet.range("A1").options(index=False).value =  drug_predictions  # Write DataFrame to the new sheet

    process_data(drugName)

if __name__ == "__main__":
    xw.Book("HCC_Predictions.xlsm").set_mock_caller()  # Set the mock caller for testing


# Run RX predictive functions
drugPredictions('hcc_rx', 'humira_predictions', 'humira')
drugPredictions('hcc_rx', 'dupixent_predictions', 'dupixent')
drugPredictions('hcc_rx', 'ovidrel_predictions', 'ovidrel')
drugPredictions('hcc_rx', 'skyrizi_predictions', 'skyrizi')
drugPredictions('hcc_rx', 'stelera_predictions', 'stelera')

# Run Med predictive functions
procedurePredictions('hcc_medical', 'keytruda_predictions', 'J9271')
procedurePredictions('hcc_medical', 'initialchemo_predictions', '96413')
procedurePredictions('hcc_medical', 'hemodialysis_predictions', '90935')
