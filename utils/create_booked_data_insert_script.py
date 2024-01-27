import pandas as pd
from datetime import datetime
import os


def read_excel_file(spreadsheet, *tabnames):

    try:
        if len(tabnames) < 2:
            print("Please provide at least two sheet names.")
            return 400, "Please provide at least two sheet names."

        dfs = []
        for tabname in tabnames:
            df = pd.read_excel(
                f"./ExcelData/{spreadsheet}", sheet_name=tabname, skiprows=4, usecols="A:S", nrows=12)
            df["tabname"] = tabname
            df["CompanyName"] = "XYZ"
            df = df.drop(df.columns[2:15], axis=1)
            dfs.append(df)
        return 200, dfs
    except Exception as e:
        return 500, f"Error reading Excel file: {e}"


def process_data(dfs):
    try:
        # Concatenate the dataframes
        df_combined_booked_data = pd.concat(dfs, ignore_index=True)

        # Dropping rows with all NaN values
        df_combined_booked_data = df_combined_booked_data.dropna(
            axis=0, how="all")

        # Adding the remaining columns to the dataframe
        df_combined_booked_data['LineOfBusiness'] = df_combined_booked_data.apply(
            lambda row: 'General Liability' if 'GL' in row['tabname'] else 'Motor/Accident', axis=1)
        df_combined_booked_data['Currency'] = 'EUR'
        df_combined_booked_data['DWCreatedDate'] = datetime.today().strftime(
            '%Y-%m-%d')
        df_combined_booked_data['DWCreatedBy'] = 'Christian'

        # Calculating the Ultimate Loss Ratio
        df_combined_booked_data["UltimateLossRatio"] = df_combined_booked_data["Paid losses"] + \
            df_combined_booked_data["Case reserves"] + \
            df_combined_booked_data["IBNR"]

        # Renaming the columns to match those of the target table
        df_combined_booked_data = df_combined_booked_data.rename(columns={
            "Earned premium": "EarnedPremium",
            "Paid losses": "PaidLosses",
            "Case reserves": "CaseReserves",
            "U/W year": "Year",
            "Gross written premium": "GrossWrittenPremium",
        })
        return 200, df_combined_booked_data

    except Exception as e:
        print(f"Error processing data: {e}")
        return 500, f"Error processing data: {e}"


def create_insert_statements(df_combined_booked_data):
    try:
        # Creating the insert statements
        insert_statements = []
        for index, row in df_combined_booked_data.iterrows():
            values = f"{row['LineOfBusiness']}, {row['Year']}, {row['GrossWrittenPremium']}, {row['EarnedPremium']}, {row['PaidLosses']}, {row['CaseReserves']}, {row['IBNR']}, {row['UltimateLossRatio']}, \"{row['DWCreatedDate']}\", \"{row['DWCreatedBy']}\""
            insert_statement = f"INSERT INTO [ClaimDevelopment].[FactData] ([LineOfBusiness], [Year], [GrossWrittenPremium], [EarnedPremium], [PaidLosses], [CaseReserves], [IBNR], [UltimateLossRatio], [DWCreatedDate], [DWCreatedBy]) VALUES ({values});"
            insert_statements.append(insert_statement)

        # Creating the folder if it doesn't exist
        folder_path = "./InsertScripts"
        os.makedirs(folder_path, exist_ok=True)
        file_path = os.path.join(folder_path, "booked_data_insert_script.sql")

        # Writing the insert statements to a file
        with open(file_path, "w") as sql_file:
            for statement in insert_statements:
                sql_file.write(statement + "\n")

        print("Script created successfully")
        return 200, file_path
    except Exception as e:
        return 500, f"Error creating insert statements: {e}"


def main(spreadsheet, *tabnames):
    status, dfs = read_excel_file(spreadsheet, *tabnames)
    if status != 200:
        return f"Error reading Excel file: {dfs}"

    status, df_combined_booked_data = process_data(dfs)
    if status != 200:
        return f"Error processing data: {df_combined_booked_data}"

    status, file_path = create_insert_statements(df_combined_booked_data)
    if status != 200:
        return f"Error creating the insert script: {file_path}"

    print("Data processing and script creation completed successfully.")


if __name__ == "__main__":
    main("Howden_CompanyXYZ_2021_Data.xlsx", "GL-np", "MA-np")
