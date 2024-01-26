import pandas as pd
from datetime import datetime
import os


def read_excel_file(spreadsheet, *tabnames):

    try:
        if len(tabnames) < 2:
            return {400, "Please provide at least two sheet names."}

        dfs = []
        for tabname in tabnames:
            df = pd.read_excel(
                f"./ExcelData/{spreadsheet}", sheet_name=tabname, skiprows=4, usecols="A:N", nrows=12)
            # Adding identifier column
            df["tabname"] = tabname
            # Dropping the Gross written premium column
            df = df.drop(df.columns[1], axis=1)
            dfs.append(df)
        return dfs
    except Exception as e:
        return {500, f"Error reading Excel file: {e}"}


def process_data(dfs):
    try:
        # Concatenate the dataframes
        df_combined = pd.concat(dfs, ignore_index=True)

        df_melted = pd.melt(df_combined, id_vars=[
                            "U/W year", "tabname"], var_name="Development Period", value_name="Loss Ratio")

        # Changing NaN values to 0 to avoid errors when converting to integers
        df_melted["U/W year"] = df_melted["U/W year"].fillna(0)
        df_melted["Development Period"] = df_melted["Development Period"].fillna(
            0)
        df_melted["Loss Ratio"] = df_melted["Loss Ratio"].fillna(0)

        # Converting the relevant columns to integers
        df_melted["U/W year"] = df_melted["U/W year"].astype(int)
        df_melted["Development Period"] = df_melted["Development Period"].astype(
            int)

        # Renaming the columns to match those of the target table
        df_melted = df_melted.rename(columns={"U/W year": "Year", "Gross written premium": "GrossWrittenPremium",
                                     "Development Period": "DevelopmentMonth", "Loss Ratio": "LossIncurredRatio"})

        # Adding the remaining columns to df_melted
        df_melted["CompanyName"] = "XYZ"
        df_melted["LineOfBusiness"] = df_melted.apply(
            lambda row: "General Liability" if "GL" in row["tabname"] else "Motor/Accident", axis=1)
        df_melted["Currency"] = "EUR"
        df_melted["DWCreatedDate"] = datetime.today().strftime("%Y-%m-%d")
        df_melted["DWCreatedBy"] = "Christian"

        return df_melted

    except Exception as e:
        return {500, f"Error processing data: {e}"}


def create_insert_statements(df_melted):
    try:
        # Creating the insert statements
        insert_statements = []
        for index, row in df_melted.iterrows():
            values = f'"{row["CompanyName"]}", "{row["LineOfBusiness"]}", "{row["Currency"]}", {row["Year"]}, {row["LossIncurredRatio"]}, {row["DevelopmentMonth"]}, "{row["DWCreatedDate"]}", "{row["DWCreatedBy"]}"'
            insert_statement = f'INSERT INTO [ClaimDevelopment].[FactStatistical] ([CompanyName], [LineOfBusiness], [Currency], [Year], [LossIncurredRatio], [DevelopmentMonth], [DWCreatedDate], [DWCreatedBy]) VALUES ({values});'
            insert_statements.append(insert_statement)

        # Writing the insert statements to a file in the "InsertScripts" folder
        folder_path = "./InsertScripts"
        os.makedirs(folder_path, exist_ok=True)
        file_path = os.path.join(
            folder_path, "insert_script_statistical_data.sql")

        # Writing the insert statements to a file
        with open(file_path, "w") as sql_file:
            for statement in insert_statements:
                sql_file.write(statement + "\n")

        print("Script created successfully")
        return 200

    except Exception as e:
        # Internal server error status code
        return {500, f"Error creating the insert script {e}"}


def main(spreadsheet, *tabnames):
    dfs = read_excel_file(spreadsheet, *tabnames)
    df_melted = process_data(dfs)
    create_insert_statements(df_melted)


if __name__ == "__main__":
    main("Howden_CompanyXYZ_2021_Data.xlsx", "GL-np", "MA-np")
# CHANGE THE WRITE PATH TO A FIXED FOLDER
