import pandas as pd
from datetime import datetime


def booked_data_insert_script(spreadsheet, *tabnames):
    try:
        if len(tabnames) < 2:
            print("Please provide at least two sheet names.")
            return 400  # Bad request status code

        dfs = []
        for tabname in tabnames:
            df = pd.read_excel(
                f'./ExcelData/{spreadsheet}', sheet_name=tabname, skiprows=4, usecols="A:S", nrows=12)
            df['tabname'] = tabname
            df['CompanyName'] = 'XYZ'
            df = df.drop(df.columns[2:15], axis=1)
            dfs.append(df)

        df_combined_booked_data = pd.concat(dfs, ignore_index=True)

        df_combined_booked_data = df_combined_booked_data.dropna(
            axis=0, how='all')

        df_combined_booked_data['LineOfBusiness'] = df_combined_booked_data.apply(
            lambda row: 'General Liability' if 'GL' in row['tabname'] else 'Motor/Accident', axis=1)
        df_combined_booked_data['Currency'] = 'EUR'
        df_combined_booked_data['DWCreatedDate'] = datetime.today().strftime(
            '%Y-%m-%d')
        df_combined_booked_data['DWCreatedBy'] = 'Christian'

        df_combined_booked_data['UltimateLossRatio'] = df_combined_booked_data['Paid losses'] + \
            df_combined_booked_data['Case reserves'] + \
            df_combined_booked_data['IBNR']

        df_combined_booked_data = df_combined_booked_data.rename(columns={
            'Earned premium': 'EarnedPremium',
            'Paid losses': 'PaidLosses',
            'Case reserves': 'CaseReserves',
            'U/W year': 'Year',
            'Gross written premium': 'GrossWrittenPremium',
        })

        insert_statements = []

        for index, row in df_combined_booked_data.iterrows():
            values = f"{row['LineOfBusiness']}, {row['Year']}, {row['GrossWrittenPremium']}, {row['EarnedPremium']}, {row['PaidLosses']}, {row['CaseReserves']}, {row['IBNR']}, {row['UltimateLossRatio']}, '{row['DWCreatedDate']}', '{row['DWCreatedBy']}'"
            insert_statement = f"INSERT INTO [ClaimDevelopment].[FactData] ([LineOfBusiness], [Year], [GrossWrittenPremium], [EarnedPremium], [PaidLosses], [CaseReserves], [IBNR], [UltimateLossRatio], [DWCreatedDate], [DWCreatedBy]) VALUES ({values});"
            insert_statements.append(insert_statement)

        with open("insert_script_booked_data.sql", "w") as sql_file:
            for statement in insert_statements:
                sql_file.write(statement + "\n")

        print('Script created successfully')
        return 200  # Success status code
    except Exception as e:
        print(f"An error occurred: {e}")
        return 500  # Error status code


booked_data_insert_script('Howden_CompanyXYZ_2021_Data.xlsx', 'GL-np', 'MA-np')

# CHANGE THE WRITE PATH TO A FIXED FOLDER
