import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import FuncFormatter


def read_excel_file(tabname):
    try:
        # Read the data from the Excel file
        df_booked_data = pd.read_excel(
            f'./ExcelData/Howden_CompanyXYZ_2021_Data.xlsx', sheet_name=tabname, skiprows=4, usecols="A:S", nrows=12)
        df_booked_data = df_booked_data.drop(
            df_booked_data.columns[2:15], axis=1)
    except Exception as e:
        return 500, f"Error reading Excel file: {e}"
    return 200, df_booked_data


def process_data(df_booked_data):
    try:
        df_booked_data['UltimateLossRatio'] = df_booked_data['Paid losses'] + \
            df_booked_data['Case reserves'] + df_booked_data['IBNR']

        df_booked_data = df_booked_data.rename(columns={
            'Earned premium': 'EarnedPremium',
            'Paid losses': 'PaidLosses',
            'Case reserves': 'CaseReserves',
            'U/W year': 'Year',
            'Gross written premium': 'GrossWrittenPremium',
        })
    except Exception as e:
        return 500, f"Error processing data: {e}"
    return 200, df_booked_data


def plot_data(df_booked_data):
    try:
        fig, ax1 = plt.subplots(figsize=(15, 8))

        # First y-axis for Ultimate Loss Ratio
        # Plot Ultimate Loss Ratio on the left y-axis
        df_booked_data.plot(x='Year', y=['PaidLosses', 'CaseReserves', 'IBNR'],
                            kind='bar', stacked=True, ax=ax1, width=0.8, position=1)
        ax1.set_ylabel('Ultimate loss ratio split')
        # Format the right y-axis as percentages
        formatter = FuncFormatter(lambda y, pos: f'{y:.0%}')
        ax1.yaxis.set_major_formatter(formatter)

        # Second y-axis for Earned Premium
        ax2 = ax1.twinx()
        df_booked_data['EarnedPremium'].plot(
            ax=ax2, marker='o', linestyle='-', color='b', label='Earned Premium')
        ax2.set_ylabel('Earned Premium')

        # x-axis label and title
        ax1.set_xlabel('U/W year')
        plt.title('Ultimate Loss Ratio')

        # Set legend
        ax1.legend(loc="upper center", bbox_to_anchor=(0.4, -0.15), ncol=4)
        ax2.legend(loc="upper center", bbox_to_anchor=(0.8, -0.15), ncol=1)

        plt.show()

    except Exception as e:
        return 500, f"Error creating plot: {e}"


def main(tabname):
    status, df_booked_data = read_excel_file(tabname)
    if status != 200:
        return f"Error reading Excel file: {df_booked_data}"

    status, df_booked_data = process_data(df_booked_data)
    if status != 200:
        return f"Error processing data: {df_booked_data}"

    status = plot_data(df_booked_data)
    if status != 200:
        return "Error creating plot"


if __name__ == "__main__":
    main('GL-np')
