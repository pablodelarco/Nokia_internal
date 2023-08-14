import pandas as pd
import xlsxwriter

def process_worksheet(df, kpi_name):
    # Sort the DataFrame by "Period start time" and pivot it for better structure
    kpi_df = df[["Period start time", "NRCEL name", kpi_name]].sort_values(by="Period start time")
    kpi_df = kpi_df.pivot(index="Period start time", columns="NRCEL name", values=kpi_name)
    return kpi_df

def main():
    excel_file = "testdata.xlsx"
    worksheet_names = ["Data for Active cell throughput", "Data for Maximum MAC SDU cell t"]

    kpi_data = {}  # Dictionary to hold processed data
    kpi_numbers = {}  # Dictionary to hold KPI numbers

    # Read each worksheet and process the data
    for worksheet_name in worksheet_names:
        df = pd.read_excel(excel_file, sheet_name=worksheet_name)
        kpi_names = df.columns[4:]
        kpi_numbers[worksheet_name] = df.iloc[0, 4:].tolist()

        # Process each KPI and store data along with its number
        for kpi_name, kpi_number in zip(kpi_names, kpi_numbers[worksheet_name]):
            kpi_df = process_worksheet(df, kpi_name)
            kpi_data[kpi_name] = (kpi_df, kpi_number)

    output_file = "combined_tables.xlsx"
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        workbook = writer.book

        for kpi_name, (kpi_df, kpi_number) in kpi_data.items():
            new_kpi_name = kpi_number  # Format KPI number for worksheet name

            worksheet = writer.book.add_worksheet(new_kpi_name)

            kpi_df.to_excel(writer, sheet_name=new_kpi_name, index=True)

            # Set column widths for specific columns
            worksheet.set_column(0, 0, 20)  # Period start time column
            worksheet.set_column(2, len(kpi_df.columns) + 1, 15)  # KPI value columns

            worksheet.write_string(1, 1, '')  # Empty cell
            merge_format = workbook.add_format({
                "bold": 1,
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            })
            worksheet.merge_range("C2:E2", kpi_name, merge_format)

    print(f"Combined tables saved to '{output_file}'")

if __name__ == "__main__":
    main()
