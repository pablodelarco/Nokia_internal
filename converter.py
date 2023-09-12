import pandas as pd
import xlsxwriter
import datetime
from datetime import time
import os

OUTPUT_FOLDER = "output"

kpi_number_list = ["NR_578C", "NR_579C", "NR_5258A", "NR_5259A", "NR_5739A", "NR_5740A"]
kpi_number_reference = ["NR_5192D", "NR_5193B", "NR_716A", "NR_718A"]


def process_worksheet(df, kpi_name):
    kpi_df = df[["Period start time", "NRCEL name", kpi_name]].sort_values(
        by="Period start time"
    )
    kpi_df = kpi_df.pivot(
        index="Period start time", columns="NRCEL name", values=kpi_name
    )

    kpi_df["Time"] = pd.to_datetime(kpi_df.index).time
    kpi_df["Date"] = pd.to_datetime(kpi_df.index).date

    return kpi_df


def generate_output_filename(date_of_output):
    date_timestamp = pd.Timestamp(date_of_output)  # Replace with your actual Timestamp
    date_str = date_timestamp.strftime("%Y-%m-%d %H:%M:%S")

    # Parse the input date string into a datetime object
    date_of_output = datetime.datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S")

    # Extract month and day from the input date
    month_day = date_of_output.strftime("%B_%d")

    # Create the output file name
    output_file = f"python_excel_{month_day}.xlsx"

    return output_file


def main(excel_file, excel_file_reference, output_folder, time_ranges_input):
    # time_ranges_input = input("Enter time ranges (e.g., '19:35, 22:25, 23:00, 6:55'): ")

    time_ranges = [
        time(int(t.split(":")[0]), int(t.split(":")[1]))
        for t in time_ranges_input.split(", ")
    ]

    excel_data = pd.ExcelFile(excel_file)
    worksheet_names = excel_data.sheet_names

    excel_data_2 = pd.ExcelFile(excel_file_reference)
    worksheet_names_2 = excel_data_2.sheet_names

    kpi_data = {}
    kpi_numbers = {}
    kpi_data_2 = {}
    kpi_numbers_2 = {}

    for worksheet_name in worksheet_names_2:
        if "Data for" in worksheet_name:
            df = pd.read_excel(excel_file_reference, sheet_name=worksheet_name)
            kpi_names_2 = df.columns[4:]
            kpi_numbers_2[worksheet_name] = df.iloc[0, 4:].tolist()

            for kpi_name, kpi_number in zip(kpi_names_2, kpi_numbers_2[worksheet_name]):
                if kpi_number in kpi_number_reference:
                    new_df = process_worksheet(df, kpi_name)
                    new_df.drop(columns=["Time"], inplace=True)
                    new_df.drop(columns=["Date"], inplace=True)
                    kpi_data_2[kpi_name] = (new_df, kpi_number)

    for worksheet_name in worksheet_names:
        if "Data for" in worksheet_name:
            df = pd.read_excel(excel_file, sheet_name=worksheet_name)
            kpi_names = df.columns[4:]
            kpi_numbers[worksheet_name] = df.iloc[0, 4:].tolist()
            date_of_output = df.iloc[4]["Period start time"]

            for kpi_name, kpi_number in zip(kpi_names, kpi_numbers[worksheet_name]):
                if kpi_number in kpi_number_list:
                    kpi_df = process_worksheet(df, kpi_name)
                    kpi_data[kpi_name] = (kpi_df, kpi_number)

    output_file = generate_output_filename(date_of_output)

    try:
        output_file_path = os.path.join(output_folder, output_file)

        with pd.ExcelWriter(output_file_path, engine="xlsxwriter") as writer:
            workbook = writer.book

            for kpi_name, (kpi_df, kpi_number) in kpi_data.items():
                worksheet = writer.book.add_worksheet(kpi_number)
                merge_format = workbook.add_format(
                    {
                        "bold": 1,
                        "border": 1,
                        "align": "center",
                        "valign": "vcenter",
                    }
                )

                worksheet.merge_range(0, 2, 0, 4, kpi_name, merge_format)
                worksheet.merge_range(1, 2, 1, 4, kpi_number, merge_format)

                color_index = 0
                for row_num, (index, row) in enumerate(kpi_df.iterrows(), start=3):
                    time_value = row["Time"]
                    date_value = row["Date"]

                    if pd.notna(time_value):
                        for i in range(0, len(time_ranges) - 1, 2):
                            start_time = time_ranges[i]
                            end_time = time_ranges[i + 1]

                            start_date = kpi_df.loc[
                                kpi_df["Time"] == start_time, "Date"
                            ].item()
                            end_date = kpi_df.loc[
                                kpi_df["Time"] == end_time, "Date"
                            ].item()

                            red_background_format = workbook.add_format(
                                {"bg_color": "#FFC7CE"}
                            )
                            green_background_format = workbook.add_format(
                                {"bg_color": "#C6EFCE"}
                            )

                            if (
                                (start_time <= time_value <= end_time)
                                or not ((start_time > time_value > end_time))
                                and start_date != end_date
                            ):
                                if color_index == 0:
                                    worksheet.conditional_format(
                                        row_num,
                                        2,
                                        row_num,
                                        4,
                                        {
                                            "type": "cell",
                                            "criteria": ">=",
                                            "value": 0,
                                            "format": red_background_format,
                                        },
                                    )
                                    if time_value == end_time:
                                        color_index = 1
                                else:
                                    worksheet.conditional_format(
                                        row_num,
                                        2,
                                        row_num,
                                        4,
                                        {
                                            "type": "cell",
                                            "criteria": ">=",
                                            "value": 0,
                                            "format": green_background_format,
                                        },
                                    )
                                    if time_value == end_time:
                                        color_index = 0
                                break

                kpi_df.drop(columns=["Time"], inplace=True)
                kpi_df.drop(columns=["Date"], inplace=True)

                kpi_df.to_excel(
                    writer,
                    sheet_name=kpi_number,
                    startrow=2,
                    index=True,
                    freeze_panes=(1, 0),
                    na_rep="",
                )
                worksheet.set_column(0, 0, 20)
                worksheet.set_column(2, len(kpi_df.columns) + 1, 15)
                worksheet.write_string(3, 1, "")

                idx = 2
                for kpi_name_2, (kpi_df_2, kpi_number_2) in kpi_data_2.items():
                    start_col = len(kpi_df.columns) + idx
                    kpi_df_2.to_excel(
                        writer,
                        sheet_name=kpi_number,
                        startcol=start_col,
                        startrow=2,
                        index=False,
                    )
                    worksheet.write_string(3, start_col, "")
                    worksheet.merge_range(
                        0,
                        start_col + 1,
                        0,
                        start_col + len(kpi_df.columns) - 1,
                        kpi_name_2,
                        merge_format,
                    )
                    worksheet.merge_range(
                        1,
                        start_col + 1,
                        1,
                        start_col + len(kpi_df.columns) - 1,
                        kpi_number_2,
                        merge_format,
                    )
                    worksheet.set_column(start_col + 3, len(kpi_df.columns) + 1, 15)

                    idx = idx + len(kpi_df_2.columns)

        print(f"Combined tables saved to '{output_file}'")
    except Exception as e:
        print(f"Error creating the file: {str(e)}")

    return output_file_path


if __name__ == "__main__":
    main()
