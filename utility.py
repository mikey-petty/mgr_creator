import pandas as pd
from constants import *


# ----------------------------- MISCELLANEOUS FUNCTIONS ------------------------------
def num_row_to_letter(num_col: int) -> str:
    start_index = 1  #  it can start either at 0 or at 1
    letter = ""
    while num_col > 25 + start_index:
        letter += chr(65 + int((num_col - start_index) / 26) - 1)
        num_col = num_col - (int((num_col - start_index) / 26)) * 26
    letter += chr(65 - start_index + (int(num_col)))
    return letter


# ----------------------------- FORMULA WRITING FUNCTIONS ------------------------------


def create_summary_sheet(sheets_list) -> pd.DataFrame:
    """Write a summary sheet for mgr using excel formulas.

    Args:
        writer (pd.ExcelWriter): writer for writing excel formulas
    """
    summary_rows = []
    for sheet in sheets_list:
        if sheet[1] == "Grade":
            summary_rows.append(int(sheet[0]))

    summary_rows.append("Totals")

    summary_data = {
        key: [None for x in range(0, len(summary_rows))] for key in SUMMARY_COLS
    }
    summary_df = pd.DataFrame(summary_data)

    summary_df["Grade"] = summary_rows

    (max_row, max_col) = summary_df.shape

    summary_df.index = pd.RangeIndex(start=0, stop=max_row)

    summary_df = summary_df.reset_index()

    # write number of students failing 0 to file
    for grade in summary_rows[:-1]:
        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 0"
        ] = "=COUNTIFS('Full School MGR'!D2:D1000, \"=0\", 'Full School MGR'!A2:A1000, \"={}\")".format(
            grade
        )
        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 1"
        ] = "=COUNTIFS('Full School MGR'!D2:D1000, \"=1\", 'Full School MGR'!A2:A1000, \"={}\")".format(
            grade
        )
        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 2+"
        ] = "=COUNTIFS('Full School MGR'!D2:D1000, \">=2\", 'Full School MGR'!A2:A1000, \"={}\")".format(
            grade
        )

        # write total grade levels enrolled
        summary_df.loc[
            summary_df["Grade"] == grade, "Total Enrolled"
        ] = "=COUNTIF('Full School MGR'!A2:A1000, \"={}\")".format(grade)

        # write percentage failing 0
        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 0 - Pct"
        ] = "=B{}/I{}".format(
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
        )

        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 1 - Pct"
        ] = "=D{}/I{}".format(
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
        )

        summary_df.loc[
            summary_df["Grade"] == grade, "Failing 2+ - Pct"
        ] = "=F{}/I{}".format(
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
        )

        # write pass percentage
        summary_df.loc[
            summary_df["Grade"] == grade, "Pass Rate"
        ] = "=(B{}+D{})/I{}".format(
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
            summary_df.loc[summary_df["Grade"] == grade].index[0] + 2,
        )

    # write totals to file
    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 0"] = (
        "=SUM(B2:B" + str(max_row) + ")"
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 0 - Pct"] = (
        "=B" + str(max_row + 1) + "/I" + str(max_row + 1)
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 1"] = (
        "=SUM(D2:D" + str(max_row) + ")"
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 1 - Pct"] = (
        "=D" + str(max_row + 1) + "/I" + str(max_row + 1)
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 2+"] = (
        "=SUM(F2:F" + str(max_row) + ")"
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Failing 2+ - Pct"] = (
        "=F" + str(max_row + 1) + "/I" + str(max_row + 1)
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Total Enrolled"] = (
        "=SUM(I2:I" + str(max_row) + ")"
    )

    summary_df.loc[summary_df["Grade"] == "Totals", "Pass Rate"] = (
        "=(B" + str(max_row + 1) + "+D" + str(max_row + 1) + ")/I" + str(max_row + 1)
    )

    summary_df = summary_df.drop(columns="index")

    return summary_df


def write_to_file(df: pd.DataFrame, sheet_name: str, writer: pd.ExcelWriter) -> bool:
    """Writes each grades gradebook to an individual sheet

    Args:
        df (pd.DataFrame): dataframe containing data to write
        sheet_name (str): sheet name to
        writer (pd.ExcelWriter): writer for excel

    Returns:
        bool: returns whether or not it was successful
    """
    # Get the xlsxwriter workbook and worksheet objects.
    df.columns = df.columns.str.replace("_", " ")
    df.to_excel(writer, sheet_name=sheet_name, index=False)


def create_honor_roll_table(honor_roll_col: int) -> pd.DataFrame:
    col_letter = num_row_to_letter(honor_roll_col + 1)
    honor_roll_df = pd.DataFrame(
        data={
            "name": ["High Honor Roll", "Honor Roll", "Total"],
            "total": [None, None, None],
        }
    )

    honor_roll_df.loc[
        honor_roll_df["name"] == "High Honor Roll", "total"
    ] = "=COUNTIF('Full School MGR'!{}2:{}1000, \"=High Honor Roll\")".format(
        col_letter, col_letter
    )
    honor_roll_df.loc[
        honor_roll_df["name"] == "Honor Roll", "total"
    ] = "=COUNTIF('Full School MGR'!{}2:{}1000, \"=Honor Roll\")".format(
        col_letter, col_letter
    )
    honor_roll_df.loc[honor_roll_df["name"] == "Total", "total"] = "=L1+L2"

    return honor_roll_df


def write_formulas_grade_sheet(df: pd.DataFrame) -> pd.DataFrame:
    df.index = pd.RangeIndex(len(df.index))
    (num_row, num_col) = df.shape

    average_formula_list = []
    failing_formula_list = []
    honor_roll_formula_list = []
    for i in range(0, num_row):
        average_formula_list.append(
            "=ROUND(AVERAGE({}{}:{}{})/100,3)".format(
                num_row_to_letter(
                    df.columns.get_loc("# Failing") + 2
                ),  # first course col number
                i + 2,  # student row number
                num_row_to_letter(
                    df.columns.get_loc("Average")
                ),  # last course col number
                i + 2,  # student column number
            )
        )

        honor_roll_formula_list.append(
            '=IF({}{} >= 0.93, "High Honor Roll", IF({}{} >= 0.85, "Honor Roll",""))'.format(
                num_row_to_letter(
                    df.columns.get_loc("Average") + 1
                ),  # first course col number
                i + 2,  # student row number
                num_row_to_letter(
                    df.columns.get_loc("Average") + 1
                ),  # first course col number
                i + 2,  # student row number
            )
        )

        failing_formula_list.append(
            '=COUNTIF({}{}:{}{}, "<70")'.format(
                num_row_to_letter(
                    df.columns.get_loc("# Failing") + 2
                ),  # first course col number
                i + 2,  # student row number
                num_row_to_letter(
                    df.columns.get_loc("Average")
                ),  # last course col number
                i + 2,  # student column number
            )
        )

    # write formulas directly to df
    df.loc[:, "Average"] = average_formula_list
    df.loc[:, "# Failing"] = failing_formula_list
    df.loc[:, "Honor Roll Status"] = honor_roll_formula_list

    return df


# ----------------------------- FORMATTING FUNCTIONS ------------------------------


def format_grades(df: pd.DataFrame, sheet_name: str, writer: pd.ExcelWriter):
    """conditionally formats grades sheets

    Args:
        df (pd.DataFrame): which sheet to format from
        sheet_name (str): sheet name to write to
        writer (pd.ExcelWriter): xlsxwriter object to allow for conditional formatting in excel
    """
    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Get the dimensions of the dataframe.
    (max_row, max_col) = df.shape

    grades_range = (
        num_row_to_letter(df.columns.get_loc("# Failing") + 2)
        + "2:"
        + num_row_to_letter(df.columns.get_loc("Average"))
        + str(max_row + 1)
    )
    average_range = (
        num_row_to_letter(df.columns.get_loc("Average") + 1)
        + "2:"
        + num_row_to_letter(df.columns.get_loc("Average") + 1)
        + str(max_row + 1)
    )
    failing_range = (
        num_row_to_letter(df.columns.get_loc("# Failing") + 1)
        + "2:"
        + num_row_to_letter(df.columns.get_loc("# Failing") + 1)
        + str(max_row + 1)
    )
    honor_roll_range = (
        num_row_to_letter(df.columns.get_loc("Honor Roll Status") + 1)
        + "2:"
        + num_row_to_letter(df.columns.get_loc("Honor Roll Status") + 1)
        + str(max_row + 1)
    )
    header_range = (
        "A1:" + num_row_to_letter(df.columns.get_loc("Honor Roll Status") + 1) + "1"
    )

    header_format = workbook.add_format(
        {
            "bg_color": "#d9e1f2",
            "font_color": "#000000",
            "font_size": 12,
            "bold": 1,
            "border": 0,
        }
    )

    average_format = workbook.add_format(
        {
            "num_format": 10,
            "font_size": 12,
        }
    )

    # give high honor roll students a dark orange fill
    high_honor_roll_format = workbook.add_format(
        {
            "bg_color": "#ed7d31",
            "font_color": "#000000",
            "font_size": 12,
        }
    )

    # give honor roll students a light orange fill
    honor_roll_format = workbook.add_format(
        {
            "bg_color": "#FFC000",
            "font_color": "#000000",
            "font_size": 12,
        }
    )

    # give failing grades a light red fill
    failing_format = workbook.add_format(
        {
            "bg_color": "#ffc7ce",
            "font_color": "#000000",
            "font_size": 12,
        }
    )

    missing_format = workbook.add_format(
        {
            "bg_color": "#ffff66",
            "font_color": "#000000",
            "font_size": 12,
        }
    )

    worksheet.conditional_format(
        header_range, {"type": "no_errors", "format": header_format}
    )

    worksheet.conditional_format(
        grades_range, {"type": "blanks", "format": missing_format}
    )

    worksheet.conditional_format(
        failing_range,
        {"type": "cell", "criteria": ">", "value": 1, "format": failing_format},
    )

    worksheet.conditional_format(
        grades_range,
        {"type": "cell", "criteria": "<", "value": 70, "format": failing_format},
    )

    worksheet.conditional_format(
        average_range,
        {"type": "cell", "criteria": "<", "value": 0.70, "format": failing_format},
    )

    worksheet.conditional_format(
        honor_roll_range,
        {
            "type": "cell",
            "criteria": "==",
            "value": '"High Honor Roll"',
            "format": high_honor_roll_format,
        },
    )

    worksheet.conditional_format(
        honor_roll_range,
        {
            "type": "cell",
            "criteria": "==",
            "value": '"Honor Roll"',
            "format": honor_roll_format,
        },
    )

    worksheet.conditional_format(
        average_range, {"type": "no_errors", "format": average_format}
    )

    worksheet.set_column(0, df.columns.get_loc("Student"), 13.33)
    worksheet.set_column(
        df.columns.get_loc("Student"), df.columns.get_loc("Student"), 29.6
    )
    worksheet.set_column(
        df.columns.get_loc("# Failing") + 2, df.columns.get_loc("Average"), 13.33
    )
    worksheet.set_column(
        df.columns.get_loc("Honor Roll Status"),
        df.columns.get_loc("Honor Roll Status"),
        18,
    )


def format_summary(df: pd.DataFrame, sheet_name: str, writer: pd.ExcelWriter):
    """conditionally formats grades sheets

    Args:
        df (pd.DataFrame): which sheet to format from
        sheet_name (str): sheet name to write to
        writer (pd.ExcelWriter): xlsxwriter object to allow for conditional formatting in excel
    """

    # Get the xlsxwriter workbook and worksheet objects.
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Get the dimensions of the dataframe.
    (max_row, max_col) = df.shape

    # header format
    header_format = workbook.add_format(
        {
            "bg_color": "#bdd6ee",
            "font_color": "#000000",
            "font_name": "Calibri",
            "font_size": 12,
            "bold": 1,
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )

    body_num_format = workbook.add_format(
        {
            "bold": 1,
            "font_size": 12,
            "align": "center",
            "valign": "vcenter",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )

    body_percent_format = workbook.add_format(
        {
            "bold": 1,
            "font_name": "Calibri",
            "font_size": 12,
            "align": "center",
            "valign": "vcenter",
            "align": "center",
            "valign": "vcenter",
            "num_format": 10,
            "border": 1,
        }
    )

    # give high honor roll students a dark orange fill
    high_honor_roll_format = workbook.add_format(
        {
            "bold": 1,
            "bg_color": "#ed7d31",
            "font_color": "#000000",
            "font_name": "Calibri",
            "font_size": 12,
            "border": 1,
        }
    )

    # give honor roll students a light orange fill
    honor_roll_format = workbook.add_format(
        {
            "bold": 1,
            "bg_color": "#FFC000",
            "font_color": "#000000",
            "font_name": "Calibri",
            "font_size": 12,
            "border": 1,
        }
    )

    # give failing grades a light red fill
    total_honor_roll_format = workbook.add_format(
        {
            "bold": 1,
            "bg_color": "#ffff00",
            "font_color": "#000000",
            "font_name": "Calibri",
            "font_size": 12,
            "border": 1,
        }
    )

    worksheet.conditional_format(
        "B2:B" + str(max_row + 1), {"type": "no_errors", "format": body_num_format}
    )
    worksheet.conditional_format(
        "D2:D" + str(max_row + 1), {"type": "no_errors", "format": body_num_format}
    )
    worksheet.conditional_format(
        "F2:F" + str(max_row + 1), {"type": "no_errors", "format": body_num_format}
    )
    worksheet.conditional_format(
        "I2:I" + str(max_row + 1), {"type": "no_errors", "format": body_num_format}
    )

    worksheet.conditional_format(
        "C2:C" + str(max_row + 1), {"type": "no_errors", "format": body_percent_format}
    )
    worksheet.conditional_format(
        "E2:E" + str(max_row + 1), {"type": "no_errors", "format": body_percent_format}
    )
    worksheet.conditional_format(
        "G2:G" + str(max_row + 1), {"type": "no_errors", "format": body_percent_format}
    )
    worksheet.conditional_format(
        "H2:H" + str(max_row + 1), {"type": "no_errors", "format": body_percent_format}
    )

    worksheet.conditional_format(
        "K1:L1", {"type": "no_errors", "format": high_honor_roll_format}
    )

    worksheet.conditional_format(
        "K2:L2", {"type": "no_errors", "format": honor_roll_format}
    )

    worksheet.conditional_format(
        "K2:L3", {"type": "no_errors", "format": total_honor_roll_format}
    )

    worksheet.conditional_format(
        "A1:I1", {"type": "no_errors", "format": header_format}
    )

    worksheet.conditional_format(
        "A1:A" + str(max_row + 1), {"type": "no_errors", "format": header_format}
    )

    worksheet.set_column(0, 25, 13.33)
