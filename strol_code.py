import pandas as pd
import docx
import re
from docx import Document
import docx.table
import docx.text.paragraph
from docx.shared import Inches, Pt, RGBColor
from typing import Dict, Union
from docxtpl import DocxTemplate


from constants import *


def validate_excel(file_path=INPUT_PATH):

    try:
        df = excel_to_dataframe(file_path)
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return False

    actual_columns = list(df.columns)

    column_match = set(actual_columns) == set(COLUMNS)

    empty_cells = df.isnull().values.any()

    if not column_match:
        print(f"❌ Column names do not match expected names.\nExpected: {COLUMNS}\nActual: {actual_columns}")

    return column_match and not empty_cells


def excel_to_dataframe(file_path=INPUT_PATH):
    return pd.read_excel(file_path)


def calculations_on_seperated_data(df: pd.DataFrame, commander):
    placeholder_to_value = {}

    df_filtered = df[df[COMMANDER_COLUMN] == commander]

    for option in OPTIONS:
        if option is not NONE_OF_THE_ABOVE_OPTION:
            count = count_occurrences(df_filtered, option)
            placeholder_to_value[OPTIONS_TO_PLACEHOLDERS[option]] = count / len(df_filtered)

    # I handle it differently as it appears in all of the questions
    handle_none_of_the_above(df_filtered, placeholder_to_value)

    for column in OPEN_QUESTIONS_COLUMNS:
        placeholder_to_value[column] = [str(item) for item in df_filtered[column].dropna().tolist()]

    return placeholder_to_value


def handle_none_of_the_above(df_filtered: pd.DataFrame, placeholder_to_value: Dict):
    for index, col in enumerate(MULTIPLE_CHOICE_COLUMNS):
        count = count_occurrences(df_filtered[col], NONE_OF_THE_ABOVE_OPTION)
        placeholder_to_value[OPTIONS_TO_PLACEHOLDERS[NONE_OF_THE_ABOVE_OPTION + f"_{index}"]] = count / len(df_filtered)


def calculate_total_percentage(df: pd.DataFrame):
    mahzor_averages = {}

    for option in OPTIONS:
        if option is not NONE_OF_THE_ABOVE_OPTION:
            mahzor_averages[OPTIONS_TO_PLACEHOLDERS[option]] = count_occurrences(df, option) / len(df) * 100

    # I handle it differently as it appears in all of the questions
    for index, col in enumerate(MULTIPLE_CHOICE_COLUMNS):
        count = count_occurrences(df[col], NONE_OF_THE_ABOVE_OPTION)
        mahzor_averages[OPTIONS_TO_PLACEHOLDERS[NONE_OF_THE_ABOVE_OPTION + f"_{index}"]] = count / len(df) * 100

    return mahzor_averages


def count_occurrences(data: Union[pd.DataFrame, pd.Series], target: str, separator: str = ",") -> int:
    """
    Count how many times a given string appears anywhere in a DataFrame.

    Args:
        df (pd.DataFrame): The DataFrame to search.
        target (str): The string to count.
        separator (str): The separator used for multiple strings in a cell (default: ',').

    Returns:
        int: Total number of times the string appears across the entire DataFrame.
    """
    count = 0
    if isinstance(data, pd.DataFrame):
        iterator = (cell for col in data.columns for cell in data[col])
    elif isinstance(data, pd.Series):
        iterator = iter(data)
    else:
        raise TypeError("data must be a pandas DataFrame or Series")

    for cell in iterator:
        if isinstance(cell, str):
            values = [v.strip() for v in cell.split(separator) if v.strip() != ""]
            # count every match (so duplicates in same cell count multiple times)
            count += sum(1 for v in values if v == target)

    return count


def validate_calculations(placeholder_to_value: Dict):
    # check if there is anything left empty
    if None in placeholder_to_value.values() or "" in placeholder_to_value.values():
        print("❌ Some placeholders have empty values.")
        return False
    return True


def generate_and_fill_commander_docx(df, placeholder_to_value, commander, template_path=TEMPLATE_PATH, output_path=OUTPUT_PATH):
    doc = Document(template_path)

    replace_placeholders(doc, placeholder_to_value)

    add_bullet_lists(DocxTemplate(TEMPLATE_PATH), df[df[COMMANDER_COLUMN] == commander], commander)

    doc.save(output_path + commander + ".docx")


def replace_placeholders_in_paragraph(paragraph: docx.text.paragraph.Paragraph, values: Dict):
    for run in paragraph.runs:
        for placeholder, value in values.items():
            print(f"Replacing {placeholder} with {value} at text {run.text}")
            for p in placeholder:
                p = "{{" + p + "}}"

                if str(p) not in run.text:
                    continue
                if str(p) in run.text:
                    run.text = run.text.replace(p, str(value))
                    print("completed")


def replace_placeholders_in_table(table: docx.table.Table, values: Dict):
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                replace_placeholders_in_paragraph(paragraph, values)


def replace_placeholders_in_section(section, values: Dict):
    header = section.header
    footer = section.footer
    for part in [header, footer]:
        if part is None:
            return

        for paragraph in part.paragraphs:
            replace_placeholders_in_paragraph(paragraph, values)
        for table in part.tables:
            replace_placeholders_in_table(table, values)


# TODO: make it work with filling a bullet list
def replace_placeholders(doc: Document, values: Dict):
    # Body paragraphs
    for paragraph in doc.paragraphs:
        replace_placeholders_in_paragraph(paragraph, values)

    # Tables
    for table in doc.tables:
        replace_placeholders_in_table(table, values)

    # Headers and footers
    for section in doc.sections:
        replace_placeholders_in_section(section, values)


def add_bullet_lists(doc: DocxTemplate, df: pd.DataFrame, commander: str):
    context = {
        "improve_command": {"points": df["נקודות חיוביות פיקוד:"]},
        "conserve_command": {"points": df["נקודות שליליות פיקוד:"]},
    }

    doc.render(context)
    # doc.save(OUTPUT_PATH + commander + ".docx")


def main(file_path=INPUT_PATH):
    if not validate_excel():
        print("Excel file validation failed.")
        return

    df = excel_to_dataframe(file_path)

    mahzor_averages = calculate_total_percentage(df)

    for commander in df[COMMANDER_COLUMN].unique():

        placeholder_to_value = calculations_on_seperated_data(df, commander)
        # merge with mahzor averages
        placeholder_to_value.update(mahzor_averages)

        if not validate_calculations(placeholder_to_value):
            print("Calculations validation failed.")
            return

        # TODO: currently the code can only do one (the second one).

        # add_bullet_lists(DocxTemplate(TEMPLATE_PATH), df[df[COMMANDER_COLUMN] == commander], commander)
        generate_and_fill_commander_docx(df, placeholder_to_value, commander)



if __name__ == "__main__":
    main(INPUT_PATH)
