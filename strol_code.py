import pandas as pd
import docx
import re
from docx import Document
import docx.table
import docx.text.paragraph
from docx.shared import Inches, Pt, RGBColor
from typing import Dict, Union
from docxtpl import DocxTemplate
import os
from constants import *





def validate_excel(file_path=INPUT_PATH) -> bool:
    try:
        df = excel_to_dataframe(file_path)
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return False

    actual_columns = list(df.columns)
    column_match = set(actual_columns) == set(COLUMNS)

    if not column_match:
        print(f"❌ Column names do not match expected names.\nExpected: {COLUMNS}\nActual: {actual_columns}")
        return False

    if df.empty:
        print("❌ Excel file is empty.")
        return False

    # Empty cells are allowed
    return True



def excel_to_dataframe(file_path=INPUT_PATH):
    return pd.read_excel(file_path)

def format_number(x: float):
    """
    If x is an integer
    Otherwise return rounded float with decimals
    """
    if x is None:
        return DEFAULT_ZERO_VALUE

    if float(x).is_integer():
        return int(x)
    rounded =  round(float(x), PERCENT_DECIMALS)
    return rounded
def compute_percent(count: int, total: int) -> float:
    """
    Returns count/total * 100, rounded to PERCENT_DECIMALS .
    """
    if total <= 0:
        return DEFAULT_ZERO_VALUE
    value = (count / total) * 100
    form_value = format_number(value)
    return f"{form_value}%"

def compute_mahzor_general_average(df: pd.DataFrame) -> float:
    """
    Computes the overall (mahzor) average of:
        'עד כמה הייתי רוצה להיות תחת פיקודו בעתיד?'
    """
    col = GENERAL_QUESTION_COLUMN

    raw = df[col]
    cleaned = raw.replace(r'^\s*$', pd.NA, regex=True)
    numeric = pd.to_numeric(cleaned, errors="coerce")
    series = numeric.dropna()

    if series.empty:
        return DEFAULT_ZERO_VALUE

    return round(float(series.mean()), 2)

def compute_commander_general_stats(df_filtered: pd.DataFrame) -> Dict:
    """
    Computes per-commander stats for:
        'עד כמה הייתי רוצה להיות תחת פיקודו בעתיד?'
    """
    col = "עד כמה הייתי רוצה להיות תחת פיקודו בעתיד?"

    raw = df_filtered[col]

    # Treat empty/whitespace as missing
    cleaned = raw.replace(r'^\s*$', pd.NA, regex=True)

    # Convert to numeric, invalid -> NaN, then drop NaN
    numeric = pd.to_numeric(cleaned, errors="coerce")
    series = numeric.dropna()

    n_valid = len(series)

    stats = {}
    if n_valid < MIN_GENERAL_ANSWERS:
        stats["average_general"] = TOO_FEW_ANSWERS_TEXT
        stats["std_general"] = TOO_FEW_ANSWERS_TEXT
        return stats

    mean_val = series.mean()
    std_val = series.std(ddof=1)

    if pd.isna(std_val):
        std_val = DEFAULT_ZERO_VALUE

    stats["average_general"] = round(float(mean_val), 2)
    stats["std_general"] = round(float(std_val), 2)

    return stats


def add_general_question_stats(df_all: pd.DataFrame,
                               df_filtered: pd.DataFrame,
                               placeholder_to_value: Dict):

    commander_stats = compute_commander_general_stats(df_filtered)
    mahzor_avg = compute_mahzor_general_average(df_all)
    placeholder_to_value.update(commander_stats)
    placeholder_to_value["total_general"] = mahzor_avg




def calculations_on_seperated_data(df: pd.DataFrame, commander):
    placeholder_to_value = {}

    df_filtered = df[df[COMMANDER_COLUMN] == commander]

    for option in OPTIONS:
        if option != NONE_OF_THE_ABOVE_OPTION:
            count = count_occurrences(df_filtered, option)
            percent_ph, _ = OPTIONS_TO_PLACEHOLDERS[option]  # split the tuple
            placeholder_to_value[percent_ph] = compute_percent(count,len(df_filtered))
    # I handle it differently as it appears in all of the questions
    handle_none_of_the_above(df_filtered, placeholder_to_value)

    for column in OPEN_QUESTIONS_COLUMNS:
        placeholder_to_value[column] = [str(item) for item in df_filtered[column].dropna().tolist()]

    add_general_question_stats(df, df_filtered, placeholder_to_value)

    return placeholder_to_value


def handle_none_of_the_above(df_filtered: pd.DataFrame, placeholder_to_value: Dict):
    for index, col in enumerate(MULTIPLE_CHOICE_COLUMNS):
        count = count_occurrences(df_filtered[col], NONE_OF_THE_ABOVE_OPTION)
        percent_ph, _ = OPTIONS_TO_PLACEHOLDERS[NONE_OF_THE_ABOVE_OPTION + f"_{index}"]
        placeholder_to_value[percent_ph] = compute_percent(count,len(df_filtered))



def calculate_total_percentage(df: pd.DataFrame):
    mahzor_averages = {}

    for option in OPTIONS:
        if option != NONE_OF_THE_ABOVE_OPTION:
            _, total_ph = OPTIONS_TO_PLACEHOLDERS[option]  # split the tuple
            count = count_occurrences(df, option)
            mahzor_averages[total_ph] = compute_percent(count,len(df))

    # "none of the above" per question
    for index, col in enumerate(MULTIPLE_CHOICE_COLUMNS):
        _, total_ph = OPTIONS_TO_PLACEHOLDERS[NONE_OF_THE_ABOVE_OPTION + f"_{index}"]
        count = count_occurrences(df[col], NONE_OF_THE_ABOVE_OPTION)
        mahzor_averages[total_ph] = compute_percent(count,len(df))

    return mahzor_averages



def count_occurrences(data: Union[pd.DataFrame, pd.Series], target: str) -> int:

    if isinstance(data, pd.DataFrame):
        iterator = (cell for col in data.columns for cell in data[col])
    elif isinstance(data, pd.Series):
        iterator = iter(data)
    else:
        raise TypeError("data must be a pandas DataFrame or Series")

    count = 0
    for cell in iterator:
        if not isinstance(cell, str):
            continue

        text = cell.strip()

        if target in text:
            count += 1

    return count


def validate_calculations(placeholder_to_value: Dict):
    # check if there is anything left empty
    if None in placeholder_to_value.values() or "" in placeholder_to_value.values():
        print("❌ Some placeholders have empty values.")
        return False
    return True


def generate_and_fill_commander_docx(df, placeholder_to_value, commander, template_path=TEMPLATE_PATH,
                                     output_path=OUTPUT_PATH):
    # First use python docx to replace placeholders and save
    doc = Document(template_path)

    replace_placeholders(doc, placeholder_to_value)

    doc.save(output_path + commander + ".docx")

    # : Use DocxTemplate (another library) to open the processed file and add bullet lists
    add_bullet_lists(DocxTemplate(output_path + commander + ".docx"), df[df[COMMANDER_COLUMN] == commander], commander)

def generate_and_fill_commander_docx(df: pd.DataFrame,
                                     placeholder_to_value: dict,
                                     commander: str,
                                     template_path: str = TEMPLATE_PATH,
                                     output_path: str = OUTPUT_PATH):
    # First use python docx to replace placeholders and save
    doc = Document(template_path)
    replace_placeholders(doc, placeholder_to_value)
    doc.save(output_path + commander + ".docx")

    df_commander = df[df[COMMANDER_COLUMN] == commander]
    tpl = DocxTemplate(output_path + commander + ".docx")
    add_bullet_lists(tpl, df_commander, commander)

    #cretes ecxel
    context = merge_bullet_lists(df_commander, commander)

    export_commander_excel(
        commander=commander,
        placeholder_to_value=placeholder_to_value,
        context=context,
    )


def replace_placeholders_in_paragraph(paragraph: docx.text.paragraph.Paragraph, values: Dict):
    text_runs = [run for run in paragraph.runs if run.text]

    if not text_runs:
        return

    original_texts = [run.text for run in text_runs]
    big_text = "".join(original_texts)

    if "{{" not in big_text:
        return

    for placeholder, value in values.items():
        token = "{{" + str(placeholder) + "}}"
        if token in big_text:
            big_text = big_text.replace(token, str(value))

    pos = 0
    for run, old_text in zip(text_runs, original_texts):
        length = len(old_text)
        run.text = big_text[pos:pos + length]
        pos += length




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


def build_basic_info_context(df_commander: pd.DataFrame, commander: str) -> Dict:
    answers_num = len(df_commander)
    return {
        'name': commander,
        'number_answers': answers_num,
    }


def rtl_embed(text: str) -> str:
    "make sure text is in RTL format"
    if not isinstance(text, str):
        return text

    text = text.strip()
    if not text:
        return text


    for ch in PUNCTUATION_CHARS:
        text = text.replace(ch, f"{RLM}{ch}{RLM}")

    return f"{RLE}{text}{PDF}"


def build_bullet_lists_context(df_commander: pd.DataFrame) -> Dict:
    context: Dict[str, Dict[str, list]] = {}

    for key, column_name in BULLET_LIST_CONTEXT.items():
        points: list[str] = []

        if column_name in df_commander.columns:
            series = df_commander[column_name].dropna().astype(str)

            cleaned_points: list[str] = []
            for raw in series:
                text = raw.strip()

                #
                if len(text) < MIN_COMMENT_LENGTH:
                    continue
                text = rtl_embed(text)

                cleaned_points.append(text)

            points = cleaned_points

        context[key] = {"points": points}

    return context


def merge_bullet_lists(df_commander: pd.DataFrame, commander: str):
    basic_info = build_basic_info_context(df_commander, commander)
    bullet_lists = build_bullet_lists_context(df_commander)
    context = {
        **basic_info,
        **bullet_lists
    }
    return context
def add_bullet_lists(doc: DocxTemplate, df_commander: pd.DataFrame, commander: str):
    context = merge_bullet_lists(df_commander,commander)
    doc.render(context)
    doc.save(OUTPUT_PATH + commander + ".docx")

def build_commander_summary_row(commander: str,
                                placeholder_to_value: dict,
                                context: dict) -> dict:
    row = {
        "name": context.get("name", commander),
        "number_answers": context.get("number_answers", None),
    }

    for key, value in placeholder_to_value.items():
        if isinstance(value, (list, dict)):
            continue
        row[key] = value

    return row

def build_commander_comments_rows_from_context(commander: str,
                                               context: dict) -> list[dict]:
    rows: list[dict] = []

    for key, column_name in BULLET_LIST_CONTEXT.items():
        entry = context.get(key, {})
        points = entry.get("points", []) if isinstance(entry, dict) else []

        for text in points:
            rows.append({
                "category_hebrew": column_name,
                "comment": text,
            })

    return rows

def build_summary_header_mapping() -> dict:
    header_map: dict[str, str] = {
        "name": "Commander Name",
        "number_answers": "Number of Respondents",
        "average_general": "General Rating – Commander Average",
        "std_general": "General Rating – Standard Deviation",
        "total_general": "General Rating – Cohort Average",
    }

    for opt_key, (percent_ph, total_ph) in OPTIONS_TO_PLACEHOLDERS.items():
        if opt_key.startswith(NONE_OF_THE_ABOVE_OPTION):
            idx = int(opt_key.split("_")[-1])
        else:
            idx = OPTIONS.index(opt_key)

        question = MULTIPLE_CHOICE_COLUMNS[idx // 5]

        if opt_key.startswith(NONE_OF_THE_ABOVE_OPTION):
            option_text = NONE_OF_THE_ABOVE_OPTION
        else:
            option_text = opt_key

        header_map[percent_ph] = f"{question} – {option_text} (% Commander)"
        header_map[total_ph] = f"{question} – {option_text} (% Cohort)"

    return header_map

def export_commander_excel(commander: str,
                           placeholder_to_value: dict,
                           context: dict,
                           base_path: str | None = None):
    if base_path is None:
        base_path = COMMANDER_EXCEL_OUTPUT_PATH

    os.makedirs(base_path, exist_ok=True)
    excel_path = os.path.join(base_path, f"{commander}.xlsx")

    summary_row = build_commander_summary_row(commander, placeholder_to_value, context)
    comments_rows = build_commander_comments_rows_from_context(commander, context)

    summary_df = pd.DataFrame([summary_row])
    comments_df = pd.DataFrame(comments_rows)

    header_map = build_summary_header_mapping()
    summary_df = summary_df.rename(columns=header_map)

    with pd.ExcelWriter(excel_path) as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        comments_df.to_excel(writer, sheet_name="Comments", index=False)

    print(f"Excel for commander {commander} written to: {excel_path}")



def main(file_path=INPUT_PATH):
    if not validate_excel():
        print("Excel file validation failed.")
        return

    df = excel_to_dataframe(file_path)
    # clen up empty rows / rows without commander name
    df = df.dropna(how="all")
    df = df.dropna(subset=[COMMANDER_COLUMN])

    mahzor_averages = calculate_total_percentage(df)

    for commander in df[COMMANDER_COLUMN].unique():

        placeholder_to_value = calculations_on_seperated_data(df, commander)
        # merge with mahzor averages
        placeholder_to_value.update(mahzor_averages)

        if not validate_calculations(placeholder_to_value):
            print("Calculations validation failed.")
            return




        generate_and_fill_commander_docx(df, placeholder_to_value, commander)


if __name__ == "__main__":
    main(INPUT_PATH)
