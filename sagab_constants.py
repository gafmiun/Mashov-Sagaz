from typing import Dict, List, Tuple

# =====================================
# File paths and basic configuration
# =====================================

INPUT_PATH: str = "answers_sagab.xlsx"
TEMPLATE_PATH: str = "mashov_sagab_template.docx"
OUTPUT_PATH: str = "output_sagab/"
COMMANDER_EXCEL_OUTPUT_PATH: str = "output_sagab/excel/"

MIN_GENERAL_ANSWERS: int = 4
TOO_FEW_ANSWERS_TEXT: str = "ענו פחות מ-4"

PERCENT_DECIMALS: int = 2
DEFAULT_ZERO_VALUE: float = 0.0

IGNORED_VALUE: float = 0.0
MIN_COMMENT_LENGTH: int = 2

# =====================================
# RTL helpers
# =====================================

RLE: str = "\u202B"  # Right-to-Left Embedding
PDF: str = "\u202C"  # Pop Directional Formatting
RLM: str = "\u200F"  # Right-to-Left Mark

PUNCTUATION_CHARS: List[str] = [
    ",", ".", "\"", "\\", "-", ":", ";", "(", ")", "!", "?", "+", "/"
]

# =====================================
# Core survey schema
# =====================================

COMMANDER_COLUMN: str = "ממלא/ת משוב על..."

GENERAL_QUESTION_COLUMN: str = (
    "עד כמה היית רוצה לניות תחת פיקודו/ה של המפקד/ת גם בעתיד? (0 משמעותו לא רלוונטי)"
)

# =====================================
# Numeric sections and questions
# =====================================

NUMERIC_SECTION_TO_QUESTIONS: Dict[str, List[str]] = {
    "חיבור מקצועי לתכני ההכשרה": [
        "באיזו מידה המפקד/ת מפגינ/ה שליטה בתכני ההכשרה וביעדים, ויודע/ת לענות בביטחון לשאלות? (0 משמעותו לא רלוונטי)",
        "עד כמה מביע/ה מפקד/ת המחלקה הזדהות וחיבור לתכני ההכשרה וערכיה? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מקשרת בין תכני ההכשרה, לבין יישומם בשטח? (0 משמעותו לא רלוונטי)",
    ],
    "נוכחות ומעורבות": [
        "באיזו מידה מורגשת נוכחותו/ה של המפקד/ת בהכשרה? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מודע/ת למצבך בתוכנית ועוקב/ת אחר התקדמותך? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת נגיש/ה וזמין/ה? (0 משמעותו לא רלוונטי)",
    ],
    "התנהלות בינאישית": [
        "באיזו מידה מגלה המפקד/ת אכפתיות כלפייך? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מתנהל/ת באופן מכבד? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מתנהל בפתיחות ובשקיפות? (0 משמעותו לא רלוונטי)",
    ],
    "פיקודיות": [
        "באיזו מידה מגלה המפקד/ת סמכותיות וביטחון בהתנהלותו/ה? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת רותמ/ת, מניע/ה ומעורר/ת בך מוטיבציה? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מהווה מודל לחיקוי עבורי? (0 משמעותו לא רלוונטי)",
    ],
    "פיתוח אישי ומקצועי": [
        "עד כמה שיחות אישיות עם המפקד/ת תורמות לך בעיבוד חוויות פיקודיות ומקצועיות בהכשרה? (0 משמעותו לא רלוונטי)",
        "עד כמה תורמ/ת המפקד/ת לקידום ההכשרה שלך? (0 משמעותו לא רלוונטי)",
        "עד כמה המפקד/ת מציב/ה סטנדרט גבוה, ודורש/ת חתירה למצוינות? (0 משמעותו לא רלוונטי)",
    ],
    "כללי": [GENERAL_QUESTION_COLUMN]
}

NUMERIC_SECTION_TO_PREFIX: Dict[str, str] = {
    "חיבור מקצועי לתכני ההכשרה": "professional",
    "נוכחות ומעורבות": "personal",
    "התנהלות בינאישית": "interpersonal",
    "פיקודיות": "command",
    "פיתוח אישי ומקצועי": "development",
    "כללי": "general"
}


# Short labels for X-axis, taken from the Word table "מדד"
SHORT_LABELS: Dict[str, str] = {
    # חיבור מקצועי לתכני ההכשרה
    "באיזו מידה המפקד/ת מפגינ/ה שליטה בתכני ההכשרה וביעדים, ויודע/ת לענות בביטחון לשאלות? (0 משמעותו לא רלוונטי)": "שליטה ובקיאות בתכני ההכשרה",
    "עד כמה מביע/ה מפקד/ת המחלקה הזדהות וחיבור לתכני ההכשרה וערכיה? (0 משמעותו לא רלוונטי)": "הזדהות וחיבור לתכנים ולערכים",
    "עד כמה המפקד/ת מקשרת בין תכני ההכשרה, לבין יישומם בשטח? (0 משמעותו לא רלוונטי)": "קישור בין תכני הכשרה ליישום בפועל",

    # נוכחות ומעורבות
    "באיזו מידה מורגשת נוכחותו/ה של המפקד/ת בהכשרה? (0 משמעותו לא רלוונטי)": "נוכחות מורגשת",
    "עד כמה המפקד/ת מודע/ת למצבך בתוכנית ועוקב/ת אחר התקדמותך? (0 משמעותו לא רלוונטי)": "מעקב אחר מצב הצוער",
    "עד כמה המפקד/ת נגיש/ה וזמין/ה? (0 משמעותו לא רלוונטי)": "נגישות וזמינות",

    # התנהלות בינאישית
    "באיזו מידה מגלה המפקד/ת אכפתיות כלפייך? (0 משמעותו לא רלוונטי)": "גילוי אכפתיות כלפי הצוער",
    "עד כמה המפקד/ת מתנהל/ת באופן מכבד? (0 משמעותו לא רלוונטי)": "התנהלות מכבדת",
    "עד כמה המפקד/ת מתנהל בפתיחות ובשקיפות? (0 משמעותו לא רלוונטי)": "התנהלות בפתיחות ושקיפות",

    # פיקודיות
    "באיזו מידה מגלה המפקד/ת סמכותיות וביטחון בהתנהלותו/ה? (0 משמעותו לא רלוונטי)": "סמכותיות וביטחון בהתנהלות",
    "עד כמה המפקד/ת רותמ/ת, מניע/ה ומעורר/ת בך מוטיבציה? (0 משמעותו לא רלוונטי)": "רותם, מניע ומעורר מוטיבציה",
    "עד כמה המפקד/ת מהווה מודל לחיקוי עבורי? (0 משמעותו לא רלוונטי)": "הצגת מודל לחיקוי",

    # פיתוח אישי ומקצועי
    "עד כמה שיחות אישיות עם המפקד/ת תורמות לך בעיבוד חוויות פיקודיות ומקצועיות בהכשרה? (0 משמעותו לא רלוונטי)": "שיחות אישיות תורמות לצוערים",
    "עד כמה תורמ/ת המפקד/ת לקידום ההכשרה שלך? (0 משמעותו לא רלוונטי)": "תורם להכשרת הצוערים",
    "עד כמה המפקד/ת מציב/ה סטנדרט גבוה, ודורש/ת חתירה למצוינות? (0 משמעותו לא רלוונטי)": "סטנדרט גבוה ודרישה לחתירה למצוינות",

    # שאלה כללית
    GENERAL_QUESTION_COLUMN: "עד כמה היית רוצה להיות תחת פיקודו בעתיד?",
}


MAX_QUESTIONS_PER_SECTION = 3

# =====================================
# Derived numeric collections
# =====================================

def build_numeric_question_columns() -> List[str]:
    cols: List[str] = []
    for questions in NUMERIC_SECTION_TO_QUESTIONS.values():
        cols.extend(questions)
    return cols


def build_numeric_question_to_section() -> Dict[str, str]:
    mapping: Dict[str, str] = {}
    for section_label, questions in NUMERIC_SECTION_TO_QUESTIONS.items():
        for col in questions:
            mapping[col] = section_label
    return mapping


NUMERIC_QUESTION_COLUMNS: List[str] = build_numeric_question_columns()
NUMERIC_QUESTION_TO_SECTION: Dict[str, str] = build_numeric_question_to_section()

# =====================================
# Placeholders for numeric stats
# =====================================

NUMERIC_PLACEHOLDER_SCOPE_COMMANDER: str = "commander"
NUMERIC_PLACEHOLDER_SCOPE_COHORT: str = "cohort"

NUMERIC_PH_MEAN_COMMANDER = "mean_commander"
NUMERIC_PH_STD_COMMANDER = "std_commander"
NUMERIC_PH_MEAN_COHORT = "mean_cohort"


def build_numeric_placeholders() -> Dict[str, Dict[str, str]]:
    mapping: Dict[str, Dict[str, str]] = {}

    for section_label, questions in NUMERIC_SECTION_TO_QUESTIONS.items():
        section_prefix = NUMERIC_SECTION_TO_PREFIX[section_label]

        for idx, question_col in enumerate(questions, start=1):
            base = f"{section_prefix}_{idx}"
            mapping[question_col] = {
                NUMERIC_PH_MEAN_COMMANDER: f"mean_{base}",
                NUMERIC_PH_STD_COMMANDER: f"std_{base}",
                NUMERIC_PH_MEAN_COHORT: f"cohort_mean_{base}",
            }

    return mapping


NUMERIC_PLACEHOLDERS: Dict[str, Dict[str, str]] = build_numeric_placeholders()

# =====================================
# Open-text questions and bullet lists
# =====================================

OPEN_TEXT_COLUMN_TO_BULLET_KEY: Dict[str, str] = {
    ":נקודות חוזק בתחום - חיבור מקצועי לתכני ההכשרה": "conserve_professional",
    ":נקודות חולשה בתחום - חיבור מקצועי לתכני ההכשרה": "improve_professional",

    ":נקודות חוזק בתחום - נוכחות ומעורבות": "conserve_personal",
    ":נקודות חולשה בתחום - נוכחות ומעורבות": "improve_personal",

    ":נקודות חוזק בתחום - התנהלות בינאישית": "conserve_interpersonal",
    ":נקודות חולשה בתחום - התנהלות בינאישית": "improve_interpersonal",

    ":נקודות חוזק בתחום - פיקודיות": "conserve_command",
    ":נקודות חולשה בתחום - פיקודיות": "improve_command",

    ":נקודות חוזק בתחום - פיתוח אישי ומקצועי": "conserve_development",
    ":נקודות חולשה בתחום - פיתוח אישי ומקצועי": "improve_development",

    "אנא נמק את בחירתך:": "general",
}


def build_open_text_columns() -> List[str]:
    return list(OPEN_TEXT_COLUMN_TO_BULLET_KEY.keys())


def build_bullet_list_context() -> Dict[str, str]:
    return {
        bullet_key: column_name
        for column_name, bullet_key in OPEN_TEXT_COLUMN_TO_BULLET_KEY.items()
    }


OPEN_QUESTIONS_COLUMNS: List[str] = build_open_text_columns()
BULLET_LIST_CONTEXT: Dict[str, str] = build_bullet_list_context()

# =====================================
# Expected Excel columns
# =====================================

def build_expected_excel_columns() -> List[str]:
    cols: List[str] = []
    cols.append("Timestamp")
    cols.append(COMMANDER_COLUMN)
    cols.extend(NUMERIC_QUESTION_COLUMNS)
    cols.extend(OPEN_QUESTIONS_COLUMNS)
    return cols


COLUMNS: List[str] = build_expected_excel_columns()

# =====================================
# Excel export constants (optional)
# =====================================

SHEET_NAME_QUANTITATIVE = "Quantitative"
SHEET_NAME_TEXTUAL = "Textual"

QUANT_HEADER_LABEL_COMMANDER = "Commander"
QUANT_HEADER_LABEL_NUM_RESPONDENTS = "Number of respondents"

QUANT_HEADER_ROWS = [
    ("meta", QUANT_HEADER_LABEL_COMMANDER, "commander_name"),
    ("meta", QUANT_HEADER_LABEL_NUM_RESPONDENTS, "num_answers"),
]

QUANT_COLUMN_QUESTION = "Question"
QUANT_COLUMN_COMMANDER_AVG = "Commander avg"
QUANT_COLUMN_COMMANDER_STD = "Commander std"
QUANT_COLUMN_COHORT_AVG = "Cohort avg"
QUANT_COLUMN_COHORT_STD = "Cohort std"
