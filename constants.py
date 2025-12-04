from typing import Dict, List, Tuple

# ===============================
# Paths
# ===============================

INPUT_PATH: str = "answers.xlsx"
TEMPLATE_PATH: str = "mashov_sagaz_template.docx"
OUTPUT_PATH: str = "output/"
COMMANDER_EXCEL_OUTPUT_PATH: str = "output/excel/"

# ===============================
# Core survey columns / schema
# ===============================

COMMANDER_COLUMN: str = "ממלא/ת משוב על..."
GENERAL_QUESTION_COLUMN: str = "עד כמה הייתי רוצה להיות תחת פיקודו/ה של מפקד/ת הצוות או החוכנ/ת גם בעתיד?"
NONE_OF_THE_ABOVE_OPTION: str = "אף אחד מההיגדים אינו נכון בעיניי"


# Mapping from question text (Excel column) to its options
QUESTION_TO_OPTIONS: Dict[str, List[str]] = {
    # Command
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד)": [
        "החלטי/ת, סמכותי/ת ובטוח/ה בעצמו/ה",
        "מעורר/ת בי מוטיבציה",
        "מהווה דוגמה אישית",
        "מייצג/ת בהתנהגותו/ה את ערכי התוכנית",
        NONE_OF_THE_ABOVE_OPTION,
    ],

    # Involvement
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).1": [
        "נוכח/ת במופעי ההכשרה באופן רציף",
        "מעורב/ת במתרחש בתוכנית",
        "נגיש/ה וזמינ/ה לשאלות",
        "עוקב/ת אחר מצבי בהכשרה",
        NONE_OF_THE_ABOVE_OPTION,
    ],

    # Personal
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).2": [
        "מגלה אכפתיות כלפיי",
        "מכיר/ה אותי לעומק",
        "מתייחס/ת בנעימות ובכבוד",
        "אני מרגיש/ה שאני מסוגל/ת לשתף אותו",
        NONE_OF_THE_ABOVE_OPTION,
    ],

    # Challenge
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).3": [
        "נותן/ת משוב ישיר וכנה",
        "מסייע/ת בעיבוד חוויות והתנסויות בהכשרה",
        "דואג/ת לפתח ולקדם אותי",
        "מציב/ה לי סטנדרט גבוה",
        NONE_OF_THE_ABOVE_OPTION,
    ],
}


# Question text -> prefix used in placeholder names
# Example: "מה נכון בפיקוד?" -> "command" -> percent_command_1, total_command_1, ...
QUESTION_TO_PREFIX: Dict[str, str] = {
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד)":   "command",
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).1": "involvement",
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).2": "personal",
    "סמנ/י את ההיגדים שלדעתך נכונים, בנוגע למפקד/ת הצוות או החונכ/ת שלך (ניתן לסמן יותר מהיגד אחד).3": "challenge",
}

# Flat list of all option texts (kept for backwards compatibility)
OPTIONS: List[str] = [
    option_text
    for options_for_question in QUESTION_TO_OPTIONS.values()
    for option_text in options_for_question
]

# ===============================
# Open-text questions and bullets
# ===============================

# Excel column -> bullet-list key in Word template
# General / free text
OPEN_TEXT_COLUMN_TO_BULLET_KEY: Dict[str, str] = {
    # Command
    "מה הן החוזקות של מפקד/ת הצוות או החונכ/ת שלך בתחום הפיקוד?":   "conserve_command",
    "מה הן החולשות של מפקד/ת הצוות או החונכ/ת שלך בתחום הפיקוד?":   "improve_command",

    # Involvement
    "מה הן החוזקות של מפקד/ת הצוות או החונכ/ת שלך בתחום הנוכחות והמעורבות?":   "conserve_involvement",
    "מה הן החולשות של מפקד/ת הצוות או החונכ/ת שלך בתחום הנוכחות והמעורבות?":   "improve_involvement",

    # Personal
    "מה הן החוזקות של מפקד/ת הצוות או החונכ/ת שלך בתחום היחס האישי?":   "conserve_personal",
    "מה הן החולשות של מפקד/ת הצוות או החונכ/ת שלך בתחום היחס האישי?":   "improve_personal",

    # Challenge
    "מה הן החוזקות של מפקד/ת הצוות או החונכ/ת שלך בתחום האתגור והפיתוח המקצועי?":   "conserve_challenge",
    "מה הן החולשות של מפקד/ת הצוות או החונכ/ת שלך בתחום האתגור והפיתוח המקצועי?":   "improve_challenge",

    # General / free text
    "נמק/י את הציון שבחרת, ואת הסיבות לבחירתך בו": "general_comments",
    "פירוט:":                                      "rating_explanation",
}



OPEN_QUESTIONS_COLUMNS: List[str] = list(OPEN_TEXT_COLUMN_TO_BULLET_KEY.keys())


BULLET_LIST_CONTEXT: Dict[str, str] = {
    bullet_key: column_name
    for column_name, bullet_key in OPEN_TEXT_COLUMN_TO_BULLET_KEY.items()
}

OPEN_QUESTIONS_PLACEHOLDERS: Dict[str, str] = {}

# ===============================
# Multiple-choice columns
# ===============================

MULTIPLE_CHOICE_COLUMNS: List[str] = list(QUESTION_TO_OPTIONS.keys())

# ===============================
# Excel columns (old COLUMNS)
# ===============================

def _build_excel_columns() -> List[str]:
    columns: List[str] = []
    columns.append("Timestamp")
    columns.append(COMMANDER_COLUMN)
    columns.extend(MULTIPLE_CHOICE_COLUMNS)
    columns.extend(OPEN_QUESTIONS_COLUMNS)
    columns.append(GENERAL_QUESTION_COLUMN)
    return columns


COLUMNS: List[str] = _build_excel_columns()

# ===============================
# Placeholders for Word template
# ===============================

# Full placeholder list, kept as in the original code
PLACEHOLDERS: List[str] = [
    "name", "number_answers",

    "percent_command_1", "percent_command_2", "percent_command_3",
    "percent_command_4", "percent_command_5",
    "total_command_1", "total_command_2", "total_command_3",
    "total_command_4", "total_command_5",
    "strong_points_command", "weak_points_command",

    "percent_involvement_1", "percent_involvement_2", "percent_involvement_3",
    "percent_involvement_4", "percent_involvement_5",
    "total_involvement_1", "total_involvement_2", "total_involvement_3",
    "total_involvement_4", "total_involvement_5",
    "strong_points_involvement", "weak_points_involvement",

    "percent_personal_1", "percent_personal_2", "percent_personal_3",
    "percent_personal_4", "percent_personal_5",
    "total_personal_1", "total_personal_2", "total_personal_3",
    "total_personal_4", "total_personal_5",
    "strong_points_personal", "weak_points_personal",

    "percent_challenge_1", "percent_challenge_2", "percent_challenge_3",
    "percent_challenge_4", "percent_challenge_5",
    "total_challenge_1", "total_challenge_2", "total_challenge_3",
    "total_challenge_4", "total_challenge_5",
    "strong_points_challenge", "weak_points_challenge",

    "average_general",
    "std_general",
    "total_general",
]

# ===============================
# Placeholders mapping (new generated logic)
# ===============================

# Indices inside the (percent_placeholder, total_placeholder) tuple
PLACEHOLDER_TUPLE_INDEX_PERCENT: int = 0  # percent_*
PLACEHOLDER_TUPLE_INDEX_TOTAL: int = 1    # total_*

# Backwards-compatible aliases
PLACEHOLDER_TYPE_COMMANDER: int = PLACEHOLDER_TUPLE_INDEX_PERCENT
PLACEHOLDER_TYPE_COHORT: int = PLACEHOLDER_TUPLE_INDEX_TOTAL


def build_option_key(option_text: str, question_index: int) -> str:
    """
    Build the internal key for an option in OPTIONS_TO_PLACEHOLDERS.

    Regular options: key is the option text itself.
    'None of the above' option: key is '<text>_<question_index>'.
    """
    if option_text == NONE_OF_THE_ABOVE_OPTION:
        return f"{NONE_OF_THE_ABOVE_OPTION}_{question_index}"
    return option_text


def _build_percent_placeholder_name(prefix: str, slot_number: int) -> str:
    return f"percent_{prefix}_{slot_number}"


def _build_total_placeholder_name(prefix: str, slot_number: int) -> str:
    return f"total_{prefix}_{slot_number}"


def _build_placeholders_for_option(prefix: str, option_position: int) -> Tuple[str, str]:
    """
    Given a prefix (e.g. 'command') and zero-based option position,
    return (percent_placeholder, total_placeholder).
    """
    slot_number = option_position + 1  # convert to 1-based numbering
    percent_name = _build_percent_placeholder_name(prefix, slot_number)
    total_name = _build_total_placeholder_name(prefix, slot_number)
    return percent_name, total_name


def _build_options_to_placeholders_mapping() -> Dict[str, Tuple[str, str]]:
    """
    Build the mapping:
        option_key -> (percent_placeholder_name, total_placeholder_name)
    using QUESTION_TO_OPTIONS and QUESTION_TO_PREFIX.
    """
    mapping: Dict[str, Tuple[str, str]] = {}

    for question_index, (question_text, options) in enumerate(QUESTION_TO_OPTIONS.items()):
        prefix = QUESTION_TO_PREFIX[question_text]

        for option_position, option_text in enumerate(options):
            key = build_option_key(option_text, question_index)
            mapping[key] = _build_placeholders_for_option(prefix, option_position)

    return mapping


OPTIONS_TO_PLACEHOLDERS: Dict[str, Tuple[str, str]] = _build_options_to_placeholders_mapping()

# ===============================
# General numeric / RTL constants
# ===============================

MIN_GENERAL_ANSWERS: int = 4
TOO_FEW_ANSWERS_TEXT: str = "ענו פחות מ-4"

PERCENT_DECIMALS: int = 1
MIN_COMMENT_LENGTH: int = 2
DEFAULT_ZERO_VALUE: float = 0.0

RLE: str = "\u202B"  # Right-to-Left Embedding
PDF: str = "\u202C"  # Pop Directional Formatting
RLM: str = "\u200F"  # Right-to-Left Mark

PUNCTUATION_CHARS: List[str] = [
    ",", ".", "\"", "\\", "-", ":", ";", "(", ")", "!", "?", "+", "/",
]

# ===============================
# Excel export constants
# ===============================

SHEET_NAME_QUANTITATIVE: str = "Quantitative"
SHEET_NAME_TEXTUAL: str = "Textual"

OPTIONS_PER_QUESTION: int = 5
OPTION_BLOCK_WIDTH: int = 3

QUANT_HEADER_LABEL_COMMANDER: str = "Commander"
QUANT_HEADER_LABEL_NUM_RESPONDENTS: str = "Number of respondents"
QUANT_HEADER_LABEL_AVG_GENERAL: str = "General question – commander average"
QUANT_HEADER_LABEL_STD_GENERAL: str = "General question – commander std"
QUANT_HEADER_LABEL_COHORT_GENERAL: str = "General question – cohort average"

QUANT_HEADER_ROWS = [
    ("meta",        QUANT_HEADER_LABEL_COMMANDER,        "commander_name"),
    ("meta",        QUANT_HEADER_LABEL_NUM_RESPONDENTS,  "num_answers"),
    ("placeholder", QUANT_HEADER_LABEL_AVG_GENERAL,      "average_general"),
    ("placeholder", QUANT_HEADER_LABEL_STD_GENERAL,      "std_general"),
    ("placeholder", QUANT_HEADER_LABEL_COHORT_GENERAL,   "total_general"),
]

QUANT_COLUMN_QUESTION: str = "Question"

QUANT_SUBHEADER_STATEMENT: str = "Statement"
QUANT_SUBHEADER_COMMANDER_PERCENT: str = "Commander %"
QUANT_SUBHEADER_COHORT_PERCENT: str = "Cohort %"
