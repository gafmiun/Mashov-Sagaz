COLUMNS = ["Timestamp", "שם המפקד:",
           "מה נכון בפיקוד?", "נקודות חיוביות פיקוד:", "נקודות שליליות פיקוד:",
           "מה נכון בנוכחות ומעורבות?", "נקודות חיוביות נוכחות:", "נקודות שליליות נוכחות:",
           "מה נכון ביחס אישי?", "נקודות חיוביות יחס אישי:", "נקודות שליליות יחס אישי:",
           "מה נכון באתגור ופיתוח מקצועי?", "נקודות חיוביות אתגור:", "נקודות שליליות אתגור:",
           "עד כמה הייתי רוצה להיות תחת פיקודו בעתיד?", "הערות כלליות:"]

INPUT_PATH = "answers.xlsx"

TEMPLATE_PATH = "new_template2.docx"

OUTPUT_PATH = "output/"

COMMANDER_COLUMN = "שם המפקד:"

PLACEHOLDERS = [

                "name", "number_answers",

                "percent_command_1", "percent_command_2", "percent_command_3", "percent_command_4", "percent_command_5",
                 "total_command_1", "total_command_2", "total_command_3", "total_command_4", "total_command_5",
                 "strong_points_command", "weak_points_command",

                "percent_involvement_1", "percent_involvement_2", "percent_involvement_3", "percent_involvement_4", "percent_involvement_5",
                "total_involvement_1", "total_involvement_2", "total_involvement_3", "total_involvement_4", "total_involvement_5",
                "strong_points_involvement", "weak_points_involvement",

                "percent_personal_1", "percent_personal_2", "percent_personal_3", "percent_personal_4", "percent_personal_5",
                "total_personal_1", "total_personal_2", "total_personal_3", "total_personal_4", "total_personal_5",
                "strong_points_personal", "weak_points_personal",

                "percent_challenge_1", "percent_challenge_2", "percent_challenge_3", "percent_challenge_4", "percent_challenge_5",
                "total_challenge_1", "total_challenge_2", "total_challenge_3", "total_challenge_4", "total_challenge_5",
                "strong_points_challenge", "weak_points_challenge",

                "average_general",
                 "std_general",
                 "total_general"

                ]

OPTIONS = ["החלטי/ת, סמכותי/ת ובטוח/ה בעצמו/ה", "מעורר/ת בי מוטיבציה", "מהווה דוגמה אישית", "מייצג/ת בהתנהגותו/ה את ערכי התוכנית", "אף אחד מההיגדים אינו נכון בעיניי",
           "נוכח/ת במופעי ההכשרה באופן רציף", "מעורב/ת במתרחש בתוכנית", "נגיש/ה וזמינ/ה לשאלות", "עוקב/ת אחר מצבי בהכשרה", "אף אחד מההיגדים אינו נכון בעיניי",
           "מגלה אכפתיות כלפיי", "מכיר/ה אותי לעומק", "מתייחס/ת בנעימות ובכבוד", "אני מרגיש/ה שאני מסוגל/ת לשתף אותו", "אף אחד מההיגדים אינו נכון בעיניי",
           "נותן/ת משוב ישיר וכנה", "מסייע/ת בעיבוד חוויות והתנסויות בהכשרה", "דואג/ת לפתח ולקדם אותי", "מציב/ה לי סטנדרט גבוה", "אף אחד מההיגדים אינו נכון בעיניי",
           ]

OPTIONS_TO_PLACEHOLDERS = {
                            "החלטי/ת, סמכותי/ת ובטוח/ה בעצמו/ה": ("percent_command_1", "total_command_1"),
                           "מעורר/ת בי מוטיבציה": ("percent_command_2", "total_command_2"),
                           "מהווה דוגמה אישית": ("percent_command_3", "total_command_3"),
                            "מייצג/ת בהתנהגותו/ה את ערכי התוכנית": ("percent_command_4", "total_command_4"),
                            "אף אחד מההיגדים אינו נכון בעיניי_0": ("percent_command_5", "total_command_5"),
                            "נוכח/ת במופעי ההכשרה באופן רציף": ("percent_involvement_1", "total_involvement_1"),
                            "מעורב/ת במתרחש בתוכנית": ("percent_involvement_2", "total_involvement_2"),
                            "נגיש/ה וזמינ/ה לשאלות": ("percent_involvement_3", "total_involvement_3"),
                            "עוקב/ת אחר מצבי בהכשרה": ("percent_involvement_4", "total_involvement_4"),
                            "אף אחד מההיגדים אינו נכון בעיניי_1": ("percent_involvement_5", "total_involvement_5"),
                            "מגלה אכפתיות כלפיי": ("percent_personal_1", "total_personal_1"),
                            "מכיר/ה אותי לעומק": ("percent_personal_2", "total_personal_2"),
                            "מתייחס/ת בנעימות ובכבוד": ("percent_personal_3", "total_personal_3"),
                            "אני מרגיש/ה שאני מסוגל/ת לשתף אותו": ("percent_personal_4", "total_personal_4"),
                            "אף אחד מההיגדים אינו נכון בעיניי_2": ("percent_personal_5", "total_personal_5"),
                            "נותן/ת משוב ישיר וכנה": ("percent_challenge_1", "total_challenge_1"),
                            "מסייע/ת בעיבוד חוויות והתנסויות בהכשרה": ("percent_challenge_2", "total_challenge_2"),
                            "דואג/ת לפתח ולקדם אותי": ("percent_challenge_3", "total_challenge_3"),
                            "מציב/ה לי סטנדרט גבוה": ("percent_challenge_4", "total_challenge_4"),
                            "אף אחד מההיגדים אינו נכון בעיניי_3": ("percent_challenge_5", "total_challenge_5"),
                            }

NONE_OF_THE_ABOVE_OPTION = "אף אחד מההיגדים אינו נכון בעיניי"

MULTIPLE_CHOICE_COLUMNS = [
    "מה נכון בפיקוד?",
    "מה נכון בנוכחות ומעורבות?",
    "מה נכון ביחס אישי?",
    "מה נכון באתגור ופיתוח מקצועי?"
]

OPEN_QUESTIONS_COLUMNS = ["נקודות חיוביות פיקוד:", "נקודות שליליות פיקוד:",
                          "נקודות חיוביות נוכחות:", "נקודות שליליות נוכחות:",
                          "נקודות חיוביות יחס אישי:", "נקודות שליליות יחס אישי:",
                          "נקודות חיוביות אתגור:", "נקודות שליליות אתגור:",
                          "הערות כלליות:"]


OPEN_QUESTIONS_PLACEHOLDERS = {}

# Hebrew Excel column → English placeholder name in Word template
BULLET_LIST_CONTEXT = {
    # פיקוד (Command)
    "conserve_command": "נקודות חיוביות פיקוד:",
    "improve_command":  "נקודות שליליות פיקוד:",

    # נוכחות ומעורבות (Involvement)
    "conserve_involvement": "נקודות חיוביות נוכחות:",
    "improve_involvement":  "נקודות שליליות נוכחות:",

    # יחס אישי (Personal)
    "conserve_personal": "נקודות חיוביות יחס אישי:",
    "improve_personal":  "נקודות שליליות יחס אישי:",

    # אתגור ופיתוח מקצועי (Challenge)
    "conserve_challenge": "נקודות חיוביות אתגור:",
    "improve_challenge":  "נקודות שליליות אתגור:",

    # הערות כלליות (General)
    "general_comments": "הערות כלליות:",
}


# ===== General constants =====

# Column with the 1–6 general rating
GENERAL_QUESTION_COLUMN = "עד כמה הייתי רוצה להיות תחת פיקודו בעתיד?"


MIN_GENERAL_ANSWERS = 4


TOO_FEW_ANSWERS_TEXT = "ענו פחות מ-4"


PERCENT_DECIMALS = 2


MIN_COMMENT_LENGTH = 2

DEFAULT_ZERO_VALUE = 0.0

RLE = '\u202B'  # Right-to-Left Embedding
PDF = '\u202C'  # Pop Directional Formatting
RLM = '\u200F'

PUNCTUATION_CHARS = [
    ",", ".", "\"", "\\", "-", ":", ";", "(", ")", "!", "?", "+", "/"
]
# ===== Excel export constants =====

COMMANDER_EXCEL_OUTPUT_PATH = "output/excel/"

SHEET_NAME_QUANTITATIVE = "Quantitative"
SHEET_NAME_TEXTUAL = "Textual"

OPTIONS_PER_QUESTION = 5
OPTION_BLOCK_WIDTH = 3

QUANT_HEADER_LABEL_COMMANDER = "Commander"
QUANT_HEADER_LABEL_NUM_RESPONDENTS = "Number of respondents"
QUANT_HEADER_LABEL_AVG_GENERAL = "General question – commander average"
QUANT_HEADER_LABEL_STD_GENERAL = "General question – commander std"
QUANT_HEADER_LABEL_COHORT_GENERAL = "General question – cohort average"

QUANT_COLUMN_QUESTION = "Question"

QUANT_SUBHEADER_STATEMENT = "Statement"
QUANT_SUBHEADER_COMMANDER_PERCENT = "Commander %"
QUANT_SUBHEADER_COHORT_PERCENT = "Cohort %"
