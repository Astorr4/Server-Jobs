import warnings
from Functions import delete_replace_files, Custom_table, Create_table
# Исключение предупреждения об устаревших функциях
warnings.filterwarnings('ignore', category=FutureWarning)
# Удаление старого файла с джобами из папки Old
delete_replace_files("N:\\folder\\path\\Jobs\\Old",
                     "N:\\folder\\path\\Jobs")
# Создание таблицы
Create_table("N:\\folder\\path\\Jobs_ASK\\ASK_Jobs.xlsx", "N:\\folder\\path\\Jobs\\Jobs(VBA).xlsm",
             'N:\\folder\\path\\Jobs_ASK\\Files\\Jobs routes.xlsx')
# Форматирование таблицы
Custom_table("N:\\folder\\path\\Jobs\\Jobs(VBA).xlsm")
