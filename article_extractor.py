import os

# Получите путь к директории, где находится текущий скрипт
script_directory = os.path.dirname(os.path.abspath(__file__))

# Получите список файлов в директории
all_files = os.listdir(script_directory)

# Отфильтруйте только .txt файлы
txt_files = [file for file in all_files if file.endswith('.txt')]

# Путь к главному XLSX файлу
main_xlsx_file_name = 'catalog-full.xlsx'

# Импортируйте и используйте функцию create_xlsx_from_txt для каждого файла
import pandas as pd

def create_xlsx_from_txt(txt_file_name, main_xlsx_file_path):
    try:
        # Определите, в каком столбце находятся артикулы в главном XLSX файле
        article_column_name = 'Артикул'  # Замените на имя столбца с артикулами

        # Получите абсолютные пути к файлам
        txt_file_path = os.path.join(script_directory, txt_file_name)
        main_xlsx_file_path = os.path.join(script_directory, main_xlsx_file_name)

        # Создайте новый XLSX файл с таким же именем, как у TXT, но с расширением .xlsx
        xlsx_file_name = os.path.splitext(txt_file_name)[0] + '.xlsx'
        xlsx_file_path = os.path.join(script_directory, xlsx_file_name)

        # Прочитайте главный XLSX файл в DataFrame
        main_df = pd.read_excel(main_xlsx_file_path)

        # Прочитайте артикулы из TXT файла
        with open(txt_file_path, 'r') as txt_file:
            selected_articles = [line.strip().upper() for line in txt_file]

        # Отфильтруйте строки по артикулам
        result_df = main_df[main_df[article_column_name].str.strip().str.upper().isin(selected_articles)]

        # Сохраните новый XLSX файл
        result_df.to_excel(xlsx_file_path, index=False)
        return True, xlsx_file_path
    except Exception as e:
        return False, str(e)

for txt_file in txt_files:
    success, new_xlsx_name = create_xlsx_from_txt(txt_file, main_xlsx_file_name)
    if success:
        print(f'Создан новый XLSX файл: {new_xlsx_name}')
    else:
        print(f'Произошла ошибка при создании XLSX файла: {new_xlsx_name}')

