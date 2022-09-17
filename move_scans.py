from msilib.schema import Error
import os
import shutil


SCANER_FOLDER = 'C:/Test' # папка где лежат сканы
ARCHIVE_FOLDER = 'C:/Test/Dest' # папка куда нужно их переместить
FILTER = ['.jpg', '.png'] # расширения файлов которые нужно переместить


move_list = []
file_list = os.listdir(SCANER_FOLDER)


for file in file_list:
    for filter_el in FILTER:
        if filter_el in file:
            move_list.append(file)

for file in move_list:
    try:
        shutil.move(f'{SCANER_FOLDER}/{file}', ARCHIVE_FOLDER)
    except Exception as e:
        print(e)
