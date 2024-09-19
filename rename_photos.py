import pandas as pd
import os
import shutil

# Чтение Excel файла
file_path = 'Список студентов  ЦК на 17.09.2024.xlsx'
sheets = ['общий список', 'МО - БД (только ШК21)']

# Папки, где хранятся фото
folders = ['Форма в канале', 'Форма регистрации']

for sheet in sheets:
    # Чтение листа Excel
    df = pd.read_excel(file_path, sheet_name=sheet)

    # Убираем возможные лишние пробелы в названиях колонок
    df.columns = df.columns.str.strip()

    # Проходим по каждой строке
    for index, row in df.iterrows():
        status = row['Статус']
        email = row['E-mail корпоративный']  # Без лишнего пробела в названии колонки
        photo_link = row['Фото']

        # Проверяем статус
        if status == "Изменено имя":
            print(f"Строка {index} пропущена, статус уже 'Изменено имя'")
            continue  # Пропускаем, если уже обработано

        if pd.isna(photo_link):
            continue  # Пропускаем, если ссылка на фото отсутствует

        # Извлекаем часть до @ из e-mail
        email_name = email.split('@')[0]

        # Проверяем ссылку с "Yandex.Forms"
        if 'Yandex.Forms' in photo_link:
            # Находим подчеркивание и .jpg/.jpeg/.heic в ссылке
            start = photo_link.find('_')
            end_jpg = photo_link.find('.jpg')
            end_jpeg = photo_link.find('.jpeg')
            end_heic = photo_link.find('.heic')
            end_jfif = photo_link.find('.jfif')
            end_gif = photo_link.find('.gif')
            end_png = photo_link.find('.png')
            end_pdf = photo_link.find('.pdf')
            # Определяем, какое расширение присутствует
            if end_jpg != -1:
                end = end_jpg + 4  # Длина ".jpg"
            elif end_jpeg != -1:
                end = end_jpeg + 5  # Длина ".jpeg"
            elif end_heic != -1:
                end = end_heic + 5  # Длина ".heic"
            elif end_jfif != -1:
                end = end_jfif + 5 
            elif end_gif != -1:
                end = end_gif + 4
            elif end_png != -1:
                end = end_png + 4
            elif end_pdf != -1:
                end = end_pdf + 4 
            else:
                continue  # Пропускаем, если ни .jpg, ни .jpeg, ни .heic не найдено

            if start != -1 and end != -1:
                # Извлекаем имя файла вместе с расширением
                file_name_in_link = photo_link[start:end]

                # Убираем начальное подчеркивание, если оно есть
                if file_name_in_link.startswith('_'):
                    file_name_in_link = file_name_in_link[1:]

                # Поиск файла в локальных папках
                found_file = False
                for folder_path in folders:
                    for file_name in os.listdir(folder_path):
                        if file_name_in_link in file_name:
                            # Переименовываем файл
                            extension = file_name.split('.')[-1]  # Оставляем исходное расширение
                            new_file_name = f'{email_name}.{extension}'  # Переименовываем с оригинальным расширением
                            old_file_path = os.path.join(folder_path, file_name)
                            new_file_path = os.path.join(folder_path, new_file_name)
                            shutil.move(old_file_path, new_file_path)
                            print(f'Файл {file_name} переименован в {new_file_name}')

                            # Обновляем статус в DataFrame
                            df.at[index, 'Статус'] = 'Изменено имя'
                            found_file = True
                            break
                    if found_file:
                        break

        # Проверяем ссылку, которая начинается с "https://disk.yandex.ru/i/"
        elif 'disk.yandex.ru' in photo_link:
            print(f"Ссылка {photo_link} обработана: статус изменен на 'Изменено имя'")
            # Обновляем статус в DataFrame
            df.at[index, 'Статус'] = 'Изменено имя'

    # Сохраняем обновленный DataFrame в Excel
    with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=sheet, index=False)

