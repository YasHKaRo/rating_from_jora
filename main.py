import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Настройки
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SOURCE_SHEET_ID = '11BRMFjVlYCuVs5VrOgmWYat-wd-FAJYlvAWAGy3ioa4'  # ID таблицы от Жоры
DEST_SHEET_ID = '1jbDd2lNrzQVwrL75w_wuaKjAiKx5EFk8Ig4ullK0uto'    # ID нашей клановой таблицы

# Цвета для фильтрации
ACTIVE_COLOR = {'red': 207/255, 'green': 255/255, 'blue': 233/255}  # #cfffe9
ABSENT_COLOR = {'red': 255/255, 'green': 211/255, 'blue': 199/255}  # #ffd3c7

def auth_service_account():
    """Аутентификация через сервисный аккаунт"""
    creds = Credentials.from_service_account_file(
        'credentials.json',
        scopes=SCOPES
    )
    print(creds)
    return gspread.authorize(creds)

def get_sheets_service():
    """Создает сервис для работы с Google Sheets API"""
    creds = Credentials.from_service_account_file(
        'credentials.json',
        scopes=SCOPES
    )
    return build('sheets', 'v4', credentials=creds)

def rgb_to_hex(rgb_color):
    """Конвертирует RGB в HEX строку"""
    if not rgb_color:
        return None
    r = int(rgb_color.get('red', 0) * 255)
    g = int(rgb_color.get('green', 0) * 255)
    b = int(rgb_color.get('blue', 0) * 255)
    return f'#{r:02x}{g:02x}{b:02x}'

def colors_are_similar(color1, color2, tolerance=0.01):
    """Сравнивает два цвета с допуском"""
    if not color1 or not color2:
        return False
    
    return (abs(color1.get('red', 0) - color2.get('red', 0)) < tolerance and
            abs(color1.get('green', 0) - color2.get('green', 0)) < tolerance and
            abs(color1.get('blue', 0) - color2.get('blue', 0)) < tolerance)

def get_players_data(client):
    """Получение данных об игроках с фильтрацией по цвету"""
    try:
        # Открываем исходную таблицу
        source_sheet = client.open_by_key(SOURCE_SHEET_ID)
        wars_worksheet = source_sheet.worksheet("Wars")
        
        # Получаем все данные
        all_values = wars_worksheet.get_all_values()
        
        # Создаем сервис для работы с форматом ячеек
        sheets_service = get_sheets_service()
        
        # Определяем диапазон для получения формата
        last_row = len(all_values)
        if last_row < 4:
            return {}
        
        # Запрашиваем информацию о формате ячеек
        request = sheets_service.spreadsheets().get(
            spreadsheetId=SOURCE_SHEET_ID,
            ranges=[f"Wars!A4:A{last_row}"],
            fields="sheets(data(rowData(values(effectiveFormat.backgroundColor))))"
        )
        response = request.execute()
        
        # Извлекаем информацию о цветах
        sheets = response.get('sheets', [])
        if not sheets:
            return {}
        
        data = sheets[0].get('data', [])
        if not data:
            return {}
        
        row_data = data[0].get('rowData', [])
        
        all_players = {}
        
        # Проходим по всем строкам начиная с A4
        for i in range(3, last_row):
            row = all_values[i]
            
            # Проверяем, что есть хотя бы 2 колонки (A и B)
            if len(row) >= 2 and row[0] and row[1]:
                player_name = row[0].strip()
                player_tag = row[1].strip()
                
                # Получаем цвет ячейки A для этой строки
                if i - 3 < len(row_data) and row_data[i - 3]:
                    cell_data = row_data[i - 3].get('values', [])
                    if cell_data and cell_data[0]:
                        bg_color = cell_data[0].get('effectiveFormat', {}).get('backgroundColor', {})
                        
                        # Проверяем, совпадает ли цвет с активными игроками
                        if colors_are_similar(bg_color, ACTIVE_COLOR):
                            all_players[player_tag] = player_name
        
        print(f"Найдено активных игроков: {len(all_players)}")
        return all_players
        
    except Exception as e:
        print(f"Произошла ошибка при получении данных игроков: {e}")
        return {}

def save_players_to_wars(client, players_dict):
    """Сохраняет словарь игроков в лист Wars, очищая его предварительно"""
    try:
        # Открываем вашу таблицу
        dest_sheet = client.open_by_key(DEST_SHEET_ID)
        
        # Получаем или создаем лист "Wars"
        try:
            wars_worksheet = dest_sheet.worksheet("Wars")
        except gspread.WorksheetNotFound:
            # Создаем лист, если он не существует
            wars_worksheet = dest_sheet.add_worksheet(
                title="Wars", 
                rows=1000, 
                cols=2
            )
        
        # ОЧИСТКА: Полностью очищаем лист
        wars_worksheet.clear()
        print("Лист 'Wars' очищен")
        
        # Добавляем заголовки
        headers = [["Player tag", "Player name"]]
        wars_worksheet.update(values=headers, range_name='A1:B1')
        print("Заголовки добавлены")
        
        # Подготавливаем данные для записи
        data = []
        for player_tag, player_name in players_dict.items():
            data.append([player_tag, player_name])  # Имя в колонке A, тег в колонке B
        
        # Записываем данные, начиная со строки 2 (после заголовков)
        if data:
            wars_worksheet.update(values=data, range_name=f'A2:B{len(data)+1}')
            print(f"Записано {len(data)} игроков в лист 'Wars'")
        else:
            print("Нет данных для записи")
            
        # Форматирование (опционально)
        try:
            # Жирный шрифт для заголовков
            wars_worksheet.format('A1:B1', {
                'textFormat': {'bold': True},
                'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9}
            })
            print("Форматирование применено")
        except Exception as e:
            print(f"Не удалось применить форматирование: {e}")
            
        return True
        
    except Exception as e:
        print(f"Ошибка при сохранении в лист 'Wars': {e}")
        return False

def sync_sheets():
    client = auth_service_account()
    try:
        # Пытаемся открыть таблицу жоры
        source_sheet = client.open_by_key(SOURCE_SHEET_ID)
        print("Я смог подключиться к таблице Жоры!")

        # Открываем таблицу клана
        dest_sheet = client.open_by_key(DEST_SHEET_ID)
        print("Успешный доступ к таблице клана!")

        # Синхронизация данных
        for src_worksheet in source_sheet.worksheets():
            data = src_worksheet.get_all_values()
            # Проверяем существование листа в целевой таблице
            try:
                dst_worksheet = dest_sheet.worksheet(src_worksheet.title)
                print(src_worksheet.title)
            except gspread.WorksheetNotFound:
                dst_worksheet = dest_sheet.add_worksheet(
                    title=src_worksheet.title,
                    rows=len(data),
                    cols=len(data[0]) if data else 1
                )

            # Обновляем данные
            if src_worksheet.title == "Wars":
                # Получаем данные игроков и сохраняем в нашу таблицу
                pl_dict = get_players_data(client)
                save_players_to_wars(client, pl_dict)
                print("Начальник, игроки были получены!")
            elif data:
                dst_worksheet.clear()
                dst_worksheet.update(values=data, range_name='A1')
                print(f"Обновлен лист: {src_worksheet.title}")

    except gspread.SpreadsheetNotFound:
        print("Ошибка: Не удалось получить доступ к таблице")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

if __name__ == '__main__':
    sync_sheets()
