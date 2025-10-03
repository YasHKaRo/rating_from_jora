import gspread
from google.oauth2.service_account import Credentials

# Настройки
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SOURCE_SHEET_ID = '11BRMFjVlYCuVs5VrOgmWYat-wd-FAJYlvAWAGy3ioa4'  # ID таблицы от Жоры
DEST_SHEET_ID = '1jbDd2lNrzQVwrL75w_wuaKjAiKx5EFk8Ig4ullK0uto'    # ID нашей клановой таблицы

def auth_service_account():
    """Аутентификация через сервисный аккаунт"""
    creds = Credentials.from_service_account_file(
        'credentials.json',
        scopes=SCOPES
    )
    print(creds)
    return gspread.authorize(creds)

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
            except gspread.WorksheetNotFound:
                dst_worksheet = dest_sheet.add_worksheet(
                    title=src_worksheet.title,
                    rows=len(data),
                    cols=len(data[0]) if data else 1
                )

            # Обновляем данные
            if data:
                dst_worksheet.clear()
                dst_worksheet.update('A1', data)
                print(f"Обновлен лист: {src_worksheet.title}")

    except gspread.SpreadsheetNotFound:
        print("Ошибка: Не удалось получить доступ к таблице")
    except Exception as e:
        print(f"Произошла ошибка: {e}")

if __name__ == '__main__':
    sync_sheets()
