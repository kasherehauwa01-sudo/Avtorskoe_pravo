import json
from io import BytesIO
from pathlib import Path
from typing import List, Tuple

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

REQUIRED_COLUMNS = ["Код", "Поставщик", "Менеджер"]
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
DEFAULT_SHEET_ID = "19hfmYJtv9FCzS6vSJ6OcwTRK8TuSNfgjCjEdF8JRs1Y"
CONFIG_PATH = Path("config.json")
CREDENTIALS_PATH = Path("credentials.json")


def load_config_sheet_id() -> str:
    """Читает ID Google Таблицы из config.json, если файл существует."""
    if not CONFIG_PATH.exists():
        return DEFAULT_SHEET_ID

    with CONFIG_PATH.open("r", encoding="utf-8") as file:
        data = json.load(file)

    return str(data.get("google_sheet_id", DEFAULT_SHEET_ID)).strip() or DEFAULT_SHEET_ID



def load_excel_file(uploaded_file) -> pd.DataFrame:
    """Загружает Excel-файл, проверяет обязательные колонки и нормализует данные."""
    suffix = Path(uploaded_file.name).suffix.lower()
    content = uploaded_file.getvalue()
    buffer = BytesIO(content)

    if suffix == ".xlsx":
        dataframe = pd.read_excel(buffer, dtype=str)
    elif suffix == ".xls":
        dataframe = pd.read_excel(buffer, engine="xlrd", dtype=str)
    else:
        raise ValueError("Поддерживаются только файлы формата .xls и .xlsx")

    dataframe.columns = [str(column).strip() for column in dataframe.columns]
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in dataframe.columns]
    if missing_columns:
        raise ValueError(
            "В Excel отсутствуют обязательные столбцы: " + ", ".join(missing_columns)
        )

    dataframe = dataframe[REQUIRED_COLUMNS].copy()
    dataframe = dataframe.fillna("")

    for column in REQUIRED_COLUMNS:
        dataframe[column] = dataframe[column].astype(str).str.strip()

    dataframe = dataframe[dataframe["Код"] != ""]
    return dataframe.reset_index(drop=True)



def load_service_account_info() -> dict:
    """Получает данные service account из Streamlit secrets или credentials.json."""
    secret_section = st.secrets.get("gcp_service_account")
    if secret_section:
        return dict(secret_section)

    if CREDENTIALS_PATH.exists():
        with CREDENTIALS_PATH.open("r", encoding="utf-8") as file:
            return json.load(file)

    raise FileNotFoundError(
        "Не найдены данные сервисного аккаунта. Добавьте их в Streamlit secrets "
        "(секция gcp_service_account) или создайте файл credentials.json в корне проекта."
    )


@st.cache_resource(show_spinner=False)
def connect_to_google() -> gspread.Client:
    """Создает авторизованное подключение к Google Sheets API."""
    service_account_info = load_service_account_info()
    credentials = Credentials.from_service_account_info(
        service_account_info,
        scopes=SCOPES,
    )
    return gspread.authorize(credentials)



def get_worksheet(sheet_id: str):
    """Открывает Google Таблицу по ID и возвращает первый лист."""
    client = connect_to_google()
    spreadsheet = client.open_by_key(sheet_id)
    return spreadsheet.sheet1



def read_sheet_data(worksheet) -> Tuple[List[str], dict]:
    """Читает заголовки и создает индекс строк по значению столбца 'Код'."""
    values = worksheet.get_all_values()
    if not values:
        raise ValueError(
            "Google Таблица пустая. Добавьте строку заголовков: Код, Поставщик, Менеджер."
        )

    headers = [header.strip() for header in values[0]]
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in headers]
    if missing_columns:
        raise ValueError(
            "В Google Таблице отсутствуют обязательные столбцы: "
            + ", ".join(missing_columns)
        )

    code_index = headers.index("Код")
    row_index_by_code = {}
    for row_number, row in enumerate(values[1:], start=2):
        code = row[code_index].strip() if code_index < len(row) else ""
        if code:
            row_index_by_code[code] = row_number

    return headers, row_index_by_code



def update_existing_row(
    worksheet,
    row_number: int,
    headers: List[str],
    row_data: pd.Series,
) -> None:
    """Обновляет значения столбцов 'Поставщик' и 'Менеджер' для найденного кода."""
    updates = []
    for column_name in ["Поставщик", "Менеджер"]:
        column_number = headers.index(column_name) + 1
        updates.append(
            {
                "range": rowcol_to_a1(row_number, column_number),
                "values": [[row_data[column_name]]],
            }
        )

    worksheet.batch_update(updates, value_input_option="USER_ENTERED")



def add_new_row(worksheet, headers: List[str], row_data: pd.Series) -> int:
    """Добавляет новую строку в первую пустую строку и возвращает номер строки."""
    existing_values = worksheet.get_all_values()
    row_number = len(existing_values) + 1

    new_row = [""] * max(len(headers), 7)
    for column_name in REQUIRED_COLUMNS:
        column_number = headers.index(column_name)
        new_row[column_number] = row_data[column_name]

    range_name = f"A{row_number}:{rowcol_to_a1(row_number, len(new_row))}"
    worksheet.update(range_name, [new_row], value_input_option="USER_ENTERED")
    apply_gray_fill(worksheet, row_number)
    return row_number



def apply_gray_fill(worksheet, row_number: int) -> None:
    """Применяет светло-серую заливку к столбцам 1-7 добавленной строки."""
    worksheet.format(
        f"A{row_number}:G{row_number}",
        {
            "backgroundColor": {
                "red": 0.9,
                "green": 0.9,
                "blue": 0.9,
            }
        },
    )



def sync_excel_to_sheet(dataframe: pd.DataFrame, worksheet) -> List[str]:
    """Синхронизирует строки из Excel с Google Таблицей и возвращает лог выполнения."""
    headers, row_index_by_code = read_sheet_data(worksheet)
    logs: List[str] = []

    for index, row in dataframe.iterrows():
        try:
            code = row["Код"]
            if code in row_index_by_code:
                update_existing_row(worksheet, row_index_by_code[code], headers, row)
                logs.append(f"Найден код: {code} → обновлено")
            else:
                row_number = add_new_row(worksheet, headers, row)
                row_index_by_code[code] = row_number
                logs.append(f"Не найден код: {code} → добавлено")
        except Exception as error:  # noqa: BLE001
            logs.append(f"Ошибка обработки строки {index + 2}: {error}")

    return logs



def render_logs(logs: List[str]) -> None:
    """Показывает лог выполнения в интерфейсе Streamlit."""
    st.subheader("Лог выполнения")
    if not logs:
        st.info("Лог пока пуст. Загрузите файл и запустите обработку.")
        return

    for message in logs:
        if message.startswith("Ошибка"):
            st.error(message)
        elif "добавлено" in message:
            st.warning(message)
        else:
            st.success(message)



def main() -> None:
    """Главная функция Streamlit-приложения."""
    st.set_page_config(page_title="Синхронизация Excel и Google Sheets", page_icon="📄")
    st.title("Синхронизация Excel-файла с Google Таблицей")
    st.write("Загрузите xls таблицу со столбцами: Код, Поставщик, Менеджер")

    sheet_id = load_config_sheet_id()
    st.caption(f"Google Таблица по умолчанию: {sheet_id}")

    uploaded_file = st.file_uploader(
        "Выберите Excel-файл",
        type=["xls", "xlsx"],
        accept_multiple_files=False,
    )

    if "logs" not in st.session_state:
        st.session_state.logs = []

    if st.button("Запустить обработку", type="primary"):
        try:
            if not sheet_id:
                raise ValueError("Не удалось определить ID Google Таблицы из настроек приложения")
            if uploaded_file is None:
                raise ValueError("Загрузите Excel-файл перед запуском обработки")

            dataframe = load_excel_file(uploaded_file)
            worksheet = get_worksheet(sheet_id)
            st.session_state.logs = sync_excel_to_sheet(dataframe, worksheet)
            st.success("Обработка завершена")
        except Exception as error:  # noqa: BLE001
            st.session_state.logs = [f"Ошибка: {error}"]

    render_logs(st.session_state.logs)


if __name__ == "__main__":
    main()
