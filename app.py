import json
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List, Tuple

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



def read_sheet_data(worksheet) -> Tuple[List[str], Dict[str, int], int]:
    """Читает заголовки, индекс кодов и номер первой пустой строки."""
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
    row_index_by_code: Dict[str, int] = {}
    for row_number, row in enumerate(values[1:], start=2):
        code = row[code_index].strip() if code_index < len(row) else ""
        if code:
            row_index_by_code[code] = row_number

    next_row_number = len(values) + 1
    return headers, row_index_by_code, next_row_number



def build_update_requests(headers: List[str], row_number: int, row_data: pd.Series) -> List[Dict[str, Any]]:
    """Готовит пакет обновлений для существующей строки."""
    requests: List[Dict[str, Any]] = []
    for column_name in ["Поставщик", "Менеджер"]:
        column_number = headers.index(column_name) + 1
        requests.append(
            {
                "range": rowcol_to_a1(row_number, column_number),
                "values": [[row_data[column_name]]],
            }
        )
    return requests



def build_new_row_values(headers: List[str], row_data: pd.Series) -> List[str]:
    """Формирует массив значений для новой строки."""
    new_row = [""] * max(len(headers), 7)
    for column_name in REQUIRED_COLUMNS:
        column_number = headers.index(column_name)
        new_row[column_number] = row_data[column_name]
    return new_row



def execute_update_requests(worksheet, update_requests: List[Dict[str, Any]]) -> None:
    """Выполняет пакет обновлений существующих строк одним запросом."""
    if not update_requests:
        return

    worksheet.batch_update(update_requests, value_input_option="USER_ENTERED")



def append_new_rows(worksheet, rows_to_append: List[List[str]], start_row_number: int) -> None:
    """Добавляет новые строки в таблицу одним запросом и окрашивает весь диапазон."""
    if not rows_to_append:
        return

    end_row_number = start_row_number + len(rows_to_append) - 1
    range_name = f"A{start_row_number}:{rowcol_to_a1(end_row_number, len(rows_to_append[0]))}"
    worksheet.update(range_name, rows_to_append, value_input_option="USER_ENTERED")
    apply_gray_fill(worksheet, start_row_number, end_row_number)



def apply_gray_fill(worksheet, start_row_number: int, end_row_number: int) -> None:
    """Применяет светло-серую заливку к столбцам 1-7 для диапазона новых строк."""
    worksheet.format(
        f"A{start_row_number}:G{end_row_number}",
        {
            "backgroundColor": {
                "red": 0.9,
                "green": 0.9,
                "blue": 0.9,
            }
        },
    )



def sync_excel_to_sheet(
    dataframe: pd.DataFrame,
    worksheet,
    dry_run: bool = False,
) -> List[str]:
    """Синхронизирует строки из Excel с Google Таблицей и возвращает лог выполнения."""
    headers, row_index_by_code, next_row_number = read_sheet_data(worksheet)
    logs: List[str] = []
    action_suffix = " (без записи)" if dry_run else ""
    update_requests: List[Dict[str, Any]] = []
    rows_to_append: List[List[str]] = []
    pending_new_rows_by_code: Dict[str, int] = {}
    next_new_row_offset = 0

    for index, row in dataframe.iterrows():
        try:
            code = row["Код"]
            if code in pending_new_rows_by_code:
                if not dry_run:
                    rows_to_append[pending_new_rows_by_code[code]] = build_new_row_values(headers, row)
                logs.append(f"Повтор кода в Excel: {code} → обновлена подготовленная новая строка{action_suffix}")
            elif code in row_index_by_code:
                if not dry_run:
                    update_requests.extend(
                        build_update_requests(headers, row_index_by_code[code], row)
                    )
                logs.append(f"Найден код: {code} → обновлено{action_suffix}")
            else:
                planned_row_number = next_row_number + next_new_row_offset
                row_index_by_code[code] = planned_row_number
                pending_new_rows_by_code[code] = next_new_row_offset
                next_new_row_offset += 1
                if not dry_run:
                    rows_to_append.append(build_new_row_values(headers, row))
                logs.append(f"Не найден код: {code} → добавлено{action_suffix}")
        except Exception as error:  # noqa: BLE001
            logs.append(f"Ошибка обработки строки {index + 2}: {error}")

    if not dry_run:
        execute_update_requests(worksheet, update_requests)
        append_new_rows(worksheet, rows_to_append, next_row_number)

    return logs



def render_logs(logs: List[str]) -> None:
    """Показывает лог выполнения в компактном окне с собственной прокруткой."""
    st.subheader("Лог выполнения")
    if not logs:
        st.info("Лог пока пуст. Загрузите файл и запустите обработку.")
        return

    log_text = "\n".join(logs)
    st.text_area(
        "Результат обработки",
        value=log_text,
        height=280,
        disabled=True,
        label_visibility="collapsed",
    )



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
    dry_run = st.checkbox(
        "Без записи",
        help="Если галочка включена, приложение только проверяет данные и пишет лог, но не изменяет Google Таблицу.",
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
            st.session_state.logs = sync_excel_to_sheet(dataframe, worksheet, dry_run=dry_run)
            if dry_run:
                st.success("Проверка в режиме 'Без записи' завершена")
            else:
                st.success("Обработка завершена")
        except Exception as error:  # noqa: BLE001
            st.session_state.logs = [f"Ошибка: {error}"]

    render_logs(st.session_state.logs)


if __name__ == "__main__":
    main()
