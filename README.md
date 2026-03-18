# Excel → Google Sheets Sync

Приложение синхронизирует Excel-файл с Google Таблицей по колонке "Код".

## Возможности

- загрузка файлов `.xls` и `.xlsx`;
- поиск строк в Google Таблице по колонке `Код`;
- обновление значений `Поставщик` и `Менеджер`, если код найден;
- добавление новой строки, если код отсутствует;
- выделение новых строк светло-серым цветом в диапазоне столбцов `1–7`;
- вывод подробного лога обработки в интерфейсе Streamlit;
- подключение к Google Sheets API через `Streamlit secrets`;
- режим `Без записи`, в котором приложение только проверяет действия и пишет лог без изменения таблицы.

## Структура проекта

```text
/project
  app.py
  requirements.txt
  config.json
  .streamlit/
    secrets.toml
  README.md
```

## Установка

```bash
pip install -r requirements.txt
```

## Запуск

```bash
streamlit run app.py
```

## Подключение к Google Sheets API через Streamlit secrets

1. Перейдите в Google Cloud Console:  
   https://console.cloud.google.com/
2. Создайте новый проект:  
   `Select Project` → `New Project` → укажите имя проекта → `Create`.
3. Включите API:
   - откройте `APIs & Services` → `Library`;
   - найдите и включите `Google Sheets API`;
   - найдите и включите `Google Drive API`.
4. Создайте Service Account:
   - `APIs & Services` → `Credentials`;
   - `Create Credentials` → `Service Account`;
   - введите имя → `Create`.
5. Выдайте сервисному аккаунту роль `Editor`.
6. Создайте JSON-ключ:
   - откройте созданный Service Account;
   - вкладка `Keys`;
   - `Add Key` → `Create new key`;
   - формат `JSON`;
   - скачайте файл.
7. Создайте в проекте папку `.streamlit`, если ее еще нет.
8. Создайте файл `.streamlit/secrets.toml`.
9. Скопируйте значения из скачанного JSON-ключа в секцию `gcp_service_account`.

Пример `.streamlit/secrets.toml`:

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-private-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\nYOUR_PRIVATE_KEY\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project-id.iam.gserviceaccount.com"
client_id = "1234567890"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/your-service-account%40your-project-id.iam.gserviceaccount.com"
universe_domain = "googleapis.com"
```

> Если вы разворачиваете приложение в Streamlit Community Cloud, эти же поля нужно добавить в разделе `App settings` → `Secrets`.

## Доступ к Google Таблице

1. Откройте нужную Google Таблицу.
2. Нажмите `Поделиться`.
3. Возьмите значение `client_email` из `.streamlit/secrets.toml`.
4. Добавьте этот email в доступы таблицы как редактора.

## Где взять ID таблицы

Из ссылки вида:

```text
https://docs.google.com/spreadsheets/d/XXXXXXXXXXXX/edit
```

ID таблицы:

```text
XXXXXXXXXXXX
```

По умолчанию приложение использует ID таблицы `19hfmYJtv9FCzS6vSJ6OcwTRK8TuSNfgjCjEdF8JRs1Y`. При необходимости его можно изменить в `config.json` в корне проекта.

Пример `config.json`:

```json
{
  "google_sheet_id": "19hfmYJtv9FCzS6vSJ6OcwTRK8TuSNfgjCjEdF8JRs1Y"
}
```

## Формат таблицы

Обязательные колонки в Excel и Google Таблице:

- `Код`
- `Поставщик`
- `Менеджер`

## Логика работы

1. Приложение читает Excel-файл.
2. Подключается к Google Sheets API через `Streamlit secrets`.
3. Открывает Google Таблицу по ID из `config.json` или встроенного значения по умолчанию.
4. Если включен режим `Без записи`, приложение выполняет все проверки и пишет лог, но не изменяет данные в таблице.
5. Для каждой строки из Excel:
   - берет значение из колонки `Код`;
   - ищет строку в Google Таблице по колонке `Код`;
   - если код найден — обновляет `Поставщик` и `Менеджер` или пишет в лог, что обновление было бы выполнено;
   - если код не найден — добавляет новую строку и окрашивает диапазон `A:G` в светло-серый цвет или только пишет об этом в лог в режиме без записи.
6. Показывает лог выполнения в отдельном компактном окне с собственной прокруткой.

## Возможные ошибки

- Нет доступа к таблице → проверьте, что сервисному аккаунту выдан доступ редактора.
- Неверно заполнен `.streamlit/secrets.toml` → проверьте, что все поля из JSON-ключа перенесены без изменений.
- Неверный ID таблицы → убедитесь, что в `config.json` указан правильный идентификатор.
- Неправильный формат Excel → проверьте наличие столбцов `Код`, `Поставщик`, `Менеджер`.

## Интерфейс

В приложении доступны:

- заголовок;
- инструкция: `Загрузите xls таблицу со столбцами: Код, Поставщик, Менеджер`;
- встроенный ID Google Таблицы по умолчанию без отдельного поля ввода;
- загрузка файла;
- галочка `Без записи`;
- кнопка `Запустить обработку`;
- компактное окно лога с собственной прокруткой для сообщений об обновлении, добавлении и ошибках.
