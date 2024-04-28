from datetime import datetime

from aiogoogle import Aiogoogle

from app.core.config import settings

FORMAT = "%Y/%m/%d %H:%M:%S"
ROW_COUNT = 100
COLUMN_COUNT = 11


async def spreadsheets_create(wrapper_services: Aiogoogle) -> str:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    spreadsheet_body = {
        'properties': {
            'title': f'Отчет от {now_date_time}',
            'locale': 'ru_RU',
        },
        'sheets': [
            {
                'properties': {
                    'sheetType': 'GRID',
                    'sheetId': 0,
                    'title': 'Лист1',
                    'gridProperties': {
                        'rowCount': ROW_COUNT,
                        'columnCount': COLUMN_COUNT,
                    },
                }
            }
        ],
    }
    response = await wrapper_services.as_service_account(
        service.spreadsheets.create(json=spreadsheet_body)
    )
    spreadsheetid = response['spreadsheetId']
    return spreadsheetid


async def set_user_permissions(
    spreadsheetid: str, wrapper_services: Aiogoogle
) -> None:
    permissions_body = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': settings.email,
    }
    service = await wrapper_services.discover('drive', 'v3')
    await wrapper_services.as_service_account(
        service.permissions.create(
            fileId=spreadsheetid, json=permissions_body, fields="id"
        )
    )


async def spreadsheets_update_value(
    spreadsheetid: str, projects: list, wrapper_services: Aiogoogle
) -> None:
    now_date_time = datetime.now().strftime(FORMAT)
    service = await wrapper_services.discover('sheets', 'v4')
    table_values = [
        ['Отчет на', now_date_time],
        ['Топ проектов по скорости закрытия'],
        ['Название', 'Время сбора', 'Описание'],
    ]
    for pr in projects:
        new_row = [
            str(pr['name']),
            str(pr['time_of_collecting']),
            str(pr['description']),
        ]
        table_values.append(new_row)

    update_body = {'majorDimension': 'ROWS', 'values': table_values}

    if len(table_values) > ROW_COUNT:
        raise ValueError(
            'Количество проектов превышает допустимый размер таблицы'
        )
    await wrapper_services.as_service_account(
        service.spreadsheets.values.update(
            spreadsheetId=spreadsheetid,
            range='1:30',
            valueInputOption='USER_ENTERED',
            json=update_body,
        )
    )
    print(f'https://docs.google.com/spreadsheets/d/{spreadsheetid}')
