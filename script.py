import csv
from pathlib import Path
from dataclasses import dataclass, field

import pandas as pd


@dataclass
class AppsMeta:
    month: int
    year: int = 2025
    root_path: Path = Path('/Users/worker/Downloads/PayPal')
    apps: tuple = ('app1', 'app2', 'app3', 'app4', 'app5')
    header: list | None = None
    all_files: list = field(default_factory=list)

    def get_folder_path(self, app: str) -> Path:
        return self.root_path / f'PayPal {app}' / 'Monthly reports' / f'{self.month:02d}.{self.year}'

    def get_output_path(self) -> Path:
        return self.root_path / 'Sum monthly reports' / f'Sum PayPal {self.month:02d}.{self.year}.xlsx'


def save_report_to_excel(output_excel: Path) -> None:
    print('Сохраняю файл, ожидайте,скоро всё будет готово!')
    try:
        pd.DataFrame(meta.all_files, columns=meta.header).to_excel(output_excel, index=False)
        print(f'\nФайл Sum PayPal {meta.month:02d}.2025.xlsx успешно сохранен!')
    except Exception as e:
        print(f'Ошибка при сохранении: {e}')
    meta.all_files = []


def read_app_folder(app: str) -> list:
    rows = []
    print(f'Читаю папку: {app}')
    for csv_file in meta.get_folder_path(app).glob('*.[cC][sS][vV]'):
        currency_app = csv_file.stem.split('_')[-1]

        with open(csv_file, 'r', encoding='utf-8') as file:
            reader = csv.reader(file)
            header_found = False
            for row in reader:
                if row and row[0].strip().lower() == 'date':
                    if meta.header is None:
                        meta.header = row + ['Currency', 'App', 'Month']
                    header_found = True
                    break

            if not header_found:
                print(f' Заголовок в файле: {csv_file.name} не найден, пропускаю')
            else:
                for row in reader:
                    if not row or all(cell.strip() == '' for cell in row[:5]):
                        continue
                    row.extend([currency_app, app, meta.month])
                    rows.append(row)
    return rows


def read_all_apps() -> None:
    for app_name in meta.apps:
        all_rows = read_app_folder(app_name)
        if all_rows:
            meta.all_files.extend(all_rows)
        else:
            print(f'В папке PayPal {app_name} нет данных для сохранения за выбранный месяц')


month = int(input('Введите месяц по которому нужно создать отчет (цифру 1-12)\n'))
print(f'\nНачинаю обработку...\n')

meta = AppsMeta(month)
read_all_apps()
save_report_to_excel(meta.get_output_path())