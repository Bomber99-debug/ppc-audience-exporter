from __future__ import annotations

import csv
import hashlib
import paramiko
from collections.abc import Iterator
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from openpyxl import load_workbook
from pathlib import Path
from transliterate import translit
from typing import TypeAlias

from config import config

# Значення, яке може лежати в Excel-клітинці для цієї задачі
CellValue: TypeAlias = str | int | float | bool | None

# Один рядок даних з Excel
RowData: TypeAlias = list[ CellValue ]

# Набір рядків
TableData: TypeAlias = list[ RowData ]

# Дані категорій:
# ключ = назва великої категорії
# значення = список підкатегорій, що належать до неї
CategoryMap: TypeAlias = dict[ str, list[ str ] ]


def iter_excel_files( source_dir: Path ) -> Iterator[ Path ]:
	"""
	Повертає всі .xlsx файли з вказаної директорії.
	"""
	for file_path in source_dir.iterdir():
		if file_path.is_file() and file_path.suffix.lower() == ".xlsx":
			yield file_path


def read_required_columns_from_excel( file_path: Path ) -> TableData:
	"""
	Читає перший лист Excel-файлу, знаходить рядок заголовків
	і повертає лише потрібні колонки.

	Очікувані заголовки:
	- E mail
	- ФИО
	- № телефона
	- Индекс
	- Категория
	"""
	workbook = load_workbook( filename=file_path, read_only=True, data_only=True )
	worksheet = workbook[ workbook.sheetnames[ 0 ] ]

	required_headers: list[ str ] = [
			"E mail", "ФИО", "№ телефона", "Индекс", "Категория",
			]

	header_row_number: int | None = None
	required_column_indexes: list[ int ] = [ ]

	# Шукаємо рядок, у якому знаходяться потрібні заголовки.
	# Заголовок може бути не в першому рядку, тому проходимо всі рядки зверху вниз.
	for row_number, row in enumerate( worksheet.iter_rows(), start=1 ):
		row_values: list[ str ] = [ str( cell.value ).strip() if cell.value is not None else "" for cell in row ]

		# Перша груба перевірка: якщо в рядку є хоча б перший ключовий заголовок,
		# тоді вже перевіряємо повний набір.
		if required_headers[ 0 ] in row_values:
			missing_headers: list[ str ] = [ header for header in required_headers if header not in row_values ]

			if missing_headers:
				workbook.close()
				raise ValueError(
					f"У файлі {file_path} відсутні колонки: {', '.join( missing_headers )}", )

			# Зберігаємо індекси лише потрібних колонок.
			required_column_indexes = [ row_values.index( header ) for header in required_headers ]
			header_row_number = row_number
			break

	if header_row_number is None:
		workbook.close()
		raise ValueError( f"У файлі {file_path} не знайдено рядок заголовків" )

	extracted_rows: TableData = [ ]

	# Читаємо всі рядки ПІСЛЯ заголовка.
	for row in worksheet.iter_rows( min_row=header_row_number + 1, values_only=True ):
		selected_row: RowData = [ row[ column_index ] for column_index in required_column_indexes ]
		extracted_rows.append( selected_row )

	workbook.close()
	return extracted_rows


def split_full_name( row: RowData, full_name_index: int ) -> RowData:
	"""
	Розбиває поле 'ФИО' на ім'я та прізвище.

	Було:
	[email, "Іван Петренко", phone, zip, category]

	Стає:
	[email, "Іван", "Петренко", phone, zip, category]

	Якщо ПІБ порожнє або містить лише одне слово,
	підставляється 'Unknown' для відсутньої частини.
	"""
	raw_full_name = row[ full_name_index ]

	if raw_full_name is None:
		first_name = "Unknown"
		last_name = "Unknown"
	else:
		full_name_parts = str( raw_full_name ).split()
		first_name = full_name_parts[ 0 ] if len( full_name_parts ) > 0 else "Unknown"
		last_name = full_name_parts[ 1 ] if len( full_name_parts ) > 1 else "Unknown"

	row[ full_name_index ] = first_name
	row.insert( full_name_index + 1, last_name )

	return row


def sha256_hash( value: str | None ) -> str:
	"""
	Повертає SHA-256 хеш рядка.
	Якщо значення порожнє, повертає порожній рядок.
	"""
	if not value:
		return ""

	return hashlib.sha256( value.encode( "utf-8" ) ).hexdigest()


def build_export_row( source_row: RowData ) -> RowData:
	"""
	Перетворює рядок з Excel у фінальний формат для CSV.

	Очікує вхідний формат:
	[email, first_name, last_name, phone, zip, category]

	Повертає:
	[hashed_email, hashed_phone, hashed_first_name, hashed_last_name, zip, country, category]
	"""
	email = str( source_row[ 0 ] ).replace( " ", "" ) if source_row[ 0 ] else ""
	first_name = str( source_row[ 1 ] ) if source_row[ 1 ] else ""
	last_name = str( source_row[ 2 ] ) if source_row[ 2 ] else ""
	zip_code = str( source_row[ 4 ] ) if source_row[ 4 ] else "1001"
	phone = f"380{source_row[ 3 ]}".replace( " ", "" ) if source_row[ 3 ] else ""
	category = str( source_row[ 5 ] ) if len( source_row ) > 5 and source_row[ 5 ] else ""
	country = "ua"

	return [
			sha256_hash( email ), sha256_hash( phone ), sha256_hash( first_name ), sha256_hash( last_name ), zip_code,
			country, category,
			]


def write_rows_to_csv( file_path: Path, rows: TableData ) -> None:
	"""
	Записує таблицю у CSV файл.
	"""
	with file_path.open( mode="w", newline="", encoding="utf-8" ) as csv_file:
		writer = csv.writer( csv_file, delimiter="," )
		writer.writerows( rows )

	print( f"Data written to {file_path}" )


def ensure_directory_exists( directory_path: Path ) -> None:
	"""
	Створює директорію, якщо вона ще не існує.
	"""
	directory_path.mkdir( parents=True, exist_ok=True )


def upload_file_to_sftp( local_file_path: Path, remote_file_path: str ) -> None:
	"""
	Завантажує локальний файл на SFTP сервер.

	Параметри підключення беруться зі змінних середовища:
	- SFTP_HOST
	- SFTP_PORT
	- SFTP_USERNAME
	- SFTP_PASSWORD
	"""
	sftp_host = config.sftp_host
	sftp_port = config.sftp_port
	sftp_username = config.sftp_username
	sftp_password = config.sftp_password

	if not sftp_host or not sftp_username or not sftp_password:
		raise ValueError(
			"Не задані SFTP змінні середовища: SFTP_HOST, SFTP_USERNAME, SFTP_PASSWORD", )

	transport = paramiko.Transport( (sftp_host, sftp_port) )
	transport.connect( username=sftp_username, password=sftp_password )

	sftp_client = paramiko.SFTPClient.from_transport( transport )
	try:
		sftp_client.put( str( local_file_path ), remote_file_path )
		print( f"Файл {local_file_path} успішно завантажено на сервер як {remote_file_path}" )
	finally:
		sftp_client.close()
		transport.close()


def write_all_contacts_csv( output_dir: Path, rows: TableData ) -> None:
	"""
	Створює загальний CSV файл all.csv без останньої колонки category
	і завантажує його на SFTP.
	"""
	ensure_directory_exists( output_dir )

	if not rows:
		return

	csv_rows: TableData = [ rows[ 0 ] ]

	# Для загального файлу прибираємо category з кожного рядка даних.
	for row in rows[ 1: ]:
		csv_rows.append( row[ :-1 ] )

	output_file_path = output_dir / "all.csv"
	write_rows_to_csv( output_file_path, csv_rows )

	upload_file_to_sftp( output_file_path, f"/{config.path_directory_sftp}/all.csv" )


def read_category_mapping( category_file_path: Path ) -> CategoryMap:
	"""
	Читає Excel-файл категорій.

	Кожна колонка:
	- перший рядок = назва категорії
	- решта рядків = підкатегорії цієї категорії
	"""
	workbook = load_workbook( filename=category_file_path )
	worksheet = workbook.active

	category_map: CategoryMap = { }

	for column_number in range( 1, worksheet.max_column + 1 ):
		column_values: list[ str ] = [ str( cell_value ) for (cell_value,) in worksheet.iter_rows(
			min_col=column_number, max_col=column_number, values_only=True, ) if cell_value is not None ]

		if not column_values:
			continue

		category_name = column_values[ 0 ].strip()

		if category_name:
			category_map[ category_name ] = column_values[ 1: ]
		else:
			category_map[ f"Unnamed_{column_number}" ] = column_values[ 1: ]

	workbook.close()

	if not category_map:
		print( "No data found in category mapping file" )

	return category_map


def build_safe_filename_from_category( category_name: str ) -> str:
	"""
	Будує безпечне ім'я файлу на основі назви категорії:
	- транслітерація
	- нижній регістр
	- пробіли -> _
	- видалення частини спецсимволів
	"""
	transliterated_name = translit( category_name, "uk", reversed=True )

	safe_filename = (transliterated_name.replace( "'", "" ).replace( ",", "" ).replace( " ", "_" ).lower())

	return f"{safe_filename}.csv"


def write_category_csv_files(
		category_mapping: CategoryMap, output_dir: Path, rows: TableData, ) -> None:
	"""
	Для кожної великої категорії створює окремий CSV файл
	і додає туди лише ті рядки, чия category входить у список підкатегорій.
	"""
	ensure_directory_exists( output_dir )

	if not rows:
		return

	header_row = rows[ 0 ]

	for category_name, allowed_subcategories in category_mapping.items():
		category_rows: TableData = [ header_row ]

		for row in rows[ 1: ]:
			if row and row[ -1 ] in allowed_subcategories:
				# У CSV для експорту category не записуємо.
				category_rows.append( row[ :-1 ] )

		local_file_path = output_dir / f"{category_name}.csv"
		write_rows_to_csv( local_file_path, category_rows )

		remote_filename = build_safe_filename_from_category( category_name )
		remote_file_path = f"/{config.path_directory_sftp}/{remote_filename}"

		upload_file_to_sftp( local_file_path, remote_file_path )


def collect_rows_from_excel_files( source_dir: Path ) -> TableData:
	"""
	Читає всі Excel-файли з директорії паралельно
	і об'єднує результати в один список рядків.
	"""
	collected_rows: TableData = [ ]

	tasks: list[ Future[ TableData ] ] = [ ]

	with ThreadPoolExecutor( max_workers=4 ) as executor:
		for excel_file_path in iter_excel_files( source_dir ):
			tasks.append( executor.submit( read_required_columns_from_excel, excel_file_path ) )

		# Беремо результати в міру завершення задач, а не по порядку запуску.
		for completed_task in as_completed( tasks ):
			file_rows = completed_task.result()
			collected_rows.extend( file_rows )

	return collected_rows


def transform_rows_for_export( raw_rows: TableData ) -> TableData:
	"""
	Перетворює сирі рядки з Excel у фінальний формат для експорту.
	"""
	transformed_rows: TableData = [
			[ "Email", "Phone", "First Name", "Last Name", "Zip", "Country", "Category" ],
            ]

	for raw_row in raw_rows:
		row_with_split_name = split_full_name( raw_row, full_name_index=1 )
		export_row = build_export_row( row_with_split_name )
		transformed_rows.append( export_row )

	return transformed_rows


def main() -> None:
	source_dir = Path( "./files" )
	output_dir = Path( "./category" )
	category_file_path = Path( "category.xlsx" )

	print( "[START] Починаємо обробку файлів" )

	print( f"[STEP] Читаємо Excel-файли з директорії: {source_dir}" )
	raw_rows = collect_rows_from_excel_files( source_dir )
	print( f"[INFO] Зчитано рядків: {len( raw_rows )}" )

	if not raw_rows:
		print( "[STOP] Не знайдено жодного рядка даних для обробки" )
		return

	print( "[STEP] Трансформуємо дані для експорту" )
	export_rows = transform_rows_for_export( raw_rows )
	print( f"[INFO] Підготовлено рядків для експорту: {len( export_rows )}" )

	print( "[STEP] Записуємо загальний CSV (all.csv) + відправка на SFTP" )
	write_all_contacts_csv( output_dir, export_rows )

	print( f"[STEP] Читаємо файл категорій: {category_file_path}" )
	category_mapping = read_category_mapping( category_file_path )
	print( f"[INFO] Знайдено категорій: {len( category_mapping )}" )

	if not category_mapping:
		print( "[STOP] Дані категорій не знайдено" )
		return

	print( "[STEP] Генеруємо CSV по категоріях + відправка на SFTP" )
	write_category_csv_files( category_mapping, output_dir, export_rows )

	print( "[DONE] Обробка завершена" )


if __name__ == "__main__":
	main(
)