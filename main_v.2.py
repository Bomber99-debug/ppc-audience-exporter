# -*- coding: utf-8 -*-
import csv
import hashlib
import os
import paramiko
from openpyxl import load_workbook
from transliterate import translit

from config import config as sftp_config


def file_read( file_name, sheet_index, column_numbers ):
	"""Читає файл Excel і повертає вибрані стовпці."""

	def column_selection( rows, column_numbers ):
		return [ [ row[ i - 1 ] for i in column_numbers ] for row in rows if
				 any( row[ i - 1 ] for i in column_numbers ) ]

	wb = load_workbook( filename=file_name )
	sheet = wb[ wb.sheetnames[ sheet_index ] ]
	read_lines = sheet.iter_rows( min_row=1, values_only=True )

	data = column_selection( read_lines, column_numbers )
	if not data:
		print( "No data found in file_read" )
	return data


def rewrite_headers( data ):
	"""Змінює заголовки на потрібні."""
	if data and len( data ) > 0:
		data[ 0 ] = [ 'Email', 'Phone', 'First Name', 'Last Name', 'Zip', 'Country' ]
	else:
		print( "Data is empty in rewrite_headers" )


def split_full_name( row, name_index ):
	"""Розділяє 'ФИО' на 'First Name' та 'Last Name'."""
	if row[ name_index ]:
		full_name = row[ name_index ].split()
		row[ name_index ] = full_name[ 0 ] if len( full_name ) > 0 else 'Unknown'
		row.insert( name_index + 1, full_name[ 1 ] if len( full_name ) > 1 else 'Unknown' )
	else:
		row[ name_index ] = 'Unknown'
		row.insert( name_index + 1, 'Unknown' )

	return row


def hash_object( object ):
	if object:
		hash_object = hashlib.sha256( object.encode() )
		object_hash = hash_object.hexdigest()

		return object_hash

	else:
		return ''


def add_country_and_phone( row ):
	"""Додає країну і телефонний номер у відповідні позиції."""
	Email = f'{row[ 0 ]}'.replace( ' ', '' ) if row[ 0 ] else ''
	First_Name = row[ 1 ]
	Last_Name = row[ 2 ]
	Country = 'ua'  # Додаємо країну
	Zip = row[ 4 ] if row[ 4 ] else '1001'  # Якщо zip коду немає, додає код Києва
	Phone = f'380{row[ 3 ]}'.replace( ' ', '' ) if row[ 3 ] else ''
	category = row[ 5 ] if len( row ) > 5 else ''  # Переконайтесь, що row має 6 елементів

	return [
			hash_object( Email ), hash_object( Phone ), hash_object( First_Name ), hash_object( Last_Name ), Zip,
			Country, category,
			]  # return [Email, Phone, First_Name, Last_Name, Zip, Country, category]


def category_file_read( file_name ):
	"""Читає файл категорій та повертає дані у словнику."""
	wb = load_workbook( filename=file_name )
	sheet = wb.active
	number_columns = sheet.max_column
	data = { }

	for i in range( 1, number_columns + 1 ):
		column_data = [ x[ 0 ] for x in sheet.iter_rows( min_col=i, max_col=i, values_only=True ) if
						x[ 0 ] is not None ]
		if column_data:
			header = column_data[ 0 ]
			if header:
				data[ header ] = column_data[ 1: ]
			else:
				data[ f'Unnamed {i}' ] = column_data[ 1: ]

	if not data:
		print( "No data found in category_file_read" )
	return data


def write_to_csv( file_name, data ):
	"""Записує дані у CSV файл."""
	with open( file_name, mode='w', newline='', encoding='utf-8' ) as file:
		writer = csv.writer( file, delimiter=',' )
		writer.writerows( data )
	print( f"Data written to {file_name}" )


def write_to_category_csv( category_file, recording_path, data ):
	# Перевіряємо, чи існує каталог, і створюємо його, якщо ні
	if not os.path.exists( recording_path ):
		os.makedirs( recording_path )

	"""Записує дані у категорійні CSV файли."""
	for category_name, sub_category in category_file.items():
		file_name = os.path.join( translit( category_name, 'uk', reversed=True ) + '.csv' ).replace( "'", "" ).replace(
			",", '', ).replace( ' ', '_' ).lower()
		# file_name = file_name.replace("'","").lower()
		lib_data = [ data[ 0 ] ]  # Додаємо заголовок
		for row in data[ 1: ]:
			if row[ -1 ] in sub_category:
				lib_data.append( row[ :-1 ] )

		file_path = os.path.join( recording_path, category_name + '.csv' )  # Додаємо розширення

		write_to_csv( file_path, lib_data )

		# Параметри SFTP сервера
		sftp_server = sftp_config.sftp_host
		username = sftp_config.sftp_username
		password = sftp_config.sftp_password

		# Підключення до SFTP сервера
		transport = paramiko.Transport( (sftp_server, 22) )
		transport.connect( username=username, password=password )

		sftp = paramiko.SFTPClient.from_transport( transport )

		# Вказуємо шлях до локального та віддаленого файлу
		remote_file = f'/{sftp_config.path_directory_sftp}/{file_name}'

		# Завантажуємо файл на сервер
		sftp.put( file_path, remote_file )

		print( f'Файл {file_path} успішно завантажений на сервер як {remote_file}.' )

		# Закриття з'єднання з сервером
		sftp.close()  # Закриваємо SFTP-з'єднання
		transport.close()  # Закриваємо транспортне з'єднання


def write_to_all_csv( recording_path, data ):
	# Перевіряємо, чи існує каталог, і створюємо його, якщо ні
	if not os.path.exists( recording_path ):
		os.makedirs( recording_path )

	lib_data = [ data[ 0 ] ]
	for row in data[ 1: ]:
		lib_data.append( row[ :-1 ] )

	file_path = os.path.join( recording_path, 'all' + '.csv' )
	write_to_csv( file_path, lib_data )

	# Параметри SFTP сервера
	sftp_server = sftp_config.sftp_host
	username = sftp_config.sftp_username
	password = sftp_config.sftp_password

	# Підключення до SFTP сервера
	transport = paramiko.Transport( (sftp_server, 22) )
	transport.connect( username=username, password=password )

	sftp = paramiko.SFTPClient.from_transport( transport )

	# Вказуємо шлях до локального та віддаленого файлу
	remote_file = f'/{sftp_config.path_directory_sftp}/all.csv'

	# Завантажуємо файл на сервер
	sftp.put( file_path, remote_file )

	print( f'Файл {file_path} успішно завантажений на сервер як {remote_file}.' )

	# Закриття з'єднання з сервером
	sftp.close()  # Закриваємо SFTP-з'єднання
	transport.close()  # Закриваємо транспортне з'єднання


def main():
	file_name_1 = ''  # Шлях до файлу даних у форматі *.xlsx
	category_file_name = 'category.xlsx'
	recording_path = "category/"
	sheet_index = 0
	column_numbers = [ 4, 5, 6, 7, 9 ]  # E mail, ФИО, № телефона, Индекс, Категория

	data = file_read( file_name_1, sheet_index, column_numbers )

	if data:
		data = [ add_country_and_phone( split_full_name( row, 1 ) ) for row in data[ 1: ] ]
		rewrite_headers( data )
	else:
		print( "Дані порожні після file_read або rewrite_headers" )

	write_to_all_csv( recording_path, data )

	category_file = category_file_read( category_file_name )

	if category_file:
		write_to_category_csv( category_file, recording_path, data )
	else:
		print( "Дані категорії не знайдено" )


if __name__ == '__main__':
	main()
	print( 'Усі файли записано' )
nput()
