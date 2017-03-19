#!/usr/bin/python3

import os
import csv
import zipfile
import datetime
import argparse
import xlsxwriter
import requests


class Colors:
	CYAN = '\033[96m'
	GREEN = '\033[92m'
	RED = '\033[91m'
	WHITE = '\033[97m'
	YELLOW = '\033[93m'
	ENDC = '\033[0m'


my_cwd = os.getcwd()
client_file_path = my_cwd + '/' + 'Clients.txt'
if not os.path.exists(client_file_path):
	print(Colors.RED + 'Clients.txt not found. A new file has been created with instructions.' + Colors.ENDC)
	with open(client_file_path, 'w') as my_clients:
		my_clients.write('#NAMEOFCLIENT:SCRIP1;SCRIP2;SCRIP3...' + '\n')


class DownloadBhavCopy:
	def __init__(self, bcdate):
		if bcdate is None:
			self.bcdate = datetime.date.today()
		else:
			# Accepted dates are either ddmmyy OR dd/mm/yy
			# A leading zero is expected for a single digit number
			try:
				self.bcdate = datetime.datetime.strptime(bcdate, '%d%m%y')
			except ValueError:
				try:
					self.bcdate = datetime.datetime.strptime(bcdate, '%d/%m/%y')
				except ValueError:
					print(Colors.RED + 'Date is improper OR not formatted: ' + Colors.GREEN + 'ddmmyy | dd/mm/yy' + Colors.ENDC)
					exit(1)

		self.do_the_dew()

	def do_the_dew(self):
		# Generate file name for the BSE bhavcopy
		file_date = str(self.bcdate.day).zfill(2) + str(self.bcdate.month).zfill(2) + str(self.bcdate.year)[2:]
		download_file_bse = 'EQ_ISINCODE_{}.zip'.format(file_date)
		download_link_bse = 'http://www.bseindia.com/download/BhavCopy/Equity/{}'.format(download_file_bse)

		# Generate file name for the NSE bhavcopy
		download_file_nse = 'cm{0}{1}{2}bhav.csv.zip'.format(
			str(self.bcdate.day).zfill(2),
			self.bcdate.strftime('%B')[:3].upper(),
			str(self.bcdate.year))
		download_link_nse = 'https://www.nseindia.com/content/historical/EQUITIES/{2}/{1}/cm{0}{1}{2}bhav.csv.zip'.format(
			str(self.bcdate.day).zfill(2),
			self.bcdate.strftime('%B')[:3].upper(),
			str(self.bcdate.year))

		# Download files to a new directory named according to the date
		self.final_dir = my_cwd + '/' + str(self.bcdate.day).zfill(2) + '-' + str(self.bcdate.month).zfill(2) + '-' + str(self.bcdate.year)
		try:
			os.mkdir(self.final_dir)
		except:
			pass

		# This will be iterated over repeatedly
		self.download_files = [
		{'file': self.final_dir + '/' + download_file_nse, 'link': download_link_nse, 'type': 'nse'},
		{'file': self.final_dir + '/' + download_file_bse, 'link': download_link_bse, 'type': 'bse'}]

		# Proceed to download both bhavcopies
		for i in self.download_files:
			print('Attempting to download: ' + Colors.CYAN + i['file'] + Colors.ENDC)

			""" I shifted to the requests library because urllib.requests can't seem to
			do headers easily. Especially for url retrievals """
			r = requests.get(i['link'], stream=True, headers={'User-agent': 'Mozilla/5.0'})
			if r.status_code == 200:
				with open(i['file'], 'wb') as f:
					for chunk in r.iter_content(chunk_size=1024):
						if chunk:
							f.write(chunk)
				print(Colors.GREEN + 'Download successful' + Colors.ENDC)
			elif r.status_code == 404:
				print(Colors.RED + 'Bhavcopy not found' + Colors.ENDC)
			else:
				print(Colors.RED + 'HTML error code: ' + str(r.status_code) + Colors.ENDC)

		self.extract_csv()

	def extract_csv(self):
		for i in self.download_files:
			if os.path.exists(i['file']):
				bc_zip = zipfile.ZipFile(i['file'], 'r')

				# The name of the csv file is taken from the 0 index of the zip file
				if i['type'] == 'bse':
					self.csv_path_bse = self.final_dir + '/' + bc_zip.filelist[0].filename
				if i['type'] == 'nse':
					self.csv_path_nse = self.final_dir + '/' + bc_zip.filelist[0].filename

				# It's then extracted to the newly created directory
				bc_zip.extractall(path=self.final_dir)
				bc_zip.close()

				# Delete the zip file after completion
				os.remove(i['file'])
			else:
				if i['type'] == 'bse':
					self.csv_path_bse = None
				if i['type'] == 'nse':
					self.csv_path_nse = None


class ParseBhavCopy:
	def __init__(self, csv_path, stock_exchange, file_date):
		# Open csv derived from the DownloadBhavCopy class
		# Take relevant path info from the filename
		self.csv_path = csv_path
		self.csv_dir = os.path.dirname(csv_path)
		self.csv_name = os.path.splitext(os.path.basename(csv_path))[0]
		self.stock_exchange = stock_exchange
		self.file_date = file_date

		self.parse_csv()

	def parse_csv(self):
		self.scrip_data = []
		print('Attempting to parse: ' + Colors.CYAN + self.csv_path + Colors.ENDC)
		with open(self.csv_path) as csvfile:
			bc_csv = csv.reader(csvfile, delimiter=',')
			for row in bc_csv:
				if self.stock_exchange == 'bse':
					self.scrip_data.append({
						'scrip_name': row[1].strip(),
						'scrip_open': row[4],
						'scrip_high': row[5],
						'scrip_low': row[6],
						'scrip_close': row[7]})
				elif self.stock_exchange == 'nse':
					self.scrip_data.append({
						'scrip_name': row[0].strip(),
						'scrip_open': row[2],
						'scrip_high': row[3],
						'scrip_low': row[4],
						'scrip_close': row[5]})

		self.write_xlsx()

	def write_xlsx(self):
		if self.stock_exchange == 'bse':
			workbook = xlsxwriter.Workbook(self.csv_dir + '/' + '(BSE) ' + self.file_date + '.xlsx')
		elif self.stock_exchange == 'nse':
			workbook = xlsxwriter.Workbook(self.csv_dir + '/' + '(NSE) ' + self.file_date + '.xlsx')
		bold = workbook.add_format({'bold': True})

		def write_to_worksheet(this_worksheet, line_number, col_a, col_b, col_c, col_d, col_e, set_bold):
			if set_bold is True:
				worksheet.write('A' + str(line_number), col_a, bold)
				worksheet.write('B' + str(line_number), col_b, bold)
				worksheet.write('C' + str(line_number), col_c, bold)
				worksheet.write('D' + str(line_number), col_d, bold)
				worksheet.write('E' + str(line_number), col_e, bold)
			else:
				worksheet.write('A' + str(line_number), col_a)
				worksheet.write('B' + str(line_number), col_b)
				worksheet.write('C' + str(line_number), col_c)
				worksheet.write('D' + str(line_number), col_d)
				worksheet.write('E' + str(line_number), col_e)

		# Start with writing the entire Bhavcopy to the first worksheet
		worksheet = workbook.add_worksheet('ALL SCRIPS')
		worksheet.set_column('A:E', 20)

		write_to_worksheet(worksheet, 1, 'SCRIP NAME', 'OPEN', 'HIGH', 'LOW', 'CLOSE', True)

		line_number = 2
		for key in self.scrip_data[1:]:
			write_to_worksheet(
				worksheet,
				line_number,
				key['scrip_name'],
				float(key['scrip_open']),
				float(key['scrip_high']),
				float(key['scrip_low']),
				float(key['scrip_close']),
				False)
			line_number += 1

		""" Write individual worksheets for each client
		This depends on the existence of Clients.txt in the same dir as the script """
		with open(client_file_path, 'r') as client_file:
			my_clients = client_file.readlines()
			# Iterate over the list of clients mentioned in Clients.txt
			for i in my_clients:
				client_info = i.replace('\n', '')
				client_name = client_info.split(':')[0]
				client_scrips = client_info.split(':')[1].split(';')

				# Create a new worksheet for each client
				if client_name[0] != '#':
					worksheet = workbook.add_worksheet(client_name)
					worksheet.set_column('A:E', 20)

					write_to_worksheet(worksheet, 1, 'SCRIP NAME', 'OPEN', 'HIGH', 'LOW', 'CLOSE', True)

					line_number = 2
					for key in self.scrip_data[1:]:
						""" Write to the worksheet in case the name of the scrips match
						This could also use fuzzy matching in case Clients.txt can't be trusted """
						if key['scrip_name'] in client_scrips:
							write_to_worksheet(
								worksheet,
								line_number,
								key['scrip_name'],
								float(key['scrip_open']),
								float(key['scrip_high']),
								float(key['scrip_low']),
								float(key['scrip_close']),
								False)

							line_number += 1
		print(Colors.GREEN + 'Parsing complete.' + Colors.ENDC)


def main():
	parser = argparse.ArgumentParser(description='Download (today\'s) Bhavcopy.')
	parser.add_argument('--bcdate', type=str, nargs=1, help='Specify date for Bhavcopy download', metavar='ddmmyy or dd/mm/yy')
	args = parser.parse_args()

	if args.bcdate:
		get_bhavcopy = DownloadBhavCopy(args.bcdate[0])
	else:
		get_bhavcopy = DownloadBhavCopy(None)

	file_date = str(get_bhavcopy.bcdate.day).zfill(2) + '-' + str(get_bhavcopy.bcdate.month).zfill(2) + '-' + str(get_bhavcopy.bcdate.year)

	if get_bhavcopy.csv_path_bse:
		ParseBhavCopy(get_bhavcopy.csv_path_bse, 'bse', file_date)
	if get_bhavcopy.csv_path_nse:
		ParseBhavCopy(get_bhavcopy.csv_path_nse, 'nse', file_date)


if __name__ == '__main__':
	main()
