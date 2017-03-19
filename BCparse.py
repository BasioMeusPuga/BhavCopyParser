#!/usr/bin/python

import os
import csv
import zipfile
import datetime
import argparse
import xlsxwriter
import collections
import urllib.request


my_cwd = os.getcwd()


class Colors:
	CYAN = '\033[96m'
	GREEN = '\033[92m'
	RED = '\033[91m'
	WHITE = '\033[97m'
	YELLOW = '\033[93m'
	ENDC = '\033[0m'


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

	def do_the_dew(self, custom_link):
		if custom_link is None:
			# Filename is generated from the date
			file_date = str(self.bcdate.day).zfill(2) + str(self.bcdate.month).zfill(2) + str(self.bcdate.year)[2:]
			download_file = 'EQ_ISINCODE_{}.zip'.format(file_date)
			download_link = 'http://www.bseindia.com/download/BhavCopy/Equity/{}'.format(download_file)
		else:
			# Date is taken from the filename
			download_file = custom_link.rsplit('/', 1)[1]
			file_date = os.path.splitext(download_file)[0][-6:]
			download_link = custom_link

		output_file = my_cwd + '/' + download_file

		print('Attempting to download: ' + Colors.CYAN + download_link + Colors.ENDC)
		try:
			urllib.request.urlretrieve(download_link, output_file)
		except urllib.error.HTTPError:
			""" No special provisions have been made for a specific http error code:
			We're working with the assumption that it's 404 """
			print(Colors.RED + 'Bhavcopy not found' + Colors.ENDC)
			exit(1)
		print(Colors.GREEN + 'Download successful' + Colors.ENDC)

		""" Successfully downloaded files are shifted to a new directory named according to
		the date of the download """
		self.final_dir = my_cwd + '/' + file_date
		os.mkdir(self.final_dir)

		self.final_file = self.final_dir + '/' + download_file
		os.rename(output_file, self.final_file)

		return True

	def extract_csv(self):
		bc_zip = zipfile.ZipFile(self.final_file, 'r')

		# The name of the csv file is taken from the 0 index of the zip file
		self.csv_path = self.final_dir + '/' + bc_zip.filelist[0].filename

		# It's then extracted to the newly created directory
		bc_zip.extractall(path=self.final_dir)
		bc_zip.close()


class ParseBhavCopy:
	def __init__(self, csv_path):
		# Open csv derived from the DownloadBhavCopy class
		# Take relevant path info from the filename
		self.csv_path = csv_path
		self.csv_dir = os.path.dirname(csv_path)
		self.csv_name = os.path.splitext(os.path.basename(csv_path))[0]

	def parse_csv(self):
		self.scrip_data = []
		print('Attempting to parse: ' + Colors.CYAN + self.csv_path + Colors.ENDC)
		with open(self.csv_path) as csvfile:
			bc_csv = csv.reader(csvfile, delimiter=',')
			for row in bc_csv:
				self.scrip_data.append({
					'scrip_name': row[1].strip(),
					'scrip_open': row[4],
					'scrip_high': row[5],
					'scrip_low': row[6],
					'scrip_close': row[7]})

		self.write_xlsx()

	def write_xlsx(self):
		workbook = xlsxwriter.Workbook(self.csv_dir + '/' + self.csv_name[-6:] + '.xlsx')

		# Start with writing the entire Bhavcopy to the first worksheet
		worksheet = workbook.add_worksheet('ALL SCRIPS')

		worksheet.set_column('A:E', 20)
		bold = workbook.add_format({'bold': True})

		worksheet.write('A1', 'SCRIP NAME', bold)
		worksheet.write('B1', 'OPEN', bold)
		worksheet.write('C1', 'HIGH', bold)
		worksheet.write('D1', 'LOW', bold)
		worksheet.write('E1', 'CLOSE', bold)

		line_number = 2
		for key in self.scrip_data[1:]:
			ln_str = str(line_number)

			worksheet.write('A' + ln_str, key['scrip_name'])
			worksheet.write('B' + ln_str, float(key['scrip_open']))
			worksheet.write('C' + ln_str, float(key['scrip_high']))
			worksheet.write('D' + ln_str, float(key['scrip_low']))
			worksheet.write('E' + ln_str, float(key['scrip_close']))

			line_number += 1

		""" Write individual worksheets for each client
		This depends on the existence of Clients.txt in the same dir as the script
		In case the file is not found, create it, give a message and exit gracefully """
		client_file_path = my_cwd + '/' + 'Clients.txt'
		if not os.path.exists(client_file_path):
			print(Colors.RED + 'Clients.txt not found. A new file has been created with instructions.' + Colors.ENDC)
			with open(client_file_path, 'w') as my_clients:
				my_clients.write('#NAMEOFCLIENT:SCRIP1;SCRIP2;SCRIP3...' + '\n')
			exit(1)

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
					bold = workbook.add_format({'bold': True})

					worksheet.write('A1', 'SCRIP NAME', bold)
					worksheet.write('B1', 'OPEN', bold)
					worksheet.write('C1', 'HIGH', bold)
					worksheet.write('D1', 'LOW', bold)
					worksheet.write('E1', 'CLOSE', bold)

					line_number = 2
					for key in self.scrip_data[1:]:
						ln_str = str(line_number)

						""" Write to the worksheet in case the name of the scrips match
						This could also use fuzzy matching in case Clients.txt can't be trusted """
						if key['scrip_name'] in client_scrips:
							worksheet.write('A' + ln_str, key['scrip_name'])
							worksheet.write('B' + ln_str, float(key['scrip_open']))
							worksheet.write('C' + ln_str, float(key['scrip_high']))
							worksheet.write('D' + ln_str, float(key['scrip_low']))
							worksheet.write('E' + ln_str, float(key['scrip_close']))

							line_number += 1	


def main():
	parser = argparse.ArgumentParser(description='Download (today\'s) Bhavcopy.')
	parser.add_argument('--bcdate', type=str, nargs=1, help='Specify date for Bhavcopy download', metavar='ddmmyy or dd/mm/yy')
	parser.add_argument('--custom', type=str, nargs=1, help='Link to Bhavcopy', metavar='<url>')
	args = parser.parse_args()

	if args.bcdate:
		get_bhavcopy = DownloadBhavCopy(args.bcdate[0])
	else:
		get_bhavcopy = DownloadBhavCopy(None)

	if args.custom:
		download_status = get_bhavcopy.do_the_dew(args.custom[0])
	else:
		download_status = get_bhavcopy.do_the_dew(None)

	if download_status is True:
		get_bhavcopy.extract_csv()

		bhavcopy_shenanigans = ParseBhavCopy(get_bhavcopy.csv_path)
		bhavcopy_shenanigans.parse_csv()	
		print(Colors.GREEN + 'Parsing complete.' + Colors.ENDC)


if __name__ == '__main__':
	main()
