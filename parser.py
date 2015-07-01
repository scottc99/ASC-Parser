import glob, os
import xlrd
import xlwt
from collections import OrderedDict
import simplejson as json 
import pprint
import tkFileDialog
from Tkinter import *	
import dicttoxml
import lxml.etree as etree


def convert_excel():
	
	for file in glob.glob("DataFiles/ASC/*.ASC"):
		print "Converting %s to excel..."%file
		with open(file, "r") as asc_file:
			file_name = file.split("/")[-1].split(".")[0]
			print file_name
			workbook = xlwt.Workbook()
			sheet = workbook.add_sheet('dataset')

			content = asc_file.read().split("\n")
			row_index = 0
			for row in content:
				columns = row.split()
				column_index = 0
				for column in columns:
					sheet.write(row_index, column_index, column.decode('ISO8859-1'))
					column_index += 1
				row_index += 1

			workbook.save('DataFiles/EXCEL/%s.xlsx'%file_name)
			print "Conversion for %s to excel is done."%file   

		sequence = file_name.split("_")[0]
		#file_path = os.path.abspath('/DataFiles/EXCEL/%s_1045-R1.xlsx'%sequence)
		file_pathXL = 'DataFiles/EXCEL/%s_1045-R1.xlsx'%sequence
		file_path1 = os.path.abspath(file_pathXL)
		print file_path1

		wb = xlrd.open_workbook(filename = file_path1)
		sh = wb.sheet_by_index(0)

		dilaData = {}

		dilaHeader = {}
		dilaContent = []

		# dilaHeader["Header"] = {}

		dilaHeader["Date/Time"] = {}
		dilaHeader["Date/Time"]["Date"] = sh.cell_value(0, 1)
		dilaHeader["Date/Time"]["Time"] = sh.cell_value(0, 2)
		dilaHeader["Date/Time"]["Time"] = sh.cell_value(0, 3)
		dilaHeader["Date/Time"]["Time"] = sh.cell_value(0, 4)

		dilaHeader["Operator"] = sh.cell_value(1, 1)

		dilaHeader["Laboratory"] = sh.cell_value(2, 1)

		dilaHeader["Sample"] = sh.cell_value(0, 6)
		dilaHeader["value"] = {}
		dilaHeader["value"][""] = sh.cell_value(0, 7)
		dilaHeader["value"]["unit"] = sh.cell_value(0, 8)

		dilaHeader["Reference"] = {}
		dilaHeader["Reference"][""] = sh.cell_value(1, 3)
		dilaHeader["Reference"]["value"] = sh.cell_value(1, 4)
		dilaHeader["Reference"]["unit"] = sh.cell_value(1, 5)

		dilaHeader["Atmosphere"] = {}
		dilaHeader["Atmosphere"]["Element"] = sh.cell_value(2, 3)
		dilaHeader["Atmosphere"]["value"] = sh.cell_value(2, 4)
		dilaHeader["Atmosphere"]["unit"] = sh.cell_value(2, 5)

		dilaHeader["Comment"] = {}
		dilaHeader["Comment"][""] = sh.cell_value(3, 1)
		dilaHeader["Comment"][""] = sh.cell_value(3, 2)
		dilaHeader["Comment"][""] = sh.cell_value(3, 3)
		dilaHeader["Comment"][""] = sh.cell_value(3, 4)

		dilaData["Header"] = dilaHeader

	#### Date repeated a few times here in different format. Not sure if needed or not (6/30 @14:57) ####

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################

		Channel_Header = {}
		Channel_Header["Channel Header"] = {}
		Channel_Header["Channel Header"]["Data Source"] = sh.cell_value(8, 2)
		Channel_Header["Channel Header"]["File Type"] = sh.cell_value(9, 2)
		Channel_Header["Channel Header"]["Number of Channels"] = sh.cell_value(10, 1)

		dilaContent.append(Channel_Header)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################

		
		channel_1 = {}

		channel_1["Channel 1: Type"] = sh.cell_value(11, 1)
		
		channel_1["Sensor Type"] = {}
		channel_1["Sensor Type"][""] = sh.cell_value(12, 2)
		
		channel_1["Sensor Range"] = {}
		channel_1["Sensor Range"]["Minimum"] = {}
		channel_1["Sensor Range"]["Minimum"]["value"] = sh.cell_value(13, 2)
		channel_1["Sensor Range"]["Minimum"]["unit"] = sh.cell_value(13, 4)
		channel_1["Sensor Range"]["Maximum"] = {}
		channel_1["Sensor Range"]["Maximum"]["value"] = sh.cell_value(13, 3)
		channel_1["Sensor Range"]["Maximum"]["unit"] = sh.cell_value(13, 4)

		channel_1["User Range"] = {}
		channel_1["User Range"]["Minimum"] = {}
		channel_1["User Range"]["Minimum"]["value"] = sh.cell_value(14, 2)
		channel_1["User Range"]["Minimum"]["unit"] = sh.cell_value(14, 4)
		channel_1["User Range"]["Maximum"] = {}
		channel_1["User Range"]["Maximum"]["value"] = sh.cell_value(14, 3)
		channel_1["User Range"]["Maximum"]["unit"] = sh.cell_value(14, 4)
		
		channel_1["Minimum"] = {}
		channel_1["Minimum"]["value"] = sh.cell_value(15, 1)
		channel_1["Minimum"]["unit"] = sh.cell_value(15, 2)
		
		channel_1["Maximum"] = {}
		channel_1["Maximum"]["value"] = sh.cell_value(16, 1)
		channel_1["Maximum"]["unit"] = sh.cell_value(16, 2)

		dilaContent.append(channel_1)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################


		channel_2 = {}

		channel_2["Channel 2: Type"] = [sh.cell_value(17, 1)]
		
		channel_2["Sensor Type"] = {}
		channel_2["Sensor Type"][""] = sh.cell_value(18, 2)
		
		channel_2["Sensor Range"] = {}
		channel_2["Sensor Range"]["Minimum"] = {}
		channel_2["Sensor Range"]["Minimum"]["value"] = sh.cell_value(19, 2)
		channel_2["Sensor Range"]["Minimum"]["unit"] = sh.cell_value(19, 4)
		channel_2["Sensor Range"]["Maximum"] = {}
		channel_2["Sensor Range"]["Maximum"]["value"] = sh.cell_value(19, 3)
		channel_2["Sensor Range"]["Maximum"]["unit"] = sh.cell_value(19, 4)

		channel_2["User Range"] = {}
		channel_2["User Range"]["Minimum"] = {}
		channel_2["User Range"]["Minimum"]["value"] = sh.cell_value(20, 2)
		channel_2["User Range"]["Minimum"]["unit"] = sh.cell_value(20, 4)
		channel_2["User Range"]["Maximum"] = {}
		channel_2["User Range"]["Maximum"]["value"] = sh.cell_value(20, 3)
		channel_2["User Range"]["Maximum"]["unit"] = sh.cell_value(20, 4)
		
		channel_2["Minimum"] = {}
		channel_2["Minimum"]["value"] = sh.cell_value(21, 1)
		channel_2["Minimum"]["unit"] = sh.cell_value(21, 2)
		
		channel_2["Maximum"] = {}
		channel_2["Maximum"]["value"] = sh.cell_value(22, 1)
		channel_2["Maximum"]["unit"] = sh.cell_value(22, 2)

		dilaContent.append(channel_2)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################

		channel_3 = {}

		channel_3["Channel 3: Type"] = [sh.cell_value(23, 1)]
		
		channel_3["Sensor Type"] = {}
		channel_3["Sensor Type"][""] = sh.cell_value(24, 2)
		
		channel_3["Sensor Range"] = {}
		channel_3["Sensor Range"]["Minimum"] = {}
		channel_3["Sensor Range"]["Minimum"]["value"] = sh.cell_value(25, 2)
		channel_3["Sensor Range"]["Minimum"]["unit"] = sh.cell_value(25, 4)
		channel_3["Sensor Range"]["Maximum"] = {}
		channel_3["Sensor Range"]["Maximum"]["value"] = sh.cell_value(25, 3)
		channel_3["Sensor Range"]["Maximum"]["unit"] = sh.cell_value(25, 4)

		channel_3["User Range"] = {}
		channel_3["User Range"]["Minimum"] = {}
		channel_3["User Range"]["Minimum"]["value"] = sh.cell_value(26, 2)
		channel_3["User Range"]["Minimum"]["unit"] = sh.cell_value(26, 4)
		channel_3["User Range"]["Maximum"] = {}
		channel_3["User Range"]["Maximum"]["value"] = sh.cell_value(26, 3)
		channel_3["User Range"]["Maximum"]["unit"] = sh.cell_value(26, 4)
		
		channel_3["Minimum"] = {}
		channel_3["Minimum"]["value"] = sh.cell_value(27, 1)
		channel_3["Minimum"]["unit"] = sh.cell_value(27, 2)
		
		channel_3["Maximum"] = {}
		channel_3["Maximum"]["value"] = sh.cell_value(28, 1)
		channel_3["Maximum"]["unit"] = sh.cell_value(28, 2)

		dilaContent.append(channel_3)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################


		channel_4 = {}

		channel_4["Channel 4: Type"] = {}
		channel_4["Channel 4: Type"][""] = sh.cell_value(29, 1)
		channel_4["Channel 4: Type"][""] = sh.cell_value(29, 2)
		
		channel_4["Sensor Type"] = {}
		channel_4["Sensor Type"][""] = sh.cell_value(30, 2)
		
		channel_4["Sensor Range"] = {}
		channel_4["Sensor Range"]["Minimum"] = {}
		channel_4["Sensor Range"]["Minimum"]["value"] = sh.cell_value(31, 2)
		channel_4["Sensor Range"]["Minimum"]["unit"] = sh.cell_value(31, 4)
		channel_4["Sensor Range"]["Maximum"] = {}
		channel_4["Sensor Range"]["Maximum"]["value"] = sh.cell_value(31, 3)
		channel_4["Sensor Range"]["Maximum"]["unit"] = sh.cell_value(31, 4)

		channel_4["User Range"] = {}
		channel_4["User Range"]["Minimum"] = {}
		channel_4["User Range"]["Minimum"]["value"] = sh.cell_value(32, 2)
		channel_4["User Range"]["Minimum"]["unit"] = sh.cell_value(32, 4)
		channel_4["User Range"]["Maximum"] = {}
		channel_4["User Range"]["Maximum"]["value"] = sh.cell_value(32, 3)
		channel_4["User Range"]["Maximum"]["unit"] = sh.cell_value(32, 4)
		
		channel_4["Minimum"] = {}
		channel_4["Minimum"]["value"] = sh.cell_value(33, 1)
		channel_4["Minimum"]["unit"] = sh.cell_value(33, 2)
		
		channel_4["Maximum"] = {}
		channel_4["Maximum"]["value"] = sh.cell_value(33, 1)
		channel_4["Maximum"]["unit"] = sh.cell_value(33, 2)

		dilaContent.append(channel_4)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################


		channel_5 = {}

		channel_5["Channel 5: Type"] = [sh.cell_value(35, 1)]
		
		channel_5["Sensor Type"] = {}
		channel_5["Sensor Type"][""] = sh.cell_value(36, 2)
		
		channel_5["Sensor Range"] = {}
		channel_5["Sensor Range"]["Minimum"] = {}
		channel_5["Sensor Range"]["Minimum"]["value"] = sh.cell_value(37, 2)
		channel_5["Sensor Range"]["Minimum"]["unit"] = sh.cell_value(37, 4)
		channel_5["Sensor Range"]["Maximum"] = {}
		channel_5["Sensor Range"]["Maximum"]["value"] = sh.cell_value(37, 3)
		channel_5["Sensor Range"]["Maximum"]["unit"] = sh.cell_value(37, 4)

		channel_5["User Range"] = {}
		channel_5["User Range"]["Minimum"] = {}
		channel_5["User Range"]["Minimum"]["value"] = sh.cell_value(38, 2)
		channel_5["User Range"]["Minimum"]["unit"] = sh.cell_value(38, 4)
		channel_5["User Range"]["Maximum"] = {}
		channel_5["User Range"]["Maximum"]["value"] = sh.cell_value(38, 3)
		channel_5["User Range"]["Maximum"]["unit"] = sh.cell_value(38, 4)
		
		channel_5["Minimum"] = {}
		channel_5["Minimum"]["value"] = sh.cell_value(39, 1)
		channel_5["Minimum"]["unit"] = sh.cell_value(39, 2)
		
		channel_5["Maximum"] = {}
		channel_5["Maximum"]["value"] = sh.cell_value(40, 1)
		channel_5["Maximum"]["unit"] = sh.cell_value(40, 2)

		dilaContent.append(channel_5)

	########################################################################################
	####--------------------------------------------------------------------------------####
	########################################################################################

	#### Not sure what data represents, putting in for use anyway (6/30 @16:26) ####


		begin = 41

		while True: 
			try: 	

				if sh.cell_value(begin, 0) == "None":
					break

				dataList_row = {}

				dataList_row["Data List"] = {}
				dataList_row["Data List"]["Column 1"] = sh.cell_value(begin, 0)
				dataList_row["Data List"]["Column 2"] = sh.cell_value(begin, 1)
				dataList_row["Data List"]["Column 3"] = sh.cell_value(begin, 2)

				dilaContent.append(dataList_row)
				begin += 1
				
			except: 
				print "first section: %s"%str(begin)
				break

		while True: 
			try:
			 	if sh.cell_value(begin, 0) == "PROBE":
					break	
				begin += 1
		
		begin += 1

		while True:
			try: 
				dataTable_row = {}

				dataTable_row["Data Table"] = {}
				dataTable_row["Data Table"]["Column 1"] = sh.cell_value(begin, 0)
				dataTable_row["Data Table"]["Column 2"] = sh.cell_value(begin, 1)
				dataTable_row["Data Table"]["Column 3"] = sh.cell_value(begin, 2)
				dataTable_row["Data Table"]["Column 4"] = sh.cell_value(begin, 3)
				dataTable_row["Data Table"]["Column 5"] = sh.cell_value(begin, 4)
				
				dilaContent.append(dataTable_row)
				begin += 1

			except: 
				print "second section: %s"%str(begin)
				break

		dilaData["Content"] = dilaContent

		with open('DataFiles/JSON/%s_1045-R1.json'%sequence, 'w') as f:
			f.write(json.dumps(dilaData, sort_keys=True, indent=4, separators=(',', ': ')))

		with open('DataFiles/XML/%s_1045-R1.xml'%sequence, 'w') as f:
			f.write(dicttoxml.dicttoxml(dilaData))

		x = etree.parse("DataFiles/XML/%s_1045-R1.xml"%sequence)
		
		with open('DataFiles/XML/%s_1045-R1.xml'%sequence, 'w') as f:
			f.write(etree.tostring(x, pretty_print = True))				

		print "Conversion for %s to json and xml is done."%file_name

if __name__ == "__main__":

	convert_excel()





	



