import os, csv
from datetime import datetime
from xlsxwriter import Workbook

class BlueLightComplier:
	data = {}
	file_part_name = None
	sample_name = None
	curr_dir = os.getcwd()
	
	def name_parser(self,file_name):
		gdt_file = file_name.split("\\")[-1]
		savename = gdt_file.split(".")[0]
		file_part = savename.split(" ")
		self.file_part_name = file_part[0]
		self.sample_name = " ".join(file_part[1:])
	
	def gdt_parser(self,file_name,units):
		first_run = False
		self.name_parser(file_name)
		part_key = "{0}({1})".format(self.file_part_name,units)
		if part_key not in self.data:
			self.data[part_key] = []
			first_run = True
			
		wb = self.data[part_key]
		
		with open(file_name,"r+") as f:
			content = [x.split(";") for x in f.readlines() if x != "\n"] 

		rm_chars = lambda x: x.replace('\n','').replace(' ','')
		
		values = [rm_chars(x[5]) for x in content[1:] if rm_chars(x[5]) != '']
		
		in_mm = {'in': 1 , 'mm': 25.4}
		
		values = [float(x) / in_mm[units] for x in values if float(x) != 0]
		
		values.insert(0,self.sample_name)
		
		if first_run:
			titles = [x[0] for x in content[1:]]
			titles.insert(0,"Sample ID")
			wb.append(titles[:len(values)])

		wb.append(values)
		
	def to_csv(self,location=curr_dir):
		for key, values in self.data.items():
			csv_file = "{0}//{1}.csv".format(location,key)
			with open(csv_file,"a+",newline='') as excel:
				ws = csv.writer(excel,dialect='excel')
				for val in values:
					ws.writerow(val)
					
	def _column_letter(self,col_idx):
		letters = []
		while col_idx > 0:
			col_idx, remainder = divmod(col_idx, 26)
			if remainder == 0:
				remainder = 26
				col_idx -= 1
			letters.append(chr(remainder+64))
		return ''.join(reversed(letters))
		
	def to_xlsx(self,location=curr_dir):
		for key, values in self.data.items():
			xlsx_file = "{0}//{1}.xlsx".format(location,key)
			workbook = Workbook(xlsx_file)
			worksheet = workbook.add_worksheet()
			for idx,val in enumerate(values):
				loc = "A" + str(idx + 1)
				worksheet.write_row(loc, val)
				
			#conditional formatting
			fail_color = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})
			pass_color = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
			last_cell = self._column_letter(len(values[-1])) + str(len(values))
			limits = (-0.008,0.008) if 'in' in key else (-0.2,0.2)
			worksheet.conditional_format('B2:'+last_cell,{'type' : 'cell',
														 'criteria' : 'between',
														 'minimum' : limits[0],
														 'maximum' : limits[1],
														 'format' : pass_color
														})
			worksheet.conditional_format('B2:'+last_cell,{'type' : 'cell',
														 'criteria' : 'not between',
														 'minimum' : limits[0],
														 'maximum' : limits[1],
														 'format' : fail_color
														})
			workbook.close()
					
	def load_gdt(self,location=curr_dir):
		for root, dirs, files in os.walk(location):
			for file in files:
				if file.endswith(".gdt"):
					file = os.path.join(root,file)
					self.gdt_parser(file,"in")
					self.gdt_parser(file,"mm")

data_folder = "{0}\\Data".format(os.getcwd())
		
bl = BlueLightComplier()

bl.load_gdt(data_folder)

bl.to_csv()
#bl.to_xlsx()
