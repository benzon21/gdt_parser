import os, csv
from datetime import datetime
from pymsgbox import alert

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
					
	def load_gdt(self,location=curr_dir):
		for root, dirs, files in os.walk(location):
			for file in files:
				if file.endswith(".gdt"):
					file = os.path.join(root,file)
					self.gdt_parser(file,"in")
					self.gdt_parser(file,"mm")

data_folder = "{0}\\Data".format(os.getcwd())

if not os.path.exists(data_folder):
	print("There's no data folder!")
	try:
		os.makedirs(data_folder)
	except Exception:
		alert("Please make a Data Folder","No Folder Found")
		
bl = BlueLightComplier()

bl.load_gdt(data_folder)

bl.to_csv()
