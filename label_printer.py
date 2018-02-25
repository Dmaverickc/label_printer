import csv 
import os
import labels
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase.pdfmetrics import registerFont, stringWidth
from reportlab.graphics import shapes
from reportlab.lib import colors
from Tkinter import Tk
from tkFileDialog import askopenfilename
import pandas as pd
import ntpath

specs = labels.Specification(210, 297, 2, 8, 90, 25, corner_radius=2)

##### muiltproscessing has to be done & the lenght of the str 

def file_converter(file_obj):
	
	file_obj = ntpath.basename(file_obj)  # just takes the file name from the path
	file_n, file_ex = os.path.splitext(file_obj) # splits the file into 2 name and exetionsion 	
	file_csv = pd.read_excel(file_obj, file_n, index_col=None) # this is passed the file name and the name of the sheet first sheet
	file_csv.to_csv('file_label.csv', encoding='utf-8', index=False) #  converts xlsx file to csv



def csv_reader(file_obj):

	reader = csv.DictReader(file_obj, delimiter =',') # reads in csv file
	
	prev_row = None   # sets the preivous version to null 
	prev_item = None
	group_label = []
	single_label = []

	for idx, row in enumerate(reader):       # goes through the rows of the csv
		
		item = row["letter"]   # set item as row on time from csv file

		single_label.append(item)

		if row["name"] == prev_row: # equals prev_row so label can build
		
			group_label.append(item)		# appends item to the label list
			group_label.append(prev_item)	 # appends the last item to the list so it does skip any
			group_label = sorted(set(group_label))  # sorts the removes doubles

		elif row["name"] != prev_item:

			# creates a list of all the single label values by removing the group_labels of labels from it 
			for x in group_label:
				for y in single_label:
					if x == y:
						single_label.remove(x)

			group_label = position_label(group_label)
			group_label = [] # clears the list group_list before the next incrememnt 
			
		
		prev_row = row["name"]     # sets the current version equal to to previous vresion before next increment
		prev_item = row["letter"]

	single_position(single_label)  # has to be outside the for loop so single_positon gets the end list of single_labels

def draw_label(label, width, height, obj):

	font_size = 28
	f_size = 20
	text_width = width - 5
	name_width = stringWidth(obj, "Helvetica", font_size)


	while name_width > text_width: # compresses the label to fit within its boundaries 
	  	font_size *= 0.9
		name_width = stringWidth(obj, "Helvetica", font_size)



	firstpart, secondpart = obj[:len(obj)/2], obj[len(obj)/2:]

	t = shapes.String(width/2.0, 30, obj, textAnchor="middle")
	s = shapes.String(width/2.0, 35, firstpart, textAnchor="middle")
	e = shapes.String(width/2.0, 19, secondpart, textAnchor="middle")


	e.fontSize = f_size
	s.fontSize = f_size
	t.fontSize = font_size

 ####need to figure a test for str lenght to format it##########
	if len(obj) <= 23:
		label.add(t)
	elif len(obj) <= 33:
		label.add(s)
		label.add(e)
	

def position_label(p_label):  
		# loops through the group to get a count on how many items are in each list 
		# and returns the value to csv_reader
	for i, s in enumerate(p_label):	#gets rid of all the lists before the are in complete. i.e it takes away not finished label lists 
		if p_label == 0:   #checks to see if lenght of p_label 		
			p_label = " ".join(p_label)
			sheet.add_label(p_label) # prints the p_label to sheet
		elif len(p_label) == 2: 
			p_label = " ".join(p_label)
			sheet.add_label(p_label)
		elif len(p_label) == 3:
			p_label = " ".join(p_label)
			sheet.add_label(p_label) 
		elif len(p_label) == 4:
			p_label = " ".join(p_label)
			sheet.add_label(p_label) 

def single_position(s_label):
	s = "" # creates a str without any spaces
	s = s.join(s_label) #  joins the s_label str together
	for i in s_label: # goes throught the the str of s_label
	 	print i    
		sheet.add_label(i)  #adds every individual item(i) of single_label to the label sheet	

sheet = labels.Sheet(specs, draw_label, border=True)

if __name__ == "__main__":

	Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
	filename = askopenfilename() # show an "Open" dialog box and return the path to the selected file

	file = ntpath.basename(filename) # just takes out the file name from the path 

	file = "".join(file) # turns files name into string
 
	suffix = "xlsx"  # creates a string that has the valve of xlsx

	if file.endswith(suffix) == False:   # checks to see if the file type is true
		with open(filename) as f_obj:   # opens the csv file
			# still have to get this to print to document, might try multiprocessing to do this
			csv_reader(f_obj)             # sends the csv_reader function the csv data file 
	elif file.endswith(suffix) == True:
		file_converter(filename)
		with open('file_label.csv') as f_obj:   # opens the csv file
		 	csv_reader(f_obj)        
		 	f_obj.close()   # closes f_obj
			os.remove('file_label.csv') # deletes the temp file 



	sheet.save('basic1.pdf')
	print("{0:d} label(s) output on {1:d} page(s).".format(sheet.label_count, sheet.page_count))