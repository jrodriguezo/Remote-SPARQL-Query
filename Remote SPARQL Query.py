#Imports
from tkinter import *
from tkinter import scrolledtext
from tkinter import filedialog
from tkinter import ttk
from tkinter.messagebox import showinfo
import subprocess
from subprocess import call
from subprocess import *
import os 
import xlsxwriter
from SPARQLWrapper import SPARQLWrapper, JSON


class Main:
	def __init__(self, window):

		#Initialization to know when he/she have clicked on "search"
		self.has_been_called = False
		self.wind = window
		#Full window
		self.wind.state('zoomed')


		#To use an icon
		#self.wind.iconbitmap("lens.ico")

		#Define title app
		self.wind.title("Remote SPARQL Query")

		#######################################
		######   Open file directory   ########
		#######################################

		#Creating a Frame Container -1
		self.frame = LabelFrame(self.wind, text="Dataset")
		self.frame.grid(row = 0, column = 0,padx =10,pady=2, sticky= W+E)
		#To create text and buttons
		Label(self.frame, text = "Ontology: ").grid(row = 1, column= 1)
		self.b1 = Button(self.frame, text="Choose file", command=self.mopenfile)
		self.b1.grid(row = 1, column = 2, columnspan = 2 , pady = 5, padx =10, sticky= W+E)
		self.l1 = Label(self.frame, text="No file chosen")
		self.l1.grid(row = 1, column = 4, padx = 30)

		##############################
		###### Insert query ##########
		##############################
		#Creating a Frame Container -2
		self.frame2 = LabelFrame(self.wind, text="Queries")
		self.frame2.grid(row = 2, column = 0, padx= 10,pady=2, sticky= W+E)

		#Query text entry
		self.query = scrolledtext.ScrolledText(self.frame2 , height = 15,width=90, undo=True)
		self.query.grid(row = 3, column = 1, pady = 5,padx = 10, sticky= W+E)
		#Initial text
		self.query.insert(INSERT,' SELECT * WHERE{?s ?p ?o } LIMIT 1000 ')

		self.b2 = Button(self.frame2, text="Search", command=self.mresults,width =10)
		self.b2.grid(row=4,column=1,  pady = 5, padx =10)

		##############################
		###### Methods definition ####
		##############################

	#Method to open a file
	def mopenfile(self, event=None):
		self.wind.filename =  filedialog.askopenfilename()
		#Directory where Fuseki is found
		#IMPORTANT change if user want another directory
		os.chdir(os.getcwd()+"\\fuseki")
		#Initialized server and you can access through localhost:3030
		s = subprocess.Popen(['java', '-jar', 'fuseki-server.jar','--file', self.wind.filename,'/Data'])
		#Change text when is loaded
		self.l1.config(text = self.wind.filename )
		#Button disabled, the file can not be updated
		self.b1.config(state = 'disabled')

	#Method to destroy frame 3 when click on "search" button and call mresults()
	def tdestroy(self):
		self.frame3.grid_forget()
		self.mresults()
	#Method to request to the endpoint the query  and in the reply is put in a table
	def mresults(self):
		#It happens when user has clicked two or more times in "search", 
		#i.e destroy the previous frame to show the new results
		if self.has_been_called:
			self.has_been_called = False
			self.tdestroy()
		else:
			#Creating a Frame Container -3
			self.frame3 = LabelFrame(self.wind, text="Results")
			self.frame3.grid(row = 5, column = 0, padx=10,pady=2)

			#Creating a table
			self.tree = ttk.Treeview(self.frame3)
			self.tree.grid(row = 7, column = 0, pady = 10,padx = 10)

			#The method has been called 
			self.has_been_called = True
			pass
		
			#Scrollbars in table
			ysb = ttk.Scrollbar(self.frame3, orient='vertical', command=self.tree.yview)
			ysb.grid(row=7, column = 6, sticky='ns', rowspan=3)
			xsb = ttk.Scrollbar(self.frame3, orient='horizontal', command=self.tree.xview)
			xsb.grid(row=8, column=0, sticky='we',columnspan =3)

			self.tree.configure(yscrollcommand=ysb.set)
			self.tree.configure(xscrollcommand=xsb.set)

			#Getting data from endpoint
			sparql = SPARQLWrapper("http://localhost:3030/Data/sparql")
			
			#Query inserted by user
			sparql.setQuery(self.query.get(1.0,END))

			#Getting reply in JSON format
			sparql.setReturnFormat(JSON)
			results = sparql.query().convert()

			#To put headers in table
			self.tree['show'] = 'headings'
			self.tree['columns'] = results["head"]["vars"]
			i = 0
			list= []
			#To put results in table
			for result in results["results"]["bindings"]:
				variables = []		
				for var in results["head"]["vars"]:
					if var not  in result:
						variables.append("empty")
						continue
					variables.append(result[var]["value"])
				list.append(variables)
			
			self.l3 = Label(self.frame3, text=str(len(list))+" entries")
			self.l3.grid(row = 6, column = 0,padx=10, sticky = W)

			for element in list:
				self.tree.insert("", "end", values= element)

			#Depends on the variables that have to be presented uses different sizes
			for var in results["head"]["vars"]:
				if len(results["head"]["vars"])<4:
					self.tree.heading("#"+str(i+1), text=var,anchor ="center")
					self.tree.column("#"+str(i+1), width=400, stretch="no")
				elif len(results["head"]["vars"])<7:
					
					self.tree.heading("#"+str(i+1), text=var,anchor ="center")
					self.tree.column("#"+str(i+1),minwidth =300, width=200, stretch=True)
				else:
					self.tree.heading("#"+str(i+1), text=var,anchor ="center")
					self.tree.column("#"+str(i+1),minwidth =300, width=100, stretch=True)
				i += 1
			#Variables to save in Excel, called by save_toexcel method
			self.listToSave = list
			self.resultsToSave = results["head"]["vars"]

			#Button to save
			self.save = Button(self.frame3, text="Save", command=self.save_toexcel,width =10)
			self.save.grid(row=6,column=0,pady = 5,padx = 10, sticky = E)

			self.l2 = Label(self.frame3, text="No path selected")
			self.l2.grid(row = 6, column = 0, padx= 100, sticky = E)

	#Method to save results in Excel
	def save_toexcel(self):
		#Ask the user the directory where he/she want to save the excel
		self.wind.filename =  filedialog.askdirectory()
		#Change text
		self.l2.configure(text = "Saved in: "+self.wind.filename)
	
		workbook = xlsxwriter.Workbook(self.wind.filename+'\\query_results.xlsx')
		#Create worksheet
		worksheet = workbook.add_worksheet()
		#Custom cells
		cell_format = workbook.add_format()
		cell_format.set_bold()
		cell_format.set_font_color('white')
		cell_format.set_bg_color('black')

		#Insert the headers in the first row and their corresponding columns
		col = 0
		row = 0
		for col, data in enumerate(self.resultsToSave):
			worksheet.write(row, col, data,cell_format)
			worksheet.set_column(col, len(self.resultsToSave), 60)

		#Insert the results in their corresponding columns and rows
		col = 0
		for row, data in enumerate(self.listToSave):
			worksheet.write_row(row+1, col, data)
		#End
		workbook.close()

# Class to copy, cut, and paste with ctrl+c, ctrl+x and ctrl+v in text inputs
class Test(Text):
	def __init__(self, master, **kw):
		Text.__init__(self, master, **kw)
		self.bind('<Control-c>', self.copy)
		self.bind('<Control-x>', self.cut)
		self.bind('<Control-v>', self.paste)
		
	def copy(self, event=None):
		self.clipboard_clear()
		text = self.get("sel.first", "sel.last")
		self.clipboard_append(text)
	
	def cut(self, event):
		self.copy()
		self.delete("sel.first", "sel.last")

	def paste(self, event):
		text = self.selection_get(selection='CLIPBOARD')
		self.insert('insert', text)


if __name__ == '__main__':
	#Create Window object
	#All window objects must be declared between those lines
	window = Tk()
	application = Main(window)
	#This function calls the endless cycle of the window
	#Window will wait for any user interaction until we close it
	window.mainloop()





