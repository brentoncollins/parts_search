from tkinter import *
from tkinter import messagebox
from xlrd import open_workbook


def readxls():
	book = open_workbook("Parts_List.xlsx")
	# Open workbook
	sheet_names = book.sheet_names()
	# Get sheet names
	no_sheets = book.nsheets
	# Get number of sheets
	workbook = []
	# Create empty list for all data in workbook
	sheet_num = 0
	# Set sheet number to x, add 1 when every sheet has finished.
	for x in range(0, no_sheets):
		# Assign each sheet in every loop
		sheet = book.sheet_by_index(x)
		# Clear the item list after every sheet, this wil differentiate between the lockers.
		item_list = []

		draw = 0
		# Start at draw 0, when finding the first draw in the spreadsheet add one.
		for row in range(2, sheet.nrows):
			if "Draw" in str(sheet.cell(row, 0)):
				# If found draw, move to the next draw.
				draw += 1
			else:

				if "Item" in str(sheet.cell(row, 0)) or sheet.cell(row, 0).value == "":
					# Ignore all rows that start with item or is empty.
					pass

				else:
					# Append all items to a dict with the item as the key, all aother data are the values.

					# For the numbers that are converted into floats change them to int
					if type(sheet.cell(row, 3).value) == float:
						stock_no = int(sheet.cell(row, 3).value)
					else:
						stock_no = ""

					if type(sheet.cell(row, 2).value) == float:
						part_number = int(sheet.cell(row, 2).value)
					else:
						part_number = sheet.cell(row, 2).value

					if type(sheet.cell(row, 4).value) == float:
						part_type = int(sheet.cell(row, 4).value)
					else:
						part_type = sheet.cell(row, 4).value
					# Then append the values
					item_list.append(
								[{sheet.cell(row, 0).value: ["Draw " + str(draw),
									sheet.cell(row, 1).value,
									part_number,
									stock_no,
									part_type,
									sheet.cell(row, 5).value,
									sheet_names[sheet_num]
									]}])
		# When done with the sheet all the list of items to the workbook list.
		workbook.append(item_list)
		# Move to the next sheet.
		sheet_num += 1
	# Return all data for all items.
	return workbook


	# First create application class
class Application(Frame):

	def __init__(self, master=None):
		Frame.__init__(self, master)

		self.pack()
		try:
			self.items = readxls()
		except FileNotFoundError:
			messagebox.showerror("File not found",
							"Please ensure that 'Parts_List.xlsx' is located in the same directory as the executable."
							"\n\nNo resutls will be shown.")
			self.items = []


	# Create main GUI window

		self.results = StringVar()
		self.search_var = StringVar()
		self.search_var.trace("w", self.update_list)
		self.entry = Entry(self, textvariable=self.search_var, width=20)
		self.lbox_item = Listbox(self, width=25, height=25)
		self.lbox_brand = Listbox(self, width=20, height=25)
		self.lbox_part = Listbox(self, width=20, height=25)
		self.lbox_stock = Listbox(self, width=15, height=25)
		self.lbox_type = Listbox(self, width=15, height=25)
		self.lbox_disc = Listbox(self, width=50, height=25)
		self.lbox_draw = Listbox(self, width=10, height=25)
		self.lbox_locker = Listbox(self, width=10, height=25)
		# self.button = Button(self, text="About", command=self.create_window)
		
		self.search_label = Label(self, text="Marandoo Fixed Plant Electrical Parts Search", font=("Helvetica", 20), anchor="n" )
		
		self.entry_label = Label(self, text="Search",font=("Helvetica", 20), fg = "red")
		self.lbox_item_label = Label(self, text="Item")
		self.lbox_brand_label = Label(self, text="Brand")
		self.lbox_part_label = Label(self, text="Part Number")
		self.lbox_stock_label = Label(self, text="Stock Number")
		self.lbox_type_label = Label(self, text="Type")
		self.lbox_disc_label = Label(self, text="Description")
		self.lbox_draw_label = Label(self, text="Draw")
		self.lbox_locker_label = Label(self, text="Locker")

		self.results.set("Enter search term")

		self.results_label = Label(self, textvariable=self.results, font=("Helvetica", 14), fg="blue")
		
		self.results_label.grid(row=1, column=2, padx=10, pady=3, columnspan=3)
		# self.button.grid(row=0, column=7, padx=10, pady=3)

		self.search_label.grid(row=0, column=2, padx=10, pady=3, columnspan=4)
		self.entry_label.grid(row=1, column=0, padx=10, pady=3)
		self.lbox_item_label.grid(row=3, column=0, padx=10, pady=3)
		self.lbox_locker_label.grid(row=3, column=1, padx=10, pady=3)
		self.lbox_draw_label.grid(row=3, column=2, padx=10, pady=3)
		self.lbox_brand_label.grid(row=3, column=3, padx=10, pady=3)
		self.lbox_type_label.grid(row=3, column=4, padx=10, pady=3)
		self.lbox_stock_label.grid(row=3, column=5, padx=10, pady=3)
		self.lbox_part_label.grid(row=3, column=6, padx=10, pady=3)
		self.lbox_disc_label.grid(row=3, column=7, padx=10, pady=3)

		self.entry.grid(row=1, column=1, padx=10, pady=3,columnspan=2)
		self.lbox_item.grid(row=4, column=0, padx=10, pady=3)
		self.lbox_locker.grid(row=4, column=1, padx=10, pady=3)
		self.lbox_draw.grid(row=4, column=2, padx=10, pady=3)
		self.lbox_brand.grid(row=4, column=3, padx=10, pady=3)
		self.lbox_part.grid(row=4, column=4, padx=10, pady=3)
		self.lbox_stock.grid(row=4, column=5, padx=10, pady=3)
		self.lbox_type.grid(row=4, column=6, padx=10, pady=3)
		self.lbox_disc.grid(row=4, column=7, padx=10, pady=3)

		# Function for updating the list/doing the search.
		# It needs to be called here to populate the listbox.
		self.update_list()
		
	def update_list(self, *args):

		self.results_label = Label(self,font=("Helvetica", 20), fg="red")

		search_term = self.search_var.get()

		self.lbox_item.delete(0, END)
		self.lbox_brand.delete(0, END)
		self.lbox_part.delete(0, END)
		self.lbox_stock.delete(0, END)
		self.lbox_type.delete(0, END)
		self.lbox_disc.delete(0, END)
		self.lbox_draw.delete(0, END)
		self.lbox_locker.delete(0, END)
		for y in self.items:
			for item in y:
				for key, value in item[0].items():
					for val in value:
						if search_term.lower() in str(val).lower() or search_term.lower() in key.lower():
							self.lbox_item.insert(END, key)
							self.lbox_draw.insert(END, value[0])
							self.lbox_brand.insert(END, value[1])
							self.lbox_type.insert(END, value[2])

							self.lbox_stock.insert(END, value[3])
							self.lbox_part.insert(END, value[4])
							self.lbox_disc.insert(END, value[5])
							self.lbox_locker.insert(END, value[6])
							
							break

		for x in range(0, self.lbox_item.size()):
			if x & 1:
				self.lbox_locker.itemconfig(x, {'bg': 'light blue'})
				self.lbox_draw.itemconfig(x, {'bg': 'light blue'})
				self.lbox_disc.itemconfig(x, {'bg': 'light blue'})
				self.lbox_type.itemconfig(x, {'bg': 'light blue'})
				self.lbox_stock.itemconfig(x, {'bg': 'light blue'})
				self.lbox_part.itemconfig(x, {'bg': 'light blue'})
				self.lbox_brand.itemconfig(x, {'bg': 'light blue'})
				self.lbox_item.itemconfig(x, {'bg': 'light blue'})
		self.results.set("{} Results Found".format(int(self.lbox_item.size())))

	# def create_window(self):
	# 	t = Toplevel(self)
	# 	# t.overrideredirect(True)
	# 	self.created_by = Label(t, text ="Created by Brenton Collins")
	# 	self.contact = Label(t, text="Any issues or improvement ideas?"
	# 							"\nContact : brenton.collins@riotinto.com")
	#
	# 	# self.exit = Button(t,text = "Close", command = t.destroy())
	#
	# 	self.created_by.grid(row=1, column=1, padx=10, pady=3, columnspan=2)
	# 	self.contact.grid(row=2, column=1, padx=10, pady=3, columnspan=2)
	# 	# self.exit.grid(row=3, column=1, padx=10, pady=3, columnspan=2)




root = Tk()
root.title('Electrical workshop parts search.')

app = Application(master=root)
print('Starting mainloop()')

app.mainloop()


# c:\python27\Scripts\pyinstaller.exe --onefile --hidden-import=xlrd -w --windowed --noconsole search.py