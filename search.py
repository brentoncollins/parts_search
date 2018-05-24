from tkinter import *
from tkinter import messagebox
from xlrd import open_workbook
from openpyxl import load_workbook
from openpyxl.styles.borders import Border, Side


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
class Application:

	def __init__(self, master):
		self.master = master

		try:
			self.items = readxls()
		except FileNotFoundError:
			messagebox.showerror(
				"File not found",
				"Please ensure that 'Parts_List.xlsx' is located in the same directory as the executable."
				"\n\nNo results will be shown.")

			self.items = []

	# Create main GUI window

		self.results = StringVar()
		self.search_var = StringVar()

		self.search_var.trace("w", self.update_list)
		self.entry = Entry(master, textvariable=self.search_var, width=20)
		self.lbox_item = Listbox(master, width=25, height=25)
		self.lbox_brand = Listbox(master, width=20, height=25)
		self.lbox_part = Listbox(master, width=20, height=25)
		self.lbox_stock = Listbox(master, width=15, height=25)
		self.lbox_type = Listbox(master, width=15, height=25)
		self.lbox_disc = Listbox(master, width=50, height=25)
		self.lbox_draw = Listbox(master, width=10, height=25)
		self.lbox_locker = Listbox(master, width=10, height=25)
		self.button = Button(master, text="Add Entry", command=NewEntry)
		
		self.search_label = Label(
			master, text="Marandoo Fixed Plant Electrical Parts Search", font=("Helvetica", 20), anchor="n")
		
		self.entry_label = Label(master, text="Search",font=("Helvetica", 20), fg = "red")
		self.lbox_item_label = Label(master, text="Item")
		self.lbox_brand_label = Label(master, text="Brand")
		self.lbox_part_label = Label(master, text="Part Number")
		self.lbox_stock_label = Label(master, text="Stock Number")
		self.lbox_type_label = Label(master, text="Type")
		self.lbox_disc_label = Label(master, text="Description")
		self.lbox_draw_label = Label(master, text="Draw")
		self.lbox_locker_label = Label(master, text="Locker")

		self.results.set("Enter search term")

		self.results_label = Label(master, textvariable=self.results, font=("Helvetica", 14), fg="blue")
		
		self.results_label.grid(row=1, column=2, padx=10, pady=3, columnspan=3)
		self.button.grid(row=0, column=7, padx=10, pady=3)

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
		
	def update_list(self):

		self.results_label = Label(self.master, font=("Helvetica", 20), fg="red")

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


class NewEntry(object):

	def __init__(self):
		self.window = Toplevel(root)

		draws = [
			'Draw One', 'Draw Two',
			'Draw Three', 'Draw Four',
			'Draw Five', 'Draw Six',
			'Draw Seven', 'Draw Eight',
			'Draw Nine']
		lockers = [
					'Locker A', 'Locker B',
					'Locker C', 'Locker D',
					'Locker E', 'Locker F',
					'Locker G', 'Locker H']

		self.item_var = StringVar()
		self.brand_var = StringVar()
		self.part_var = StringVar()
		self.stock_var = IntVar()
		self.type_var = StringVar()
		self.desc_var = StringVar()
		self.draw_var = StringVar()
		self.locker_var = StringVar()
		self.question_var = StringVar()

		self.lbox_item_entry = Entry(self.window, textvariable=self.item_var, width=25)
		self.lbox_brand_entry = Entry(self.window, textvariable=self.brand_var, width=20)
		self.lbox_part_entry = Entry(self.window, textvariable=self.part_var, width=20)
		self.lbox_stock_entry = Entry(self.window, textvariable=self.stock_var, width=15)
		self.lbox_type_entry = Entry(self.window, textvariable=self.type_var, width=15)
		self.lbox_disc_entry = Entry(self.window, textvariable=self.desc_var, width=50)
		self.lbox_draw_entry = OptionMenu(self.window, self.draw_var, *draws)
		self.lbox_locker_entry = OptionMenu(self.window, self.locker_var, *lockers)

		self.entry_label = Label(self.window, text="Enter New Item", font=("Helvetica", 16), fg="blue")
		self.lbox_item_label_entry = Label(self.window, text="Item")
		self.lbox_brand_label_entry = Label(self.window, text="Brand")
		self.lbox_part_label_entry = Label(self.window, text="Part Number")
		self.lbox_stock_label_entry = Label(self.window, text="Stock Number")
		self.lbox_type_label_entry = Label(self.window, text="Type")
		self.lbox_disc_label_entry = Label(self.window, text="Description")
		self.lbox_draw_label_entry = Label(self.window, text="Draw")
		self.lbox_locker_label_entry = Label(self.window, text="Locker")

		self.entry_label.grid(row=0, column=0, padx=10, pady=3)
		self.lbox_item_label_entry.grid(row=3, column=0, padx=10, pady=3)
		self.lbox_locker_label_entry.grid(row=3, column=1, padx=10, pady=3)
		self.lbox_draw_label_entry.grid(row=3, column=2, padx=10, pady=3)
		self.lbox_brand_label_entry.grid(row=3, column=3, padx=10, pady=3)
		self.lbox_type_label_entry.grid(row=3, column=4, padx=10, pady=3)
		self.lbox_stock_label_entry.grid(row=3, column=5, padx=10, pady=3)
		self.lbox_part_label_entry.grid(row=3, column=6, padx=10, pady=3)
		self.lbox_disc_label_entry.grid(row=3, column=7, padx=10, pady=3)

		self.entry_label.grid(row=0, column=0, padx=10, pady=3, columnspan=2)
		self.lbox_item_entry.grid(row=4, column=0, padx=10, pady=3)
		self.lbox_locker_entry.grid(row=4, column=1, padx=10, pady=3)
		self.lbox_draw_entry.grid(row=4, column=2, padx=10, pady=3)
		self.lbox_brand_entry.grid(row=4, column=3, padx=10, pady=3)
		self.lbox_part_entry.grid(row=4, column=4, padx=10, pady=3)
		self.lbox_stock_entry.grid(row=4, column=5, padx=10, pady=3)
		self.lbox_type_entry.grid(row=4, column=6, padx=10, pady=3)
		self.lbox_disc_entry.grid(row=4, column=7, padx=10, pady=3)
		self.button = Button(self.window, text="Add Item", command=self.show_btn_lbl)

		self.button.grid(row=0, column=2, padx=10, pady=3, columnspan=2)

		self.question_label = Label(
											self.window,
											textvariable=self.question_var,
											font=("Helvetica", 14), fg="black")
		#
		self.question_label.grid(row=5, column=0, padx=10, pady=3, columnspan=2)

		self.yes_button = Button(self.window, text="Add Entry", command=self.add_item)
		self.yes_button.grid(row=5, column=3, padx=10, pady=3)

		self.no_button = Button(self.window, text="Change", command=self.remove_btn_lbl)
		self.no_button.grid(row=5, column=4, padx=10, pady=3)

		self.question_label.grid_remove()
		self.no_button.grid_remove()
		self.yes_button.grid_remove()

	def show_btn_lbl(self):
		self.question_label.grid()
		self.no_button.grid()
		self.yes_button.grid()
		try:
			self.question_var.set("Are you sure you want to add\n"
			" {} {} {} {} {} {} {} {}?".format(
				self.item_var.get().title(),
				self.draw_var.get(),
				self.locker_var.get(),
				self.brand_var.get().title(),
				self.part_var.get().upper(),
				self.stock_var.get(),
				self.type_var.get().upper(),
				self.desc_var.get().title()))
		except TclError:
			messagebox.showerror("Error", "Please only use a number for the Stock Number.")

	def remove_btn_lbl(self):
		self.question_label.grid_remove()
		self.no_button.grid_remove()
		self.yes_button.grid_remove()

	def add_item(self):

		book = open_workbook("Parts_List.xlsx")
		wb = load_workbook(filename='Parts_List.xlsx')
		# Open workbook
		sheet_names = book.sheet_names()
		# Get sheet names
		no_sheets = book.nsheets
		# Set sheet number to x, add 1 when every sheet has finished.
		for x in range(0, no_sheets):
			if sheet_names[x] == self.locker_var.get():
				sheet = book.sheet_by_index(x)
				for row in range(2, sheet.nrows):
					if self.draw_var.get() in str(sheet.cell(row, 0)):
						ws = wb.worksheets[x]
						ws.insert_rows(row + 3, 1)
						thin_border = Border(
												right=Side(style='medium'),
												bottom=Side(style=None)
												)
						try:
							self.stock_var.get()

						except TclError:
							messagebox.showerror("Error", "Please only use a number for the Stock Number.")

						ws.cell(column=1, row=row + 3, value=self.item_var.get().title()).border = thin_border
						ws.cell(column=2, row=row + 3, value=self.brand_var.get().title()).border = thin_border
						ws.cell(column=3, row=row + 3, value=self.part_var.get().upper()).border = thin_border
						ws.cell(column=4, row=row + 3, value=self.stock_var.get()).border = thin_border
						ws.cell(column=5, row=row + 3, value=self.type_var.get().upper()).border = thin_border
						ws.cell(column=6, row=row + 3, value=self.desc_var.get().title()).border = thin_border

						try:
							wb.save("Parts_List.xlsx")
						except PermissionError:
							messagebox.showerror(
								"Error",
								"No permission to write to the Parts List, please close the "
								"document or gain permission to write to the file.")
							break
						self.lbox_item_entry.delete(0, 'end')
						self.lbox_brand_entry.delete(0, 'end')
						self.lbox_part_entry.delete(0, 'end')
						self.stock_var.set(0)
						self.lbox_type_entry.delete(0, 'end')
						self.lbox_disc_entry.delete(0, 'end')
						self.items = readxls()
						print("Done")
						self.remove_btn_lbl()
						self.question_label.grid()
						self.question_var.set("Done")


root = Tk()
root.title('Electrical workshop parts search.')

app = Application(root)
print('Starting mainloop()')

root.mainloop()


# c:\python27\Scripts\pyinstaller.exe --onefile --hidden-import=xlrd -w --windowed --noconsole search.py
