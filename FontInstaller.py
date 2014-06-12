import os
import os.path
import zipfile
import re
from tkinter import *
from tkinter.filedialog import *
from shutil import *

class gui:
	"""the gui of FontInstaller"""
	def __init__(self):
		window = Tk()
		# title
		window.title('Font Installer Beta 1.0')
		# size
		window.geometry('450x280')
		# icon
		#window.iconbitmap()
		
		# menu bar
		menubar = Menu(window)
		window.config(menu = menubar)
		aboutMenu = Menu(menubar, tearoff = 0)
		# extract
		menubar.add_cascade(label = "Extract", command = self.extract)
		# install
		menubar.add_cascade(label = "Install", command = self.install)
		# other menu
		menubar.add_cascade(label = "Other", menu = aboutMenu)
		aboutMenu.add_command(label = "Help", command = self.dialogHelp)
		aboutMenu.add_separator()
		aboutMenu.add_command(label = "About", command = self.dialogAbout)
		# exit menu
		exitmenu = Menu(menubar, tearoff = 0)
		menubar.add_cascade(label = "Exit", menu = exitmenu)
		exitmenu.add_command(label = "Quit", command = window.quit)

		# popup menus
		self.popupMenu = Menu(window, tearoff = 0)
		self.popupMenu.add_command(label = "Quit", command = window.quit)
		window.bind("<Button-3>", self.popup)

		# font path selector
		frameFont = Frame(window)
		frameFont.pack()
		labelFont = Label(frameFont, text = "Font root path:")
		self.fontPath = StringVar()
		self.fontPath.set('Unselected')
		self.entryFont = Entry(frameFont, textvariable = self.fontPath, state = 'readonly')
		btFont = Button(frameFont, text = "Browse", command = self.fontPathSelect)
		# widget grid
		labelFont.grid(row = 1, column = 0, sticky = W, padx = 10, pady = 20)
		self.entryFont.grid(row = 1, column = 1, sticky = W, padx = 3, ipadx = 50)
		btFont.grid(row = 1, column = 2, sticky = W)

		# target path selector
		frameTarget = Frame(window)
		frameTarget.pack()
		labelTarget = Label(frameTarget, text = "Target root path:")
		self.targetPath = StringVar()
		self.targetPath.set('C:\Windows\Fonts')
		self.entryTarget = Entry(frameTarget, textvariable = self.targetPath, state = 'readonly')
		btTarget = Button(frameTarget, text = "Browse", command = self.targetPathSelector)
		# widget grid
		labelTarget.grid(row = 1, column = 0, sticky = W, padx = 3)
		self.entryTarget.grid(row = 1, column = 1, sticky = W, padx = 3, ipadx = 50)
		btTarget.grid(row = 1, column = 2, sticky = W)


		# Control
		frameControl = Frame(window)
		frameControl.pack()
		self.judTarget = IntVar()
		self.judTarget.set(0)
		cbTarget = Checkbutton(frameControl, text = "Extract to target folder", variable = self.judTarget)
		self.judDelete = IntVar()
		self.judDelete.set(1)
		cbDelete = Checkbutton(frameControl, text = "Delete other files", variable = self.judDelete)
		# widget grid
		cbTarget.grid(row = 1, column = 1, padx = 20, pady = 10)
		cbDelete.grid(row = 1, column = 2, padx = 20)

		# infomation
		frameInfo = Frame(window)
		frameInfo.pack()
		scrollbar = Scrollbar(frameInfo)
		scrollbar.pack(side = RIGHT, fill = Y)
		self.info = Text(frameInfo, width = 150, height = 60, yscrollcommand = scrollbar.set)
		self.info.pack()
		self.info.insert(END, 'Here are many unexpected bug becouse of beta version.\nPleae backup before use.\nUse Step:\n1.Select font root path.\n2.Select target root path.\n3.Extract.\n4.Install.\n\n')
		scrollbar.config(command = self.info.yview)

		window.mainloop()

	def extract(self):
		if not os.path.isdir(self.targetPath.get()):
			os.makedirs(self.targetPath.get())
		fontPath = self.fontPath.get()
		if not os.path.isdir(fontPath):
			pass
		else:
			for root, dirs, files in os.walk(fontPath):
				for fn in files:
					sufix = os.path.splitext(fn)[1][1:]
					if sufix == 'zip' or sufix == 'ZIP':
						self.info.insert(END, 'Extracted: '+fn+'\n')
						f = zipfile.ZipFile(root+'\\'+fn)
						if self.judTarget.get():
							f.extractall(self.targetPath.get()+'\\')
						else:
							f.extractall(root+'\\')

	def install(self):
		fontPath = self.fontPath.get()
		if self.judTarget.get():
			fontPath = self.targetPath.get() 
		targetPath = self.targetPath.get()
		if not os.path.isdir(targetPath):
			os.makedirs(targetPath)
		for root, dirs, files in os.walk(fontPath):
			for fn in files:
				sufix = os.path.splitext(fn)[1][1:]
				if sufix == 'ttf' or sufix == 'TTF' or sufix == 'otf' or sufix == 'OTF':
					temp = ''
					for c in root:
						if c == '/':
							temp += '\\'
						else:
							temp += c
					os.system('Installfont.vbs'+' '+temp+'\\'+fn)
					self.info.insert(END, 'Installfont.vbs'+' '+temp+'\\'+fn+'\n'+'Installed: '+fn+'\n')
					if isDelete:
						os.remove(root+'\\'+fn)
						self.info.insert(End, '')
				elif sufix == 'zip' or sufix == 'ZIP' or sufix == 'rar' or sufix == 'RAR':
					pass
				elif self.judDelete.get():
					os.remove(root+'\\'+fn)
					self.info.insert(END, 'Deleted: '+fn+'\n')

	def dialogAbout(self):
		pass

	def dialogHelp(self):
		pass

	def popup(self, event):
		self.popupMenu.post(event.x_root, event.y_root)

	def fontPathSelect(self):
		fontPath = askdirectory(initialdir="/",title='Pick a directory')
		self.fontPath.set(fontPath)

	def targetPathSelector(self):
		targetPath = askdirectory(initialdir="/",title='Pick a directory')
		self.targetPath.set(targetPath)

gui()