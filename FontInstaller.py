import os
import os.path
import shutil
import zipfile
import re
import tkinter.messagebox
from tkinter import *
from tkinter.filedialog import *

class gui:
	"""the gui of FontInstaller"""
	def __init__(self):
		window = Tk()
		# title
		window.title('Font Installer '+version)
		# size
		window.geometry('450x280')
		window.resizable(width = FALSE, height = TRUE)
		# icon
		window.iconbitmap('FontInstaller.ico')
		
		# menu bar
		menubar = Menu(window)
		window.config(menu = menubar)
		aboutMenu = Menu(menubar, tearoff = 0)
		# extract
		menubar.add_cascade(label = "Extract", command = self.extract)
		# install
		menubar.add_cascade(label = "Install", command = self.install)
		# pick
		menubar.add_cascade(label = "Pick", command = self.pickup)
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
		self.entryFont = Entry(frameFont, textvariable = self.fontPath)
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
		self.targetPath.set('Unselected')
		self.entryTarget = Entry(frameTarget, textvariable = self.targetPath)
		btTarget = Button(frameTarget, text = "Browse", command = self.targetPathSelector)
		# widget grid
		labelTarget.grid(row = 1, column = 0, sticky = W, padx = 3)
		self.entryTarget.grid(row = 1, column = 1, sticky = W, padx = 3, ipadx = 50)
		btTarget.grid(row = 1, column = 2, pady = 5, sticky = W)


		# Control
		frameControl = Frame(window)
		frameControl.pack()
		self.judTarget = IntVar()
		self.judTarget.set(1)
		cbTarget = Checkbutton(frameControl, text = "Extract to target folder", variable = self.judTarget)
		self.judDelete = IntVar()
		self.judDelete.set(1)
		cbDelete = Checkbutton(frameControl, text = "Delete other files", variable = self.judDelete)
		self.judRemove = IntVar()
		self.judRemove.set(0)
		cbRemove = Checkbutton(frameControl, text = "Remove original font files", variable = self.judRemove)
		# widget grid
		cbTarget.grid(row = 0, column = 0, sticky = W)
		cbDelete.grid(row = 0, column = 1, padx  = 10, sticky = W)
		cbRemove.grid(row = 1, sticky = W)

		# infomation
		frameInfo = Frame(window)
		frameInfo.pack()
		scrollbar = Scrollbar(frameInfo)
		scrollbar.pack(side = RIGHT, fill = Y)
		self.info = Text(frameInfo, width = 150, height = 60, yscrollcommand = scrollbar.set)
		self.info.pack()
		self.info.config(font = 'Courier 10')
		self.info.tag_config('blue', foreground = 'blue')
		self.info.insert(END, '-'*50+'\n')
		self.info.insert(END, '!Pleae backup before use!\nInstall:\nSelect font root path.\nExtract or Pick:\nSelect font root path as well as target path.\n', 'blue')
		self.info.insert(END, '-'*50+'\n')
		scrollbar.config(command = self.info.yview)

		window.mainloop()

	def extract(self):
		"""Extract fonts from fontPath"""
		
		self.couterExtracted = 0
		# judge fontpath
		fontPath = self.fontPath.get()
		if not os.path.exists(fontPath):
			self.errorFontpath()
			return
		# judge targetpath
		targetPath = self.targetPath.get()
		if not os.path.exists(targetPath):
			os.makedirs(targetPath)
		# extract
		for root, dirs, files in os.walk(fontPath):
			for fn in files:
				sufix = os.path.splitext(fn)[1][1:]
				if sufix == 'zip' or sufix == 'ZIP':
					f = zipfile.ZipFile(root+'/'+fn)
					if self.judTarget.get():
						f.extractall(targetPath)
					else:
						f.extractall(root)
					self.info.insert(END, 'Extracted: '+root+'/'+fn+'\n')
					self.info.update()
					self.info.yview(END)
					self.couterExtracted += 1
		self.info.insert(END, '-'*50+'\n')
		self.info.insert(END, 'Extracted'+' '+str(self.couterExtracted)+' Zips.'+'\n')
		self.info.update()
		self.info.yview(END)

	def install(self):
		"""Install fonts from fontPath"""

		couterInstalled = 0
		couterDeleted = 0
		# judge fontpath
		fontPath = self.fontPath.get()
		if not os.path.exists(fontPath):
			self.errorFontpath()
			return
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
					self.info.insert(END, 'Installed: '+root+'/'+fn+'\n')
					self.info.update()
					self.info.yview(END)
					couterInstalled += 1
				elif sufix == 'zip' or sufix == 'ZIP' or sufix == 'rar' or sufix == 'RAR':
					pass
				elif self.judDelete.get():
					os.remove(root+'\\'+fn)
					self.info.insert(END, 'Deleted: '+root+'/'+fn+'\n')
					self.info.update()
					self.info.yview(END)
					couterDeleted += 1
		self.info.insert(END, '-'*50+'\n')
		self.info.insert(END, 'Installed'+' '+str(couterInstalled)+' Fonts.'+'\n')
		self.info.insert(END, 'Deleted'+' '+str(couterDeleted)+' Files.'+'\n')
		self.info.update()
		self.info.yview(END)

	def pickup(self):
		"""pick up fonts from fontPath"""

		couterPicked = 0
		# judge fontpath
		fontPath = self.fontPath.get()
		if not os.path.exists(fontPath):
			self.errorFontpath()
			return
		# judge targetpath
		targetPath = self.targetPath.get()
		if not os.path.exists(targetPath):
			os.makedirs(targetPath)
		# pick up
		for root, dirs, files in os.walk(fontPath):
				for fn in files:
					sufix = os.path.splitext(fn)[1][1:]
					if sufix == 'ttf' or sufix == 'TTF' or sufix == 'otf' or sufix == 'OTF':
						if self.judRemove.get():
							shutil.copy(root+'/'+fn, targetPath)
							os.remove(root+'/'+fn)
							self.info.insert(END, 'Picked and Deleted: '+root+'/'+fn+'\n')
						else:
							shutil.copy(root+'/'+fn, targetPath)
							self.info.insert(END, 'Picked: '+root+'/'+fn+'\n')
						self.info.update()
						self.info.yview(END)
						couterPicked += 1
		self.info.insert(END, '-'*50+'\n')
		self.info.insert(END, 'Picked '+str(couterPicked)+' Fonts.'+'\n')
		self.info.update()
		self.info.yview(END)

	def dialogAbout(self):
		about = Tk()
		about.resizable(width = FALSE, height = FALSE)
		about.title('About')
		about.iconbitmap('FontInstaller.ico')
		Label(about, text = '\n\nFontInstaller ver '+version+'\nA simple tool for install some fonts from selected folder.\n\nThanks for SkyWind about the usage of vbs.\nCoding by Dwayne @2014\nIf you find any bugs please contact me:\nDwayne@loliplus.com\n\n').pack(ipadx = 20)

	def dialogHelp(self):
		help = Tk()
		help.resizable(width = FALSE, height = FALSE)
		help.title('Help')
		help.iconbitmap('FontInstaller.ico')
		Label(help, text = '\n\nHere are three function in this FontInstaller.\n').grid(ipadx = 20, sticky = W)
		Label(help, text = '1.Install .ttf/.otf fonts from selected folder.').grid(ipadx = 20, sticky = W)
		Label(help, text = '2.Extract font files which in the zipfile to origin/target folder.').grid(ipadx = 20, sticky = W)
		Label(help, text = '3.Pick up font files from font folder to target folder.\n').grid(ipadx = 20, sticky = W)
		Label(help, text = 'Install only needs FontPath while both Extract and Pick need TargetPath as well as FontPath.').grid(ipadx = 20, sticky = W)
		Label(help, text = 'Backup before use to avoid unexpect error.\n\n', fg = 'blue').grid(ipadx = 20, sticky = W)

	def popup(self, event):
		"""Popup menu"""

		self.popupMenu.post(event.x_root, event.y_root)

	def fontPathSelect(self):
		"""Select fontPath"""

		fontPath = askdirectory(initialdir="/",title='Pick a directory')
		self.fontPath.set(fontPath)

	def targetPathSelector(self):
		"""Select targetPath"""

		targetPath = askdirectory(initialdir="/",title='Pick a directory')
		self.targetPath.set(targetPath)

	def errorFontpath(self):
		"""fontpath incorrect"""

		tkinter.messagebox.showerror("Error", "Font root path incorrect!")

	def errorTargetpath(self):
		"""targetpath incorrect"""

		tkinter.messagebox.showerror("Error", "Target root path incorrect!")

version = str(1.20)
if __name__ == '__main__':
	gui()