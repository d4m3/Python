import time, psutil, sys
import getpass, pyperclip
import webbrowser, tkinter
import os, sys
import xlrd, openpyxl
import smtplib, itertools
import tkinter as tk
import selenium.webdriver.support.ui as ui
import selenium.webdriver as webdriver
import selenium.webdriver.firefox 

# SentDex for ACW mock up
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import style
style.use('ggplot')
# SentDex

from PIL import Image,ImageDraw
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver import Ie
from tkinter import *
from tkinter import  ttk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.support import ui
from selenium.webdriver.common.keys import Keys
from time import sleep
from openpyxl import load_workbook
from aitertools import cycle


calmFont = 'Times Roman', 25, 'bold'
NORM_FONT = ("Verdana", 12)

class App(tk.Frame):
	def __init__(self, master):
	
        #----- FRAME -----
 		self.frame_header = ttk.Frame(master)
 		self.frame_header_new = ttk.Frame(master)

 		self.frame_header_new.grid()
 		self.frame_header.grid()

 		#----- LOGO -----
 		#self.logo1 = PhotoImage(file = '<<PATH>>').subsample(2,2)
 		self.logo1 = PhotoImage(file = '<PATH>').subsample(2,2) #2,2
 		self.logo2 = PhotoImage(file = '<PATH>').subsample(1,1)

 		# Image #006A85
 		ttk.Label(master, image = self.logo1, background = '#f20404', compound = CENTER).grid(row=0, column=1, columnspan=4, sticky='NSEW',ipadx=5,ipady=5)
 		ttk.Label(master, image = self.logo2, background = '#f20404', compound = CENTER).grid(row=0, column=2, columnspan=4, sticky='N',ipadx=20,ipady=20)
 		# Beta
 		#ttk.Label(master, text = 'beta', wraplength = 300, background='#006A85', font=('Times Roman', 14, 'italic')).grid(row=7,column=0, sticky='W', padx=5,pady=5)
 		
 		# Applications
 		'''ttk.Label(master, wraplength = 300, text = 'Applications', background = '#006A85', font=('Courier',15,'bold')).grid(row=0,column=6,padx=5)'''
 		
 		# Message
 		'''ttk.Label(self.frame_header_new, wraplength =300,text = 'There is no gurantee of functionality, this App is provided as is. Any improvments or suggestions please submit a request.',
 						background = '#006A85',font=('Times Romans',10,'italic')).grid(row=1, column=6)'''

 		# Woosah Label
 		'''ttk.Label(master, image= self.logo2, wraplength=300, 
 				 font = (calmFont), text="TEXT",background='#006A85').grid(row=8, column=0, ipadx=5,ipady=5)#columnspan=4,'''

 		# WIP: CREATE AN UPDATE BUTTON FOR PASSWORD CHANGES
 		#----- LOGO -----
 		

 		#----- MENU ITEMS -----

 		master.option_add('*tearoff',False)
 		menubar = Menu(master)
 		master.config(menu = menubar)
 		
 		file = Menu(menubar)
 		edit = Menu(menubar)
 		help_ = Menu(menubar)
 		submit = Menu(menubar)
 		acw_report = Menu(menubar)

 		# Button position layout
 		menubar.add_cascade(menu = file, label = 'File')
 		menubar.add_cascade(menu = edit, label = 'Edit')
 		menubar.add_cascade(menu = submit, label = 'Report')
 		menubar.add_cascade(menu = acw_report, label = 'ACW')
 		menubar.add_cascade(menu = help_, label = 'Help') 
 		
 		# FILE
 		file.add_command(label = 'New', command = lambda: messagebox.showinfo(title='New',
 							message = '(WIP) New Function'))
 		file.add_command(label = 'Save', command = lambda: messagebox.showinfo(title='Save',
 							message = '(WIP) Save Function'))

 		# HELP
 		help_.add_command(label = 'Contact', command = lambda: messagebox.showinfo(title = 'Info',
							message = 'To report any issues with this application.\n'
							'Please submit a form to Damian Forbes\n' 
							'Using Submit feature'))
 		help_.add_command(label = 'About', command = lambda: messagebox.showinfo(title = '<Author Deatils>',
							message = 'Basic program to launch Kaleida\'s applications'))
 		# REPORT
 		submit.add_command(label = 'Submit Issue', command = self.submitForm)

 		# ACW
 		acw_report.add_command(label = 'ACW Report', command = self.acwReport)
	
		#----- MENU ITEMS -----


 		self.createAppButtons(master)

	# CREATE BUTTON LAYOUT
	def createAppButtons(self, master):
		#----- BUTTON LAYOUT -----
		# Grid Layout for Buttons
		#			 COLUMNS 
		# +---------+---------+---------+---------+
		# | (0 , 0) | (0 , 1) | (0 , 2) | (0 , 3) | R
		# | (1 , 0) | (1 , 1) | (1 , 2) | (1 , 3) | O
		# | (2 , 0) | (2 , 1) | (2 , 2) | (2 , 3) | W
		# | (3 , 0) | (3 , 1) | (3 , 2) | (3 , 3) | S
 		# +---------+---------+---------+---------+

 		# ONESIGN
 		onesign = ttk.Button(master, text="OneSign", compound = CENTER, command=self.launchOneSign)
 		onesign.grid(row=1, column=2,padx=5, pady=5,sticky = 'NSEW') # W
 		# POINTCLICKCARE
 		pointclickcare = ttk.Button(master, text="Point Click Care", compound = CENTER, command=self.launchpcc)
 		pointclickcare.grid(row=1, column=3,padx=5, pady=5,sticky = 'NSEW') #E
 		# LAWSON LOOKUP
 		lawsonlookup = ttk.Button(master, text="Lawson Lookup", compound = CENTER, command=self.launchKpass)
 		lawsonlookup.grid(row=1, column=1,padx=5, pady=5,sticky = 'NSEW')#W
 		# CLEARTRAN
 		cleartran = ttk.Button(master, text="ClearTran", compound = CENTER, command=self.launchClearTran)
 		cleartran.grid(row=2, column=1, padx=5, pady=5,sticky = 'NSEW')  	#W	
  		# IMPRIVATA
 		cortext = ttk.Button(master, text="Imprivata", compound=CENTER, command=self.imprivataCortext)		
 		cortext.grid(row=2,column=2,padx=5,pady=5, sticky='NSEW')#W
 		# ANAKAM
 		anakam = ttk.Button(master, text="Anakam", compound=CENTER, command=self.launchAnakam)
 		anakam.grid(row=2,column=3,padx=5,pady=5,sticky='NSEW') #E
 		# STARS
 		stars = ttk.Button(master,text="Stars", compound = CENTER, command=self.launchStars)
 		stars.grid(row=3, column=1,padx=5,pady=5,sticky='NSEW')
 		# PLACEHOLDERS
 		pH1 = ttk.Button(master,text="PlaceHolder1", compound = CENTER, command=self.placeHolder_One)
 		pH1.grid(row=3, column=2,padx=5,pady=5,sticky='NSEW')
 		pH2 = ttk.Button(master,text="PlaceHolder2", compound = CENTER, command=self.placeHolder_Two)
 		pH2.grid(row=3, column=3,padx=5,pady=5,sticky='NSEW')
 		#QUIT
 		ttk.Button(master, text="QUIT", compound = CENTER,
				   command=master.destroy).grid(row=7,column=1,padx=5, pady=5, sticky='W' )#columnspan=3 W

  	 
	# CLEARTRAN
	def launchClearTran(self):
		# Not AD link -- Page Timeout after x-mins so logout pref
		# Ask for username(NAT Account)
		# Create an entry window for value 
		# TODO: CREATE ENTRY WINDOW
		searchApp = Tk()
		searchApp.title("ClearTran Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="User\'s AD Account: ").grid(row=1)
		Label(searchApp, text="Admin Password: ").grid(row=0)

		v = StringVar()
		x = StringVar() 

		entryBox = Entry(searchApp, textvariable = v)
		pswdBox = Entry(searchApp, textvariable = x, show="*")
		
		entryBox.grid(row=1, column=1, padx=5,pady=10)
		pswdBox.grid(row=0, column =1, padx =5, pady =10)

		entryBox.focus_set()
		pswdBox.focus_set()

		Button(searchApp, text="Search",command=lambda:callback()).grid(row=2,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=2,column=2, padx=5,pady=5)

		# Search input name
		def callback():
			search = entryBox.get()
			passwordBox = pswdBox.get()
			
			## ClearTran Page Behaviour
			ie_driver = webdriver.Ie("<PATH>")
			#ie_driver = webdriver.Ie(<PATH>)
			
			# Minimize IEDriverServer.exe
			#webdriverIe.manage().window().set_window_position(-2000,0)

			ie_driver.get('<PATH>')
			wait = ui.WebDriverWait(ie_driver,20)

			username_field = ie_driver.find_element_by_id('Username')
			usr = os.getlogin()
			username_field.send_keys(usr)
			password_field = ie_driver.find_element_by_id('Password')
			password_field.send_keys(passwordBox)
			password_field.send_keys(Keys.RETURN)

			wait = ui.WebDriverWait(ie_driver, 10)
			
			time.sleep(2)
			maint = ie_driver.find_element_by_id("Menubar_tdMaint").click()
			time.sleep(3)
			users = ie_driver.find_element_by_xpath("//*[@id='_ctl5_tblMaintenance']/tbody/tr/td/a[2]").click()

			# Paste in searchfield
			time.sleep(3)
			findUser = ie_driver.find_element_by_id("user_txtUserName") 
			findUser.send_keys(search)
			submit = ie_driver.find_element_by_id("user_btnFindUser").click()
			time.sleep(1)
			# Found user 
			usrName = ie_driver.find_element_by_xpath("//*[@id='user_grdUsers']/tbody/tr[2]/td[1]/a").click()
			
			# CHANGE PASSWORD
			changePswd = ie_driver.find_element_by_xpath("//*[@id='user_txtPassword']")

			# SELECT ALL, for current password and delete
			changePswd.send_keys(Keys.CONTROL,'a')
			changePswd.send_keys(Keys.DELETE)
			changePswd.send_keys("Password1")

			confirmPswd = ie_driver.find_element_by_xpath("//*[@id='user_txtPasswordConfirm']")
			confirmPswd.send_keys(Keys.CONTROL,'a')
			confirmPswd.send_keys(Keys.DELETE)
			confirmPswd.send_keys("Password1")

			# CHECK LOCK-OUT STATE
			# http://stackoverflow.com/questions/14442636/how-can-i-check-if-a-checkbox-is-checked-in-selenium-python-webdriver
			locked_out = ie_driver.find_element_by_xpath("//*[@id='user_chkIsLocked']")
			if locked_out.is_selected():
			# If selected, uncheck the check box	
				locked_out.click()
				update = ie_driver.find_element_by_xpath("//*[@id='user_btnupdate']").click()
			else:
				update = ie_driver.find_element_by_xpath("//*[@id='user_btnupdate']").click()

			#update = ie_driver.find_element_by_xpath("//*[@id='user_btnupdate']").click()

			# CREATE A POPUP STATING PASSWORD WAS CHANGED
			messagebox.showinfo(title = "Password Changed", message = 'Password has been changed to:\n\tPassword1')

			# Sign out of ClearTran
			signOff = ie_driver.find_element_by_xpath("//*[@id='Menubar_btnLogoff']").click()
			time.sleep(3)
			ie_driver.close()

			PROCNAME = "IEDriverServer.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()
			## TEST

		# Clear Entry box
		def clear_entryBox():
			entryBox.delete(0,'end')

		mainloop()
	# DIRECTOR COMPATABILITY ISSUES
	def launchDirector(self):
		"""ie_driver = webdriver.Ie("<PATH>")
		#ie_driver = webdriver.Ie("<PATH>")
		ie_driver.get('<PATH>')
		wait = ui.WebDriverWait(ie_driver, 20)

		username_field = ie_driver.find_element_by_id('UserName')
		usr = os.getlogin()
		username_field.send_keys(usr)

		password_field = ie_driver.find_element_by_id('Password')
		password_field.send_keys('*********')
		password_field.send_keys(Keys.TAB)

		domain_field = ie_driver.find_element_by_id('Domain')
		domain_field.send_keys('*******')
		domain_field.send_keys(Keys.RETURN)

		PROCNAME = "IEDriverServer.exe"
		for proc in psutil.process_iter():
			if proc.name() == PROCNAME:
				proc.kill()
		"""
		pass				
	# PCC
	def launchpcc(self):
		searchApp = Tk()
		searchApp.title("PCC Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="Admin\'s Password: ").grid(row=0)
		Label(searchApp, text="AD Account: ").grid(row=1)
				
		v = StringVar()
		x= StringVar()
		
		entryBox = Entry(searchApp, textvariable = v)
		passwordBox = Entry(searchApp, textvariable = x, show="*")

		passwordBox.grid(row=0,column=1, padx=5,pady=5)
		entryBox.grid(row=1,column=1, padx=5,pady=5)
		

		entryBox.focus_set()
		passwordBox.focus_set()

		#Button(searchApp, text="Quit",command=searchApp.quit).grid(row=1,column=0,padx=5)
		Button(searchApp, text="Search",command=lambda:callback()).grid(row=2,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=2,column=2, padx=5,pady=5)

		def callback():
			search = entryBox.get()
			pswdBox = passwordBox.get()

			portal = '<PATH>'
			chromeDriver = "<PATH>"
			#chromeDriver = '<PATH>'
			
			os.environ["webdriver.chrome.driver"] = chromeDriver
			ie_driver = webdriver.Chrome(chromeDriver)

			# Set browser window to see popup
			ie_driver.set_window_size(800,800)
			ie_driver.set_window_position(250,200)

			ie_driver.get(portal)
			wait = ui.WebDriverWait(ie_driver, 5)

			username_field = ie_driver.find_element_by_id('username')
			usr = os.getlogin()
			username_field.send_keys('klh.'+usr) # 'klh.' needed if cache deletes

			password_field = ie_driver.find_element_by_id('password')

			password_field.send_keys(pswdBox)
			password_field.send_keys(Keys.RETURN)

			time.sleep(2)
			std = ie_driver.find_element_by_xpath("//*[@id='QTF_emcStandardsTab']/a").click()
			time.sleep(1)
			sec_Users = ie_driver.find_element_by_xpath("//*[@id='filtersearch']")
			time.sleep(1)
			sec_Users.send_keys("sec")
			time.sleep(1)
			sec_Users_click = ie_driver.find_element_by_xpath("//*[@id='listContent']/div[13]/ul/li[1]/a").click()
			time.sleep(1)

			chkBox = ie_driver.find_element_by_id('id-accessTypeSingle').click()
			time.sleep(3)
			search_user = ie_driver.find_element_by_xpath("//*[@id='fullscreen']/form/table[2]/tbody/tr[3]/td/input[2]")

			search_user.send_keys(search)
			search_user.send_keys(Keys.RETURN) # commented as error was prompting

			#time.sleep(2)

			# Just click on link for pwd-b/c the element has db info
			reset_pwd_link = ie_driver.find_element_by_link_text('pwd').click()

			#click_pswd = ie_driver.find_element_by_xpath("//*[@id='userrow1407']/td[1]/a[5]").click()
			time.sleep(2)
			
			# Switch to POPUP WINDOW: http://stackoverflow.com/questions/10629815/handle-multiple-window-in-python
			# Get Window Handle Before
			window_handle_before = ie_driver.window_handles[0]

			# START - POPUP WINDOW
			# Get Window Handle After
			window_handle_after = ie_driver.window_handles[1] #screenpopup
			#print("Windows_Handle_After: ",window_handle_after)
			ie_driver.switch_to.window(window_handle_after)
			new_pswd = ie_driver.find_element_by_xpath("//*[@id='idESOLpasswordNew']")
			new_pswd.send_keys("Password1")
			confirm_pswd = ie_driver.find_element_by_xpath("//*[@id='idESOLpasswordChk']")
			confirm_pswd.send_keys("Password1")
			next_login = ie_driver.find_element_by_xpath("//*[@id='detail']/form/table/tbody/tr[4]/td/table/tbody/tr[4]/td[1]/input").click()
			time.sleep(2)
			save_button = ie_driver.find_element_by_xpath("//*[@id='sButton']").click()
						
			# Message about Password Reset
			messagebox.showinfo(title = 'PCC', message = 'Password has been changed to:\n\tPassword1')	
			ie_driver.set_window_size(100,100)
			time.sleep(5)

		def clear_entryBox():
			entryBox.delete(0,'end')

		mainloop()

		PROCNAME = "chromedriver.exe"
		for proc in psutil.process_iter():
		 	if proc.name() == PROCNAME:
		 		proc.kill()
	# KPASS
	def launchKpass(self):
		searchApp = Tk()
		searchApp.title("Lawson Lookup Portal")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text = 'Admin Password: ').grid(row=0)
		Label(searchApp, text = 'User\'s Lastname: ').grid(row=1)
		
		v = StringVar()
		x = StringVar()

		entryBox = Entry(searchApp, textvariable = v)
		passwordBox = Entry(searchApp, textvariable = x, show="*")

		passwordBox.grid(row=0,column=1,padx=5,pady=10)
		entryBox.grid(row=1,column=1,padx=5,pady=10)
		
		passwordBox.focus_set()
		entryBox.focus_set()

		Button(searchApp, text = 'Search', command=lambda:callback()).grid(row=2,column=1,padx=5,pady=5)#.send_keys(Keys.RETURN)
		Button(searchApp, text = 'Clear', command=lambda:clear_entryBox()).grid(row=2,column=2,padx=5,pady=5)
		
		def callback():
			search = entryBox.get()
			pswdBox = passwordBox.get()

			ie_driver = webdriver.Ie("<PATH>")
			ie_driver.get('<PATH>')
			wait = ui.WebDriverWait(ie_driver, 15)

			# Accept Certificate
			ie_driver.get("javascript:document.getElementById('overridelink').click()")
			time.sleep(1)
			usr = os.getlogin()
			username_field = ie_driver.find_element_by_xpath("//*[@id='Account_20Name']/table/tbody/tr[2]/td/input")
			username_field.send_keys(usr)

			username_field.send_keys(Keys.TAB)
			password_field = ie_driver.find_element_by_xpath("//*[@id='Password']/table/tbody/tr[2]/td/input")
			password_field.send_keys(pswdBox)
			
			password_field.send_keys(Keys.RETURN)
			time.sleep(2)
			
			# Get clipboard data
			lastname_field = ie_driver.find_element_by_xpath("//*[@id='LastName']/table/tbody/tr[2]/td/input")
			lastname_field.send_keys(search)
			lastname_field.send_keys(Keys.RETURN)
			time.sleep(30)
			ie_driver.close()

			PROCNAME = "IEDriverServer.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()

		def clear_entryBox():
			entryBox.delete(0,'end')

		mainloop()
	# IMPRIVATA CORTEXT
	def imprivataCortext(self):
		searchApp = Tk()
		searchApp.title("Imprivata Cortext Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)
		
		ttk.Label(searchApp,text="Admin username: ").grid(row=0)
		ttk.Label(searchApp, text="Admin\'s Password: ").grid(row=1)
		ttk.Label(searchApp, text="User\'s First/Lastname: ").grid(row=2) 

		usr = StringVar
		v = StringVar()
		x = StringVar()

		adminBox = Entry(searchApp, textvariable = usr)
		entryBox = Entry(searchApp, textvariable = v)
		passwordBox = Entry(searchApp, textvariable = x, show="*")


		adminBox.grid(row=0,column=1, padx=5,pady=5)
		passwordBox.grid(row=1,column=1, padx=5,pady=5)
		entryBox.grid(row=2,column=1, padx=5,pady=5)

		adminBox.focus_set()
		passwordBox.focus_set()
		entryBox.focus_set()
		
		Button(searchApp, text="Search",command=lambda:callback()).grid(row=3,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=3,column=2, padx=5,pady=5)


		def callback():
			adminUser = adminBox.get()
			pswdBox = passwordBox.get()
			searchUser = entryBox.get()
			#ie_driver = webdriver.Ie("<PATH>")
			portal = '<PATH>'
			chromeDriver = '<PATH>'
			os.environ["webdriver.chrome.driver"] = chromeDriver
			chrome_driver = webdriver.Chrome(chromeDriver)
			chrome_driver.get(portal)
			wait = ui.WebDriverWait(chrome_driver, 15)
						
			time.sleep(2)
			#username_field = chrome_driver.find_element_by_id("username")
			username_field = chrome_driver.find_element_by_xpath("//input[@id='username']")
			username_field.send_keys(adminUser)
			password_field = chrome_driver.find_element_by_id("password")
			password_field.send_keys(pswdBox)
			time.sleep(1)
			login = chrome_driver.find_element_by_xpath("//input[@id='loginSubmit']").click()
			time.sleep(2)
			# Looking for Users
			# Click on Users to search
			_users = chrome_driver.find_element_by_xpath("//*[@id='tabUsers']/a").click()
			time.sleep(2)
			
			# Search users
			search = chrome_driver.find_element_by_xpath("//*[@id='searchField']")#.click()
			editLink = chrome_driver.find_element_by_xpath("//tbody/tr/td[6]/div/div[2]/ul/li[2]/a")

			time.sleep(1)
			search.send_keys(searchUser)
			search.send_keys(Keys.RETURN)
			time.sleep(5)
			search.clear()
			#searchUser.delete(0,'end')
			# searchApp = Tk()
			# searchApp.title("Cortext Imprivata")
			# searchApp.geometry("400x125")
			# searchApp.resizable(False,False)

			# Label(searchApp, text = 'Username: ').grid(row=0)
			# Label(searchApp, text = '(hint:%) Admin Password: ').grid(row=1)

			# v = StringVar()
			# x = StringVar()

			
			# entryBox = Entry(searchApp, textvariable = v)
			# pswdBox = Entry(searchApp, textvariable = x, show="*")

			# entryBox.grid(row=0,column=1,padx=5,pady=10)
			# pswdBox.grid(row=1,column=1,padx=5,pady=10)
			
			# entryBox.focus_set()
			# pswdBox.focus_set()

			# Button(searchApp, text = 'Search', command=lambda:callback()).grid(row=1,column=1,padx=5)#.send_keys(Keys.RETURN)
			# Button(searchApp, text = 'Clear', command=lambda:clear_entryBox()).grid(row=1,column=2,padx=5)

			# http://stackoverflow.com/questions/7732125/clear-text-from-textarea-with-selenium
		# def callback():
		# 	search.send_keys(entryBox.get())
		# 	search.send_keys(Keys.RETURN)
		# 	time.sleep(5)

			# WORKING - Remove name from search box in webpage and entry box
			search.clear()
			entryBox.delete(0,'end')

		def clear_entryBox():
			print("clear_entryBox")
			search.clear()
			entryBox.delete(0,'end')
		mainloop()
		PROCNAME = "chromedriver.exe"
		for proc in psutil.process_iter():
			if proc.name() == PROCNAME:
				proc.kill()
	# ANAKAM
	def launchAnakam(self):
		searchApp = Tk()
		searchApp.title("Anakam AdminPro")
		searchApp.geometry("400x150")
		searchApp.resizable(False,False)

		Label(searchApp, text="Admin Username: ").grid(row=0)
		Label(searchApp, text="Admin Password: ").grid(row=1)
		Label(searchApp, text="AD Account: ").grid(row=2)

		# LOGO
		# bg_img = PhotoImage(file = '<PATH>').subsample(1,1)
		# bg_img = Label(searchApp, image=bg_img)
		# bg_img.pack()
		# # LOGO

		v = StringVar()
		x = StringVar() 
		usr = StringVar()

		entryBox = Entry(searchApp, textvariable = v)
		pswdBox = Entry(searchApp, textvariable = x, show="*")
		usr_entry = Entry(searchApp, textvariable = usr)

		
		entryBox.grid(row=0, column=1, padx=5,pady=10)
		pswdBox.grid(row=1, column =1, padx =5, pady =10)
		usr_entry.grid(row=2, column =1, padx =5, pady =10)

		entryBox.focus_set()
		pswdBox.focus_set()
		usr_entry.focus_set()

		Button(searchApp, text="Search",command=lambda:callback()).grid(row=3,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=3,column=2, padx=5,pady=5)

		def callback():
			adminUser = entryBox.get()
			pswd_Box = pswdBox.get()
			searchUser = usr_entry.get()

			portal = '<PATH>'
			chromeDriver = '<PATH>'
			os.environ["webdriver.chrome.driver"] = chromeDriver
			chrome_driver = webdriver.Chrome(chromeDriver)
			chrome_driver.get(portal)
			wait = ui.WebDriverWait(chrome_driver, 15)

			username_field = chrome_driver.find_element_by_id("j_username")
			username_field.send_keys(adminUser)

			password_field = chrome_driver.find_element_by_id("j_password")
			password_field.send_keys(pswd_Box)
			password_field.send_keys(Keys.RETURN)
			time.sleep(1)
			chrome_driver.find_element_by_xpath("//*[@id='menubar']/ul/li[3]/a").click()
			
			# User Key
			userKey = chrome_driver.find_element_by_xpath("//*[@id='nField']")
			userKey.send_keys(searchUser)
			userKey.send_keys(Keys.RETURN)

			# Results and Click AD Account
			foundUser = chrome_driver.find_element_by_xpath("//*[@id='users']/tbody/tr/td[2]/a").click()

			# Clear AD Account
			searchUser.clear()

		def clear_entryBox():
			pass
		# TODO: CREATE A RADIOBUTTON TO CHECK REPORT		

		mainloop()
		PROCNAME = "chromedriver.exe"
		for proc in psutil.process_iter():
			if proc.name() == PROCNAME:
				proc.kill()
	# IMPRIVATA 1SIGN
	def launchOneSign(self):
		searchApp = Tk()
		searchApp.title("Lawson Lookup Portal")
		searchApp.geometry("400x135")
		searchApp.resizable(False,False)

		Label(searchApp, text = '(AD)Admin Username: ').grid(row=0)
		Label(searchApp, text = '(AD)Admin Password: ').grid(row=1)
		Label(searchApp, text = '(AD)Username').grid(row=2)
		
		v = StringVar()
		x = StringVar()
		u = StringVar()

		entryBox = Entry(searchApp, textvariable = v)
		passwordBox = Entry(searchApp, textvariable = x, show="*")
		usr = Entry(searchApp, textvariable = u)

		entryBox.grid(row=0,column=1,padx=5,pady=5)
		passwordBox.grid(row=1,column=1,padx=5,pady=5)
		usr.grid(row=2,column=1,padx=5,pady=5)
		
		passwordBox.focus_set()
		entryBox.focus_set()

		Button(searchApp, text = 'Search', command=lambda:callback()).grid(row=3,column=1,padx=5,pady=5)#.send_keys(Keys.RETURN)
		Button(searchApp, text = 'Clear', command=lambda:clear_entryBox()).grid(row=3,column=2,padx=5,pady=5)

		def callback():
			adminEntry = entryBox.get()
			pswd_Box = passwordBox.get()
			user = usr.get()

			ie_driver = webdriver.Ie("<PATH>")
			ie_driver.get('<PATH>')
			ie_driver.get("javascript:document.getElementById('overridelink').click()")

			wait = ui.WebDriverWait(ie_driver, 15)
			time.sleep(2)

			document.getElementById("LOGIN").element["userid"]

			#username_field = ie_driver.find_element_by_id('LOGIN')
			#username_field.send_keys(adminEntry)
			

		def clear_entryBox():
			pass
	# STARS
	def launchStars(self):	
		# TODO: ask for username/admin username/admin pswd
		searchApp = Tk()
		searchApp.title("STARS Event Review")
		searchApp.geometry("400x175")
		searchApp.resizable(False,False)

		# Label(searchApp, text="Admin Username: ").grid(row=0) --- Use AD account
		Label(searchApp, text="(hint:&)Admin Password: ").grid(row=0)
		Label(searchApp, text="Client Name: ").grid(row=1)

		# adminUsername = StringVar()
		adminPassword = StringVar()
		clientUser = StringVar()

		# adminEntryBox = Entry(searchApp, textvariable = adminUsername)
		pwdEntryBox = Entry(searchApp, textvariable = adminPassword, show ="*")
		clientEntryBox = Entry(searchApp, textvariable = clientUser)

		pwdEntryBox.grid(row=0, column =1, padx=5,pady=5)
		clientEntryBox.grid(row=1, column =1, padx=5,pady=5)

		# adminEntryBox.focus_set()
		pwdEntryBox.focus_set()
		clientEntryBox.focus_set()

		Button(searchApp, text="Search", command=lambda:callback()).grid(row=2,column=1,padx=5,pady=5)
		Button(searchApp, text="Clear", command=lambda:clear()).grid(row=2,column=2,padx=5,pady=5)

		def callback():
			defaultID = 'K525'
			loginAdmin = os.getlogin()
			loginPwd = pwdEntryBox.get()
			searcClient = clientEntryBox.get()

			# TODO: login to STARS [application tab > general], if KaleidaScope login not needed got directly
			portal = '<PATH>'
			# TODO: using IE (for now until test in Chrome/FF)
			ie_driver = webdriver.Ie("<PATH>")
			#ie_driver = webdriver.Ie("<PATH>")
			ie_driver.get(portal)
			delay = 3
			
			# K525
			k525 = ie_driver.find_element_by_xpath("//*[@id='txtClientID']")
			k525.send_keys(defaultID)

			# AD Account
			admin = ie_driver.find_element_by_xpath("//*[@id='txtUserID']")
			admin.send_keys(loginAdmin)

			# PASSWORD
			pwd = ie_driver.find_element_by_xpath("//*[@id='txtPassword']")
			pwd.send_keys(loginPwd)
			pwd.send_keys(Keys.RETURN)

			# Wait for page to load/use JavaScript
			time.sleep(45)
			print('45s passed')
			time.sleep(2)
			# try:
			# 	WebDriverWait(ie_driver, delay).until(EC.presence_of_element_located(ie_driver.find_element_by_link_text('ADMIN').click()))
			# except TimeoutException:
			# 	print("Page took too long to load!")

			# JavaScript test
			#ie_driver("javascript:document.getElementById('Admin').click()")

			#ie_driver.find_element_by_link_text("ADMIN").click()
			#ie_driver.find_element_by_id("orionSlCtrl").click()
			
			# Click on Admin
			_clickADMIN = ie_driver.find_element_by_xpath("//*[@type='ADMIN']")
			_clickADMIN.click()

			# Another window opens and loads, wait 10s?
			time.sleep(45)

			# Window before
			window_handle_before = ie_driver.window_handles[0]
			
			# (opened window) Get window handle
			window_handle_after = ie_driver.window_handles[1]

			ie_driver.switch_to.window(window_handle_after)

			# Click on Users
			users = ie_driver.find_element_by_link_text("Users").click()
			
			time.sleep(2)

	# PLACEHOLDER FUNCTION 
	def placeHolder_One(self):
		pass
	# PLACEHOLDER FUNCTION
	def placeHolder_Two(self):
		pass
	# SUBMIT FUNCTION		
	def submitForm(self):
		gui = Tk()
		gui.title('Submit Issue')
		gui.geometry('800x400')
		gui.resizable(True,True)
		gui_label = Label(text = 'Submission Form')

		self.frame_content = ttk.Frame(gui)
		self.frame_content.pack()

		ttk.Label(self.frame_content, text = 'Name:').grid(row=0, column=0, padx=5,pady=5,sticky='sw')
		ttk.Label(self.frame_content, text = 'Email:').grid(row=0, column=1, padx=5,pady=5,sticky='sw')
		ttk.Label(self.frame_content, text = 'Phone Number:').grid(row=0, column=2, padx=5,pady=5,sticky='sw')
		ttk.Label(self.frame_content, text = 'Comments:').grid(row=2, column=0, padx=5,pady=5,sticky='sw')

		self.entry_name = ttk.Entry(self.frame_content, width=24,font= NORM_FONT) # ('Arial', 10)
		self.entry_email = ttk.Entry(self.frame_content, width=24,font= NORM_FONT)
		self.text_comments = Text(self.frame_content, width=50,height=10,font= NORM_FONT)
		self.phone_number = ttk.Entry(self.frame_content, width=24,font= NORM_FONT)

		self.entry_name.grid(row=1, column=0, padx=5)
		self.entry_email.grid(row=1, column=1, padx=5)
		self.phone_number.grid(row=1, column=2, padx=5)
		self.text_comments.grid(row=3,column=0,columnspan=3,padx=5)
	
		ttk.Button(self.frame_content, text = 'Submit',
					command = self.submitButton).grid(row=4, column=0,padx=5,pady=5,sticky='e')
		ttk.Button(self.frame_content, text = 'Clear',
					command = self.clearButton).grid(row=4, column=1,padx=5,pady=5,sticky='w')
		gui.mainloop()
	def submitButton(self):
		nameInfo = self.entry_name.get()
		emaiAddress = self.entry_email.get()
		comment = self.text_comments.get(1.0, 'end')
		phone = self.phone_number.get()

		# SEND EMAIL
		smtpObj = smtplib.SMTP_SSL('smtp.gmail.com', 465)

		#print('\nehlo ---->', smtpObj.ehlo())
		#print('gmail logging in')

		# Dummy Account
		smtpObj.login('<Username>','<Password>')

		#print('*'*10)
		#print('Login completed')
		#print('Sending Email')
		# Forward emails to this account
		smtpObj.sendmail('Username', '<email>@gmail.com','Subject: Issue Submitted\n'
							+'\nName: '+str(nameInfo)+'\nEmail Address: '+str(emaiAddress)+'\nPhone Number:' + str(phone)
							+'\n\nComments:\n '+str(comment))
		#print("Comments sent")
		messagebox.showinfo(title = 'Comments', message = 'Your comments have been submitted.')	
		time.sleep(5)
		self.clearButton()
	def clearButton(self):
		self.entry_name.delete(0, 'end')
		self.entry_email.delete(0, 'end')
		self.phone_number.delete(0, 'end')
		self.text_comments.delete(1.0, 'end')
	
	def popupmsg(msg):
		popup = tk.Tk()
		popup.wm_title('Notice')
		popup.geometry('250x75')
		label = ttk.Label(popup, text = "Window close after 45s", compound = CENTER, font = NORM_FONT)
		label.pack(side='top', fill='x', pady=5)
		button = ttk.Button(popup,text='Ok', command = popup.destroy)
		button.pack()
		popup.mainloop()

	#############################################################################################
	# ANALYST DATA
	def acwData(self):
		radioButton = v.get()
		#file = ('<PATH>')
		file = '<PATH>'
		df = pd.read_excel(file)
		df.to_excel('<PATH>', header=None)
		df = pd.read_excel('<PATH>', index_col=[0,1])

		# ACD Calls - Unnamed:2 |	Avg ACD Time - Unnamed:3 | Avg ACW Time - Unnamed:4 | 
		# % Agent Occupancy w/o ACW - Unnamed:5 | Extn in Calls - Unnamed:6 | Avg Extn In Time - Unnamed:7 
		# Extn Out Calls - Unnamed:8 | Avg Extn Out Time - Unnamed:9 | ACD Time - Unnamed:10 
		# ACW Time - Unnamed:11 | Agent Ring Time: Unnamed:12 | Other Time - Unnamed:13 | 
		# AUX Time - Unnamed:14 Avail Time: Unnamed:15 | Staffed Time - Unnamed:16
		del df['Unnamed: 2']
		del df['Unnamed: 3']
		del df['Unnamed: 5']
		del df['Unnamed: 6']
		del df['Unnamed: 7']
		del df['Unnamed: 8']
		del df['Unnamed: 9']
		del df['Unnamed: 10']
		del df['Unnamed: 11']
		del df['Unnamed: 12']
		del df['Unnamed: 13']
		del df['Unnamed: 14']
		del df['Unnamed: 15']
		del df['Unnamed: 16']
		del df['Unnamed: 17']
		df = df.rename(columns = {'Unnamed: 4':'Avg ACW Time'})
		# Remove Totals (-1)
		df.drop(df.index[0], inplace=True)
		
		for index, row in df.iterrows():
			if radioButton == 1 and 'Lastname, Firstname':
				print('Anaylst: ',index,row['Avg ACW Time'])
			elif (v.get()) == 2 and 'Lastname, Firstname':
				print("Joe")
			else:
				print("No Option")
			break
			
		# Option for choice, input may not work
		# error input() : http://stackoverflow.com/questions/12547683/python-3-eof-when-reading-a-line-sublime-text-2-is-angry
		
		df.columns
		#print(df)
		print("\nRadioButton selected: ", newButton)
		
	#############################################################################################

	def acwReport(self):
		# Select Anaylst
		master = Tk()
		master.title("Select Analyst")
		master.geometry("400x200")
		master.resizable(False,False)
		
		global v
		v = IntVar()
		v.set(1)

		# compound=CENTER, sticky='NSEW',
		ttk.Radiobutton(master, text = 'Lastname, Firstname', command = lambda:self.acwData(), variable = v, value=1).grid(row=0,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname', command = lambda:self.acwData(), variable = v, value=2).grid(row=0,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname', command = lambda:self.acwData(), variable = v, value=3).grid(row=0,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=4).grid(row=1,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=5).grid(row=1,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=6).grid(row=1,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=7).grid(row=2,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=8).grid(row=2,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=9).grid(row=2,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Lastname, Firstname',command = lambda:self.acwData(), variable = v, value=10).grid(row=3,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)	
		ttk.Radiobutton(master, text = 'None', command = lambda:self.acwData(), variable = v, value=11).grid(row=3,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
	
		mainloop()			

def main():
	root = Tk()
	# Window Details
	root.title("TOOLS (beta)")
	root.resizable(True,True)
	root.geometry("400x350")
	root.configure(background = '#f20404') #006A85

	app = App(root)
	root.mainloop()

	
if __name__ == "__main__": main()

	
	
	
	
	
	
	
	
	
	
	
	