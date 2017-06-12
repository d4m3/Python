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
import pyautogui 
import re
from credentials import SENDER_EMAIL, SENDER_PASSWORD # Email account
import js2py
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import style
from selenium.common.exceptions import NoSuchElementException
from PIL import Image,ImageDraw
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options

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
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from twilio.rest import TwilioRestClient # Text Message
from credentials import account_sid, auth_token, my_cell, my_twilio
from selenium.webdriver.common.action_chains import ActionChains
from pywinauto.application import Application
from pywinauto.keyboard import SendKeys



style.use('ggplot')


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
 		#self.logo1 = PhotoImage(file = '<PATH>').subsample(2,2)
 		self.logo1 = PhotoImage(file = '<PATH>').subsample(2,2) #2,2
 		self.logo2 = PhotoImage(file = '<PATH>').subsample(1,1)

 		# Image #006A85 #add8e6 #f20404
 		ttk.Label(master, image = self.logo1, background = '#f20404', compound = CENTER).grid(row=0, column=1, columnspan=4, sticky='NSEW',ipadx=5,ipady=5)
 		ttk.Label(master, image = self.logo2, background = '#f20404', compound = CENTER).grid(row=0, column=2, columnspan=4, sticky='N',ipadx=20,ipady=20)

 		#----- MENU ITEMS -----
 		master.option_add('*tearoff',False)
 		menubar = Menu(master)
 		master.config(menu = menubar)
 		
 		file = Menu(menubar)
 		edit = Menu(menubar)
 		help_ = Menu(menubar)
 		submit = Menu(menubar)
 		acw_report = Menu(menubar)
 		sms = Menu(menubar)

 		# Button position layout
 		menubar.add_cascade(menu = file, label = 'File')
 		menubar.add_cascade(menu = edit, label = 'Edit')
 		menubar.add_cascade(menu = submit, label = 'Feedback')
 		menubar.add_cascade(menu = acw_report, label = 'ACW')
 		menubar.add_cascade(menu = sms, label='SMS')
 		menubar.add_cascade(menu = help_, label = 'Help') 
 		
 		# FILE
 		file.add_command(label = 'New', command = lambda: messagebox.showinfo(title='New',
 							message = '(WIP) New Function'))
 		file.add_command(label = 'Save', command = lambda: messagebox.showinfo(title='Save',
 							message = '(WIP) Save Function'))

 		# HELP
 		help_.add_command(label = 'Contact', command = lambda: messagebox.showinfo(title = 'Info',
							message = 'To report any issues/suggestions with this application.\n'
							'Please submit a form to Damian Forbes\n' 
							'using Feedback feature.'))
 		help_.add_command(label = 'About', command = lambda: messagebox.showinfo(title = '<Author Deatils>',
							message = 'Basic program to launch Tool\'s applications.\nThis program assumes you\'re handling one user at a time!'))
 		# REPORT
 		submit.add_command(label = 'Submit Feedback', command = self.submitForm)

 		# ACW
 		acw_report.add_command(label = 'ACW Report', command = self.acwReport)

 		# SMS - using TWILIO
 		sms.add_command(label = 'SMS with Twilio', command = self.launchTextMsg)

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
 		cortext = ttk.Button(master, text="Imprivata Cortext", compound=CENTER, command=self.imprivataCortext)		
 		cortext.grid(row=2,column=2,padx=5,pady=5, sticky='NSEW')#W
 		# ANAKAM
 		anakam = ttk.Button(master, text="Anakam", compound=CENTER, command=self.launchAnakam)
 		anakam.grid(row=2,column=3,padx=5,pady=5,sticky='NSEW') #E
 		# STARS
 		stars = ttk.Button(master,text="Stars", compound = CENTER, command=self.launchStars)
 		stars.grid(row=3, column=1,padx=5,pady=5,sticky='NSEW')
 		# INTERQUAL
 		interqual = ttk.Button(master,text="Interqual", compound = CENTER, command=self.launchInterqual)
 		interqual.grid(row=3, column=2,padx=5,pady=5,sticky='NSEW')
 		# TIMELESS
 		timeless = ttk.Button(master,text="Timeless Breastmilk", compound = CENTER, command=self.launchTimeless)
 		timeless.grid(row=3, column=3,padx=5,pady=5,sticky='NSEW')
 		# INFOCLIQUE
 		infoclique = ttk.Button(master,text="Infoclique", compound = CENTER, command=self.launchInfoclique)
 		infoclique.grid(row=4, column=1,padx=5,pady=5,sticky='NSEW')

 		# Placeholders
 		ph2 = ttk.Button(master,text="Placeholder 2", compound = CENTER, command=self.ph2)
 		ph2.grid(row=4, column=2,padx=5,pady=5,sticky='NSEW')

 		ph3 = ttk.Button(master,text="Placeholder 3", compound = CENTER, command=self.launchTimeless)
 		ph3.grid(row=4, column=3,padx=5,pady=5,sticky='NSEW')

 		ph4 = ttk.Button(master,text="Placeholder 4", compound = CENTER, command=self.launchTimeless)
 		ph4.grid(row=5, column=1,padx=5,pady=5,sticky='NSEW')
 	
 		ph5 = ttk.Button(master,text="Placeholder 5", compound = CENTER, command=self.launchTimeless)
 		ph5.grid(row=5, column=2,padx=5,pady=5,sticky='NSEW')

 		ph6 = ttk.Button(master,text="Placeholder 6", compound = CENTER, command=self.launchTimeless)
 		ph6.grid(row=5, column=3,padx=5,pady=5,sticky='NSEW')

 		#QUIT
 		ttk.Button(master, text="QUIT", compound = CENTER,
				   command=master.destroy).grid(row=7,column=1,columnspan=4,padx=5, pady=5)#columnspan=3 W
 	# CREATE BUTTON LAYOUT
  	 
	# CLEARTRAN
	def launchClearTran(self):
		searchApp = Tk()
		searchApp.title("ClearTran Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="User\'s AD Account: ").grid(row=1)
		Label(searchApp, text="(hint:&) Admin Password: ").grid(row=0)

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

			# CREATE A POPUP STATING PASSWORD WAS CHANGED
			messagebox.showinfo(title = "ClearTran Password Changed", message = 'Password has been changed to:\n\tPassword1')

			# Sign out of ClearTran
			signOff = ie_driver.find_element_by_xpath("//*[@id='Menubar_btnLogoff']").click()
			time.sleep(3)
			ie_driver.close()

			PROCNAME = "IEDriverServer.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()
		# Clear Entry box
		def clear_entryBox():
			entryBox.delete(0,'end')
		searchApp.mainloop()
		#mainloop()
	
	# DIRECTOR COMPATABILITY ISSUES
	def launchDirector(self):
		pass				
	
	# PONT CLICK CARE
	def launchpcc(self):
		searchApp = Tk()
		searchApp.title("Point Click Care Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="(AD) Admin\'s Password: ").grid(row=0)
		Label(searchApp, text="User's AD Account: ").grid(row=1)
				
		v = StringVar()
		x= StringVar()
		
		entryBox = Entry(searchApp, textvariable = v)
		passwordBox = Entry(searchApp, textvariable = x, show="*")

		passwordBox.grid(row=0,column=1, padx=5,pady=5)
		entryBox.grid(row=1,column=1, padx=5,pady=5)
		
		entryBox.focus_set()
		passwordBox.focus_set()

		Button(searchApp, text="Search",command=lambda:callback()).grid(row=2,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=2,column=2, padx=5,pady=5)

		def callback():
			search = entryBox.get()
			pswdBox = passwordBox.get()

			portal = '<PATH>'
			chromeDriver = "<PATH>"
			
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

			# Just click on link for pwd-b/c the element has db info
			reset_pwd_link = ie_driver.find_element_by_link_text('pwd').click()

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
			messagebox.showinfo(title = 'Point Click Care', message = 'Password has been changed to:\n\tPassword1')	
			ie_driver.set_window_size(100,100)
			time.sleep(5)
			ie_driver.close()

			PROCNAME = "chromedriver.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()

		def clear_entryBox():
			entryBox.delete(0,'end')

		mainloop()

	# LAWSON LOOKUP
	def launchKpass(self):
		searchApp = Tk()
		searchApp.title("Lawson Lookup Portal")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text = '(AD) Admin Password: ').grid(row=0)
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
		
		# Enable this feature that if selected RESETS the account and grab password/reset procedure
		#checkRadio = ttk.Radiobutton(searchApp,text='Reset Password? (WIP)')
		#checkRadio.grid(row=1,column=2, padx=5,pady=5)
		#checkCmd = IntVar()
		#checkCmd.set(0)

		usr = StringVar
		v = StringVar()
		x = StringVar()

		adminBox = Entry(searchApp, textvariable = usr)
		passwordBox = Entry(searchApp, textvariable = x, show="*")
		entryBox = Entry(searchApp, textvariable = v)

		adminBox.grid(row=0,column=1, padx=5,pady=5)
		passwordBox.grid(row=1,column=1, padx=5,pady=5)
		entryBox.grid(row=2,column=1, padx=5,pady=5)

		adminBox.focus_set()
		passwordBox.focus_set()
		entryBox.focus_set()
		
		Button(searchApp, text="Lookup User",command=lambda:callback()).grid(row=0,column=2, padx=5,pady=5) #       row =3 col=0
		Button(searchApp, text="Clear",command=lambda:clear_entryBox()).grid(row=1,column=2, padx=5,pady=5)	#	    row =3 col=1
		Button(searchApp, text="Reset Password",command=lambda:resetAccount()).grid(row=2,column=2, padx=5,pady=5)# row =3 col=2

		def callback():
			adminUser = adminBox.get()
			pswdBox = passwordBox.get()
			searchUser = entryBox.get()
			
			portal = '<PATH>'
			chromeDriver = '<PATH>'
			os.environ["webdriver.chrome.driver"] = chromeDriver
			chrome_driver = webdriver.Chrome(chromeDriver)
			chrome_driver.get(portal)
			wait = ui.WebDriverWait(chrome_driver, 15)
						
			time.sleep(2)
			username_field = chrome_driver.find_element_by_xpath("//input[@id='username']")
			username_field.send_keys(adminUser+'@kaleidahealth.org')
			password_field = chrome_driver.find_element_by_id("password")
			password_field.send_keys(pswdBox)
			time.sleep(1)
			login = chrome_driver.find_element_by_xpath("//input[@id='loginSubmit']").click()
			time.sleep(2)
			
			# Click on Users to search
			_users = chrome_driver.find_element_by_xpath("//*[@id='tabUsers']/a").click()
			time.sleep(2)
			
			# Search users
			search = chrome_driver.find_element_by_xpath("//*[@id='searchField']")#.click()
			editLink = chrome_driver.find_element_by_xpath("//tbody/tr/td[6]/div/div[2]/ul/li[2]/a")

			time.sleep(1)
			search.send_keys(searchUser)
			search.send_keys(Keys.RETURN)
			time.sleep(2)
			search.clear()
			entryBox.delete(0,'end')

		def clear_entryBox():
			entryBox.delete(0,'end')

		def resetAccount():
			adminUser = adminBox.get()
			pswdBox = passwordBox.get()
			searchUser = entryBox.get()
			defaultPwd = "Password1"
						
			portal = '<PATH>'
			chromeDriver = '<PATH>'
			os.environ["webdriver.chrome.driver"] = chromeDriver
			chrome_driver = webdriver.Chrome(chromeDriver)
			chrome_driver.get(portal)

			#chrome_driver.maximize_window()
			wait = ui.WebDriverWait(chrome_driver, 15)
						
			time.sleep(2)
			username_field = chrome_driver.find_element_by_xpath("//input[@id='username']")
			username_field.send_keys(adminUser+'@kaleidahealth.org')
			password_field = chrome_driver.find_element_by_id("password")
			password_field.send_keys(pswdBox)
			time.sleep(1)
			login = chrome_driver.find_element_by_xpath("//input[@id='loginSubmit']").click()
			time.sleep(2)
			
			# Click on Users to search
			_users = chrome_driver.find_element_by_xpath("//*[@id='tabUsers']/a").click()
			time.sleep(1) #2
			
			# Search users
			search = chrome_driver.find_element_by_xpath("//*[@id='searchField']")#.click()
			editLink = chrome_driver.find_element_by_xpath("//tbody/tr/td[6]/div/div[2]/ul/li[2]/a")

			time.sleep(1)
			search.send_keys(searchUser)
			search.send_keys(Keys.RETURN)
			time.sleep(1) #2
			search.clear()
			entryBox.delete(0,'end')
			
			# Hidden Menu
			hidden_submenu = chrome_driver.find_element_by_xpath("//*[@id='userTable']/tbody/tr/td[2]/div[1]")
			hover = ActionChains(chrome_driver).move_to_element(hidden_submenu)
			hover.perform()
			time.sleep(1) #2

			# Click Edit - New Window
			edit = chrome_driver.find_element_by_xpath("//*[@id='userTable']/tbody/tr/td[5]/a").click()
			time.sleep(1) #2
			
			# Check if No username
			cortextUsername = chrome_driver.find_element_by_xpath("//*[@id='jidUserName']")
			
			#ERROR CHECK IF NO USERNAME
			if cortextUsername == '':
				messagebox.showinfo(title='Cortext Password Reset', message = 'No Username available for: '+ searchUser + '\nPlease notify your Cortext administrator!')
				PROCNAME = "chromedriver.exe"
				for proc in psutil.process_iter():
					if proc.name() == PROCNAME:
						proc.kill()
				
			# TEST Action Chain, double click > right-click > copy
			actionDoubleClick = ActionChains(chrome_driver).move_to_element(cortextUsername)
			actionDoubleClick.double_click(cortextUsername).perform()
			actionDoubleClick.send_keys(Keys.CONTROL,"a").perform()
			actionDoubleClick.send_keys(Keys.CONTROL,"c").perform()
			
			# USED FOR CORTEXT APPLICATION 
			# Getting username from clipboad, b/c ActionChain CTRL + a & c
			try:
				import tkinter as tk
			except ImportError:
				import tkinter as tk
			copy_and_paste_username = tk.Tk()
			copy_and_paste_username.withdraw()
			capu = copy_and_paste_username.clipboard_get()
			
			time.sleep(2) #3
			# Test getting username from clipboad
			copyUser = actionDoubleClick.send_keys(Keys.CONTROL,"v").perform()

			# Cancel/go out
			cancel = chrome_driver.find_element_by_xpath("//*[@id='cancelEditUserBtn']").click()

			# Show Hidden Menu, again
			hidden_out_submenu = chrome_driver.find_element_by_xpath("//*[@id='userTable']/tbody/tr/td[2]/div[1]")
			hover_out = ActionChains(chrome_driver).move_to_element(hidden_out_submenu).perform()

			time.sleep(1)
			
			# Hover over More
			more_submenu = chrome_driver.find_element_by_xpath("//*[@id='userTable']/tbody/tr/td[6]/div/span")
			hover_menu = ActionChains(chrome_driver).move_to_element(more_submenu).perform()
			time.sleep(1)

			# Click Reset Password/Link Text?
			resetPassword = chrome_driver.find_element_by_xpath("//*[@id='userTable']/tbody/tr/td[6]/div/div[2]/ul/li[2]/a").click()

			time.sleep(1) #2

			# Enter Admin Password
			adminPassword = chrome_driver.find_element_by_xpath("//*[@id='currentPassword']")
			adminPassword.send_keys(pswdBox)

			# Click Print
			printPassword = chrome_driver.find_element_by_xpath("//*[@id='printResetCodeBtn']").click()
			######################################### RESET PASSWORD COMPLETED ######################################### 

			# COPY PASSWORD FROM WEB PAGE
			# Implementing a function to use Cortext application to reset password (launch program >)
			time.sleep(1) #2
			# Switch to new tab
			
			# TEST WINDOW HANDLES - Able to access correct page
			window_handle_before = chrome_driver.window_handles[0]
			window_handle_after = chrome_driver.window_handles[1]
			chrome_driver.switch_to.window(window_handle_after)
			# TEST WINDOW HANDLES

			time.sleep(1)
			
			# Copy Password(Page)
			passDetail = chrome_driver.find_element_by_xpath("//*[@id='print']/div[2]/table/tbody/tr[4]/td")
			get_User_Password = ActionChains(chrome_driver).move_to_element(passDetail)
			get_User_Password.double_click(passDetail).perform()
			get_User_Password.send_keys(Keys.CONTROL,"a").perform()
			get_User_Password.send_keys(Keys.CONTROL,"c").perform()

			try:
				import tkinter as tk
			except ImportError:
				import tkinter as tk
			copy_and_paste_password = tk.Tk()
			copy_and_paste_password.withdraw()
			capp = copy_and_paste_password.clipboard_get()
			
			# REGEX TO MATCH PAASWORD LENGTH (Automate the boring stuff Lesson 29)
			# 12 letters and 1 number Format of password: ABCDEFGHIJKL1 
			# TODO: Create a regex for all capital leeter
			#resetPass = re.compile (r'?<![^A-Z])')

			#passRegex = re.compile (r'(?<![A-Z])[A-Z]{12}')
			passRegex = re.compile (r'(?<![A-Z0-9])[A-Z0-9]{13}')

			# Get password in raw format, need to remove ['']
			pW = passRegex.findall(capp) # ['']

			# Raw format for password
			cortextAppPassword = (''.join(pW))
			cortextAppUsername = capu

			# Close password page
			chrome_driver.close()
						
			time.sleep(1)

			# Switch to origanl page and close
			chrome_driver.switch_to.window(window_handle_before)
			chrome_driver.close()

			# Terminate process Cortext Web
			PROCNAME = "chromedriver.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()
			
			time.sleep(1) #2
			#------------------------ 
			# CORTEXT APPLICATION --- RESET PASSWORD VIA LAUNCHING APPLICATION
			# OR SEND USER THE EMAIL FOR PASSWORD RESET 
			#-----------------------
			app = Application()
			app.start(r"C:\Program Files (x86)\Imprivata\Cortext\Cortext.exe")
			
			cortextApp = app.window_(title_re='.*Imprivata Cortext*')
			cortextApp.SetFocus()
			# Username
			cortextApp.TypeKeys(cortextAppUsername)
			SendKeys("{TAB}")
			# Password
			cortextApp.TypeKeys(cortextAppPassword)
			SendKeys("{ENTER}")

			time.sleep(3)

			# --------------- USING PYWINAUTOGUI ---------------
			# Click Username
			clickUserName = pyautogui.locateCenterOnScreen('<PATH>')
			pyautogui.click((clickUserName))	
			time.sleep(1)
			# Click Account
			account = pyautogui.locateCenterOnScreen('<PATH>')
			pyautogui.click((account))			
			time.sleep(1)
			# Click Change Password
			changepassword = pyautogui.locateCenterOnScreen('<PATH>')
			pyautogui.click((changepassword))
			time.sleep(1)
			# Password Fields			
			pyautogui.typewrite(str(cortextAppPassword), interval=0.05)
			pyautogui.press('tab')
			pyautogui.typewrite(defaultPwd, interval=0.05)
			pyautogui.press('tab')
			pyautogui.typewrite(defaultPwd, interval=0.05)
			pyautogui.press('tab')
			pyautogui.press('tab')
			# Save
			pyautogui.press('enter') 

			time.sleep(3)
			
			# Terminate Cortext
			PROCNAME = "Cortext(32 bit)"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()
			# --------------- USING PYWINAUTOGUI ---------------

			# Change cortextAppPassword to defaultPswd
			messagebox.showinfo(title='Cortext Password Reset', message = 'Cortext Imprivata Username: '+cortextAppUsername+'\nCortext Imprivata Password: ' + cortextAppPassword+'\n\nPassword Changed to: '+ defaultPwd
								+'\nPlease close Cortext Imprivata!'	)
		
		searchApp.mainloop()
		
	# ANAKAM
	def launchAnakam(self):
		# Add Logo
		self.anakamLogo = PhotoImage(file = '<PATH>').subsample(2,2) #2,2
		# Add Logo
		
		searchApp = Tk()
		searchApp.title("Anakam AdminPro")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="(jsmith)Admin Username: ").grid(row=0)
		Label(searchApp, text="(non AD sync)Admin Password: ").grid(row=1)
		Label(searchApp, text="AD Account: ").grid(row=2)

		v = StringVar()
		x = StringVar() 
		usr = StringVar()

		entryBox = Entry(searchApp, textvariable = v)
		pswdBox = Entry(searchApp, textvariable = x, show="*")
		usr_entry = Entry(searchApp, textvariable = usr)

		
		entryBox.grid(row=0, column=1, padx=5,pady=5)
		pswdBox.grid(row=1, column =1, padx =5, pady =5)
		usr_entry.grid(row=2, column =1, padx =5, pady =5)

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
			search.clear()
			entryBox.delete(0,'end')

		mainloop()
		PROCNAME = "chromedriver.exe"
		for proc in psutil.process_iter():
			if proc.name() == PROCNAME:
				proc.kill()
	
	# IMPRIVATA ONESIGN
	def launchOneSign(self):
		searchApp = Tk()
		searchApp.title("Imprivata OneSign")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text = '(AD)Admin Password: ').grid(row=0)
		Label(searchApp, text = '(AD)Username').grid(row=1)
		
		x = StringVar()
		u = StringVar()

		passwordBox = Entry(searchApp, textvariable = x, show="*")
		usr = Entry(searchApp, textvariable = u)

		passwordBox.grid(row=0,column=1,padx=5,pady=5)
		usr.grid(row=1,column=1,padx=5,pady=5)
		
		passwordBox.focus_set()
		usr.focus_set()

		Button(searchApp, text = 'Search', command=lambda:callback()).grid(row=2,column=1,padx=5,pady=5)#.send_keys(Keys.RETURN)
		Button(searchApp, text = 'Clear', command=lambda:clear_entryBox()).grid(row=2,column=2,padx=5,pady=5)

		def callback():
			adminEntry = os.getlogin()
			pswd_Box = passwordBox.get()
			user = usr.get()

			ie_driver = webdriver.Ie("<PATH>")
			ie_driver.get('<PATH>')
			ie_driver.get("javascript:document.getElementById('overridelink').click()")

			wait = ui.WebDriverWait(ie_driver, 15)
			time.sleep(2)

			username_field = ie_driver.find_element_by_id("modalityLabel")
			username_field.send_keys(adminEntry)	

		def clear_entryBox():
			pass
	
	# STARS
	def launchStars(self):	
		# TODO: ask for username/admin username/admin pswd
		searchApp = Tk()
		searchApp.title("STARS Event Review")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)
		
		Label(searchApp, text="(hint:&)Admin Password: ").grid(row=0)
		Label(searchApp, text="Client Name: ").grid(row=1)
		
		adminPassword = StringVar()
		clientUser = StringVar()
		
		pwdEntryBox = Entry(searchApp, textvariable = adminPassword, show ="*")
		clientEntryBox = Entry(searchApp, textvariable = clientUser)

		pwdEntryBox.grid(row=0, column =1, padx=5,pady=5)
		clientEntryBox.grid(row=1, column =1, padx=5,pady=5)

		pwdEntryBox.focus_set()
		clientEntryBox.focus_set()

		Button(searchApp, text="Search", command=lambda:callback()).grid(row=2,column=1,padx=5,pady=5)
		Button(searchApp, text="Clear", command=lambda:clear()).grid(row=2,column=2,padx=5,pady=5)

		def callback():
			defaultID = '<default ID>'
			loginAdmin = os.getlogin()
			loginPwd = pwdEntryBox.get()
			searchClient = clientEntryBox.get()

			# TODO: login to STARS URL
			portal = '<PATH>'
			# TODO: using IE SilverLight(for now until test in Chrome/FF)
			ie_driver = webdriver.Ie("<PATH>")
			#ie_driver = webdriver.Ie("C:\\Users\\dof344\\Desktop\\IEDriverServer.exe")
			ie_driver.get(portal)
						
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
			
			# Click Admin in upper-right
			selAdmin = pyautogui.locateCenterOnScreen('<PATH>')
			pyautogui.click((selAdmin))
				
	# INTERQUAL
	def launchInterqual(self):
		searchApp = Tk()
		searchApp.title("Interqual Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		# Adming username is AD, not needed
		# password label
		Label(searchApp, text = '(&)Admin Password: ').grid(row=0)
		Label(searchApp, text = 'Username: ').grid(row=1)

		v = StringVar()
		u = StringVar()

		passwordBox = Entry(searchApp, textvariable = v, show='*')
		userBox = Entry(searchApp, textvariable = u)

		passwordBox.grid(row=0,column=1,padx=5,pady=5)
		userBox.grid(row=1,column=1,padx=5,pady=5)

		passwordBox.focus_set()
		userBox.focus_set()

		Button(searchApp, text = 'Search', command=lambda:callback()).grid(row=2,column=1,padx=5,pady=5)
		Button(searchApp, text = 'Clear', command=lambda:clear()).grid(row=2,column=2,padx=5,pady=5)

		def callback():
			_admin = os.getlogin()
			_pwd = passwordBox.get()
			_user = userBox.get()

			portal = '<PATH>'
			ie_driver = webdriver.Ie("<PATH>")
			ie_driver.get(portal)

			adminUsername = ie_driver.find_element_by_xpath("//*[@id='userName']")
			adminUsername.send_keys(_admin)

			adminPassword = ie_driver.find_element_by_xpath("//*[@id='password']")
			adminPassword.send_keys(_pwd)
			time.sleep(2)
			login = ie_driver.find_element_by_xpath("//*[@id='login']").click()

			time.sleep(10)

			ie_driver.close()

		def clear(self):
				pass
	
	# TIMELESS BREASTMILK
	def launchTimeless(self):
		searchApp = Tk()
		searchApp.title("Timeless Breastmilk Account Unlock")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="(AD) Admin Password: ").grid(row=0)
		Label(searchApp, text="User\'s AD Account: ").grid(row=1)
		
		v = StringVar()
		x = StringVar() 

		pswdBox = Entry(searchApp, textvariable = x, show="*")
		entryBox = Entry(searchApp, textvariable = v)
		
		pswdBox.grid(row=0, column =1, padx =5, pady =10)
		entryBox.grid(row=1, column=1, padx=5,pady=10)

		pswdBox.focus_set()
		entryBox.focus_set()

		Button(searchApp, text="Search",command=lambda:callback()).grid(row=2,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear()).grid(row=2,column=2, padx=5,pady=5)

		def callback():
			user = entryBox.get()
			passwordBox = pswdBox.get()

			portal = '<PATH>'
			chromeDriver = "<PATH>"
			os.environ["webdriver.chrome.driver"] = chromeDriver
			chrome_driver = webdriver.Chrome(chromeDriver)
			chrome_driver.get(portal)
			wait = ui.WebDriverWait(chrome_driver, 15)		

			adminAcc = os.getlogin()
			username_field = chrome_driver.find_element_by_xpath("//*[@id='loginName']")
			username_field.send_keys(adminAcc)

			password_field = chrome_driver.find_element_by_xpath("//*[@id='loginPassword']")
			password_field.send_keys(passwordBox)
			password_field.send_keys(Keys.RETURN)
			time.sleep(2)

			menu = chrome_driver.find_element_by_xpath("//*[@id='low-res-menu-button']").click()
			time.sleep(1)
			manage_users = chrome_driver.find_element_by_xpath("//*[@id='main_menu']/ul[2]/li[2]/a").click()
			time.sleep(1)
			manage_users_new = chrome_driver.find_element_by_xpath("//*[@id='sub_manage']/li[3]/a").click()
			userID = chrome_driver.find_element_by_xpath("//*[@id='userLoginBarcodesearch']")
			userID.send_keys(user)
			userID.send_keys(Keys.RETURN)
			edit = chrome_driver.find_element_by_xpath("//*[@id='users']/table/tbody/tr[3]/td[6]/nobr/a[1]").click()

			# Check if account locked first -- needed dummy accounts for testing
			time.sleep(2)
			'''
				Reversed logic? -- else selected the checkbox
					*note: if not
					checkbox
					//*[@id="userIncorrectLogins"]
	
					submit 
					//*[@id="submitusers"]
	
					logout
					//*[@id="logout-button"]
			'''
			try:
				locked_out = chrome_driver.find_element_by_xpath("//input[@id='userIncorrectLogins']")
				if not(locked_out.is_selected()):
					# Select the check box, send_keys(Keys.SPACE) can work
					locked_out.click()
					
					# click submit
					submit = chrome_driver.find_element_by_xpath("//*[@id='submitusers']").click()
					time.sleep(2)
					logout = chrome_driver.find_element_by_xpath("//*[@id='logout-button']").click()
					chrome_driver.close()

					messagebox.showinfo(title='Timeless Breastmilk', message = 'Account is unlocked!\nTimeless Breastmilk uses your AD account password!' +
										'\nClick Ok to close')
				else:
					messagebox.showinfo(title='Timeless Breastmilk', message = 'Account is not locked!\nTimeless Breastmilk uses your AD account password!' +
										'\nTimeless will close in 10s')
					time.sleep(10)
					logout = chrome_driver.find_element_by_xpath("//*[@id='logout-button']").click()
					time.sleep(15)
					chrome_driver.close()
			except Exception as e:
				messagebox.showinfo(title='Timeless Breastmilk', message = 'Account is not locked!\nTimeless Breastmilk uses your AD account password!' +
										'\nClick Ok to close')
				time.sleep(5)

				chrome_driver.quit()
			
			PROCNAME = "chromedriver.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()

		def clear():
			entryBox.delete(0,'end')
			
		mainloop()
	
	# INFOCLIQUE - Page need to be logged in before able to access
	# Launch Kaleidascope, then close
	def launchInfoclique(self):
		# Use JavaScript Selector feature Section 8 essential JS training
		searchApp = Tk()
		searchApp.title("Password Reset")
		searchApp.geometry("400x125")
		searchApp.resizable(False,False)

		Label(searchApp, text="Admin Password: ").grid(row=0)
		Label(searchApp, text="User\'s First/Lastname: ").grid(row=1)
		
		v = StringVar()
		x = StringVar() 

		pswdBox = Entry(searchApp, textvariable = x, show="*")
		entryBox = Entry(searchApp, textvariable = v)
		
		pswdBox.grid(row=0, column =1, padx =5, pady =10)
		entryBox.grid(row=1, column=1, padx=5,pady=10)

		pswdBox.focus_set()
		entryBox.focus_set()

		Button(searchApp, text="Search",command=lambda:callback()).grid(row=2,column=1, padx=5,pady=5)
		Button(searchApp, text="Clear",command=lambda:clear()).grid(row=2,column=2, padx=5,pady=5)

		def callback():
			_admin = os.getlogin()
			_pwd = pswdBox.get()
			_user = entryBox.get()

			# Note - Infolcque finds user name by lastname first,
			# Swap name order so Lastname, Firstname (better accuracy)
			
			# Use split for using the comma
			split_user = _user
			
			# Reversing the order word: Firstname Lastname  -> Lastname Firtname, the comma is added with .join
			reverseUser = re.split("\W+", _user)
			reverseUser.reverse()
			combineUser = ' '.join(reverseUser)
			
			split_user2 = combineUser.split(" ")
			
			# Join with a comma
			find_user = ", ".join(split_user2)

			portal = '<PATH>'
			portal2 = '<PATH>'

			# Start: Using Chrome Driver
			chromeDriver = "<PATH>"
			os.environ["webdriver.chrome.driver"] = chromeDriver
			ie_driver = webdriver.Chrome(chromeDriver)
			ie_driver.get(portal)
			wait = ui.WebDriverWait(ie_driver, 5)
			# End: Using Chrome Driver

			adminUsername = ie_driver.find_element_by_xpath("//*[@id='ucLogin_txtUser']")
			adminUsername.send_keys(_admin)
			adminPassword = ie_driver.find_element_by_xpath("//*[@id='ucLogin_txtPW']")
			adminPassword.send_keys(_pwd)
			adminPassword.send_keys(Keys.RETURN)

			time.sleep(2)
			ie_driver.get(portal2)
			ie_driver.set_window_size(800,950)
			ie_driver.set_window_position(600,100) # 250,200
			#ie_driver.maximize_window()
			
			time.sleep(3)
			
			# Take a screenshot of imaage and use loC to find
			selUser = pyautogui.locateCenterOnScreen('<PATH>')
			
			pyautogui.click((selUser))
			time.sleep(1)
			# typewrite takes string arg
			pyautogui.typewrite(str(find_user), interval=0.2)
			pyautogui.click()
			time.sleep(1)
			
			resetPassword = pyautogui.locateCenterOnScreen('<PATH>')
			pyautogui.click((resetPassword))
			
			time.sleep(5)
			
			clickYes = pyautogui.locateCenterOnScreen('<PATH>')			
			pyautogui.click((clickYes))
			time.sleep(10)
			ie_driver.close()

			PROCNAME = "chromedriver.exe"
			for proc in psutil.process_iter():
				if proc.name() == PROCNAME:
					proc.kill()

			''' ******* Development Notes *************:
			# Only need page to login. Disregard below if don't need secured application page
			#time.sleep(2)
			# Go To Application Tab <PATH>, CSS Selector
			#applicationTab = ie_driver.get('<PATH>')
			#ie_driver.find_element_by_xpath("//*[href='/applications']").click()
			appTab = ie_driver.find_element_by_css_selector("#tabs > td:nth-child(3) > table:nth-child(1) > tbody:nth-child(1) > tr:nth-child(1) > td:nth-child(1) > a:nth-child(6) > img:nth-child(1)")
			appTab.click()			
			main_window = ie_driver.window_handles[0]
			pswdResetLink = ie_driver.find_element_by_link_text("<PATH>")
			pswdResetLink.click()
			opened_window = ie_driver.window_handles[1]
			time.sleep(3)
			#ie_driver.switch_to.window(opened_window)
			# loadUser = ui.WebDriverWait(ie_driver,15).until(lambda: ie_driver.find_element_by_css_selector("#ddlUserList"))
			ie_driver.switch_to.window(opened_window)
			time.sleep(5)			
						
			* NOTE - Implement selecting the selector for grabbing the names
			New window opens
			Can begin to type name to get to user
			<SELECT id=ddlUserList style="WIDTH: 100%" name=ddlUserList> 
				<option value ="AD Name">Last, First name</option>
				No handle to new window, Window before/after?
			# Uisng JavaScript to access Select box
			#ie_driver.execute_script("window.alert('This is a TEST')");
			#selectUser = ie_driver.execute_script("document.getElementById('ddlUserList')");
			#selectUser.send_keys(_user);
			#print("Found option 4: ", selectUser)

			sel = ie_driver.find_element_by_css_selector("[name='ddlUserList']")
			sel.select_by_visible_text("APA505").click()
			#selUser = document.getElementById("ddlUserList");
			#selUser = ie_driver.find_element_by_xpath("//*[@id='ddlUserList']/option[7]")
			#selUser.click()
			for option in selUser.find_element_by_tag_name('option'):
				if option.text == 'APA505':
					Option.click()
					break
			#selUser = Select(ie_driver.find_element_by_xpath("//*[@id='ddlUserList']"))
			#selUser = ie_driver.find_element_by_xpath("/html/body/table/tbody/tr[114]/td[2]/text()").click()
			#selUser = ie_driver.find_element_by_css_selector("body > table > tbody > tr:nth-child(116) > td.line-content > span")
			#selUser.select_by_index(4)
			#selUser.send_keys(_user)
			#select = Select(ie_driver.find_element_by_id("ddlUserList"))
			# selectUser = ie_driver.find_element_by_css_selector("p.ddlUserList")
			# option = selectUser.find_element_by_css_selector("option[value$=APA505]")
			# option.click()
			'''
	
		def clear():
			entryBox.delete(0,'end')
			
		searchApp.mainloop()

	# TWILIO - Uses API to send SMS to verified phone number
	def launchTextMsg(self):
		# Create window to recieve message, recipeint# 
		msgGUI = Tk()
		msgGUI.title('Twilio SMS Messager')
		msgGUI.geometry('550x375')
		msgGUI.resizable(True,True)
		
		#msgGUI_label = Label(text='Twilio Messanger')

		self.frame_content = ttk.Frame(msgGUI)
		self.frame_content.pack()

		ttk.Label(self.frame_content, text = 'Recipient Phone Number:').grid(row=1, column=0,padx=5,pady=5,sticky='sw')
		ttk.Label(self.frame_content, text = 'Message:').grid(row=2, column=0,padx=5,pady=5,sticky='sw')

		self.entry_Recipient = ttk.Entry(self.frame_content, width=24, font= NORM_FONT)
		self.entry_Message = Text(self.frame_content, width=50, height=10, font= NORM_FONT)


		self.entry_Recipient.grid(row=1,column=1,columnspan=2,padx=5)
		self.entry_Message.grid(row=3,column=0,columnspan=3,padx=5)

		ttk.Button(self.frame_content, text = 'Send',
					command = self.sendMSG).grid(row=4, column=0,padx=5,pady=5,sticky='e')
		ttk.Button(self.frame_content, text = 'Cancel',
					command = msgGUI.destroy).grid(row=4, column=1,padx=5,pady=5,sticky='w')
		msgGUI.mainloop()
	def sendMSG(self):
		client = TwilioRestClient(account_sid, auth_token)
		_recipient = self.entry_Recipient.get()
		_msg = self.entry_Message.get(1.0, 'end')
		message = client.messages.create(to=_recipient, from_=my_twilio, body=_msg)

		'''
		# WORKS
		client = TwilioRestClient(account_sid, auth_token)
		my_msg = "Your message here (use and interface/get phonenumber to use)"
		message = client.messages.create(to=my_cell, from_=my_twilio, body=my_msg)
		#client.message.create(my_cell, my_twilio, my_msg)
		# WORKS
		'''

	# Placeholder
	def ph2(self):
		pass
	# Placeholder
	def ph3():
		pass
	# Placeholder
	def ph4():
		pass
	# Placeholder
	def ph5():
		pass


	# SUBMIT FUNCTION		
	def submitForm(self):
		gui = Tk()
		gui.title('Submit Feedback')
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
		emailAddress = self.entry_email.get()
		comment = self.text_comments.get(1.0, 'end')
		phone = self.phone_number.get()

		# SEND EMAIL: http://stackoverflow.com/questions/10147455/how-to-send-an-email-with-gmail-as-provider-using-python/27515833#27515833
		smtpObj = smtplib.SMTP_SSL('smtp.gmail.com', 465)

		smtpObj.ehlo()

		# Create dummy dccount (in settings file)
		#smtpObj.login('username','password')

		smtpObj.login(SENDER_EMAIL, SENDER_PASSWORD)


		# Forward emails to this account
		smtpObj.sendmail(SENDER_EMAIL, 'email@gmail.com','Subject: Feedback Submitted\n'
							+'\nName: '+str(nameInfo)+'\nEmail Address: '+str(emailAddress)+'\nPhone Number:' + str(phone)
							+'\n\nComments:\n '+str(comment))
		
		# Clear user info
		self.clearButton()
		
		# Popup notifying on message submission
		messagebox.showinfo(title = 'Comments', message = 'Your comments have been submitted.')	

		# Close gmail connection
		smtpObj.close()

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
			if radioButton == 1 and '<Lastname>, <Firstnam>':
				print('Anaylst: ',index,row['Avg ACW Time'])
			elif (v.get()) == 2 and '<Lastname>, <Firstnam>':
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
		ttk.Radiobutton(master, text = 'Name', command = lambda:self.acwData(), variable = v, value=1).grid(row=0,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name', command = lambda:self.acwData(), variable = v, value=2).grid(row=0,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name', command = lambda:self.acwData(), variable = v, value=3).grid(row=0,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=4).grid(row=1,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=5).grid(row=1,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=6).grid(row=1,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=7).grid(row=2,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=8).grid(row=2,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=9).grid(row=2,column=3, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
		ttk.Radiobutton(master, text = 'Name',command = lambda:self.acwData(), variable = v, value=10).grid(row=3,column=1, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)	
		ttk.Radiobutton(master, text = 'None', command = lambda:self.acwData(), variable = v, value=11).grid(row=3,column=2, sticky='NSEW', padx=5, pady=5)#.pack(anchor=W)
	
		mainloop()
		# Select Anaylst			
		# Implement a DB to read current ACW times
		# Separate script to run for these numbers (i.e Metrics excel data, conver to Access?)
		# Open in separate window?
		# Using downloaded CMS reports
		# Pull the xls. data into pandas and convert to xlsx
		##file = '<PATH>'
def main():
	root = Tk()
	# Window Details
	root.title("KH Automation version 0.3.3")
	root.resizable(False,True) # True, True
	root.geometry("400x400")
	# Background color
	root.configure(background = '#010a0d') #006A85
	app = App(root)
	root.mainloop()

	
if __name__ == "__main__": main()

	
	