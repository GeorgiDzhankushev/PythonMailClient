#!/usr/bin/python

# Python 3

import wx, smtplib, imaplib, email, wx.html
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.application import MIMEApplication
from os.path import basename

class Mail(wx.Frame):

	# calls the parent's init method
	def __init__(self, *args, **kw):
		super(Mail, self).__init__(*args, **kw)
		self.initUI()

	# attempts to send the email written by the user
	def sendMail(self, obj):
		self.msg['From'] = self.sender
		self.msg['To'] = self.recepientField.GetValue()
		self.msg['Subject'] = self.titleField.GetValue()
		self.msg.attach(MIMEText(self.bodyField.GetValue(), 'plain'))
		text = self.msg.as_string()
		try:
			self.SMTPserver.sendmail(self.sender, self.recepientField.GetValue(), text)
			lbl = wx.StaticText(self.pnl, label = 'email sent successfully!', pos = (145, 445))
		except:
			lbl = wx.StaticText(self.pnl, label = 'failed to send message.', pos = (145, 445))
	
	# selects files to include as attachments
	def attach(self, obj):
		dialog = wx.FileDialog(self.pnl, "Select files", style= wx.FD_MULTIPLE)
		if dialog.ShowModal() == wx.ID_CANCEL:
			return
		filePaths = dialog.GetPaths()
		for file in filePaths:
			part = MIMEApplication(open(file, 'rb').read(), Name = basename(file))
			part['Content-Disposition'] = 'attachment; filename="%s"' % basename(file)
			self.msg.attach(part)
		if len(filePaths) > 0:
			numberOfAttachments = 'attachments: ' + str(len(filePaths))
			lbl = wx.StaticText(self.pnl, label = numberOfAttachments , pos = (180, 410))

	# draws a window for composing a message
	def compose(self, obj):
		frame = wx.Frame(self, title = 'Compose', size = (460, 530))
		self.pnl = wx.Panel(frame)
		frame.Centre()
		frame.Show(True)
		self.msg = MIMEMultipart ()
		self.recepientField = wx.TextCtrl(self.pnl, value = 'recepient@example.com', pos = (40, 30), size = (380, 35))
		self.titleField = wx.TextCtrl(self.pnl, value = 'Title', pos = (40, 80), size = (380, 35))
		self.bodyField = wx.TextCtrl(self.pnl, value = 'Message body.', pos = (40, 130), size = (380, 250), style = wx.TE_MULTILINE)
		attachButton = wx.Button(self.pnl, label = 'attach', pos = (40, 400))
		sendButton = wx.Button(self.pnl, label = 'send', pos = (335, 400))
		sendButton.Bind(wx.EVT_BUTTON, self.sendMail)
		attachButton.Bind(wx.EVT_BUTTON, self.attach)

	# displays the contents of the selected email
	def displayMail(self, e):
		obj = e.GetEventObject()
		self.senderBox.SetValue('From: ' + self.mailList[obj.GetSelection()][1])
		self.contentsBox.SetPage(self.mailList[obj.GetSelection()][2])

	# helper function for extracting the body of an email
	def getBody(self, msg):
		if msg.is_multipart():
			return self.getBody(msg.get_payload(0))
		else:
			return msg.get_payload(None, True)

	# retrieves every email from the user's inbox
	def getMail(self, obj):
		self.pnl.Destroy()
		self.pnl = wx.Panel(self)
		self.SetSize((620, 470))
		self.SetTitle('Inbox')
		self.Centre()
		btn = wx.Button(self.pnl, pos = (70, 20), size = (100, 40), label = 'compose')
		btn.Bind(wx.EVT_BUTTON, self.compose)
		self.mailList = []
		self.IMAPserver.select("Inbox")
		tmp, data = self.IMAPserver.search(None, 'ALL')
		for num in data[0].split():
			tmp, data = self.IMAPserver.fetch(num, '(RFC822)')
			raw = email.message_from_bytes(data[0][1])
			self.mailList.append((raw['Subject'], raw['From'], self.getBody(raw)))
		titles = []
		for i in range (len(self.mailList)):
			if self.mailList[i][0] == None:
				titles.append('Untitled')
			else:
				titles.append(self.mailList[i][0])
		titlesBox = wx.ListBox(self.pnl, pos = (20, 80), size = (250, 330), choices = titles)
		self.senderBox = wx.TextCtrl(self.pnl, pos = (240, 20), size = (350, 40), style = wx.TE_READONLY)
		self.contentsBox = wx.html.HtmlWindow(self.pnl, -1, pos = (280, 80), size = (330, 330))
		titlesBox.Bind(wx.EVT_LISTBOX, self.displayMail)

	# attempts to login to the SMTP and IMAP servers
	def login(self, e):
		self.domainField.GetValue() == '@gmail.com':
		self.domainSMTP = 'smtp.gmail.com'
		self.domainIMAP = 'imap.gmail.com'
		self.port = 587
		try:
			self.SMTPserver = smtplib.SMTP(self.domainSMTP, self.port)
			self.SMTPserver.starttls()
			self.sender = self.senderField.GetValue() + self.domainField.GetValue()
			self.password = self.passwordField.GetValue()
			self.SMTPserver.login(self.sender, self.password)
			self.IMAPserver = imaplib.IMAP4_SSL(self.domainIMAP)
			self.IMAPserver.login(self.sender, self.password)
			self.getMail(self)
		except:
			self.lbl = wx.StaticText(self.pnl, label = 'Failed to login.', pos = (90, 220))

	# initializes the UI for the Login window
	def initUI(self):
		self.pnl = wx.Panel(self)
		self.senderField = wx.TextCtrl(self.pnl, value = 'name', pos = (20, 30), size = (150, 35))
		self.passwordField = wx.TextCtrl(self.pnl, value = 'password', pos = (20, 80), size = (282, 35), style= wx.TE_PASSWORD)
		self.domainField = wx.ComboBox(self.pnl, pos = (172, 30), size = (130, 35), choices = ['@gmail.com'], style = wx.CB_DROPDOWN)
		self.domainField.SetValue('@gmail.com')
		btn = wx.Button(self.pnl, label = 'Login', pos = (120, 150))
		btn.Bind(wx.EVT_BUTTON, self.login)
		self.SetSize((322, 300))
		self.SetTitle('Login')
		self.Centre()
		self.Show(True)

# creates a class instance and a loop to handle the events
def main():
	app = wx.App()
	Mail(None)
	app.MainLoop()

if __name__ == '__main__':
	main()
