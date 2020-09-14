from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5 import QtCore
import sys
import xlrd 
loc = ("C:/Users/Arun/Documents/Sublime/completed/iquizzer/questions.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 

class win(QWidget):
	def __init__(self):
		super().__init__()
		self.setGeometry(0,0,1350,690)
		self.initUI()
	def initUI(self):
		self.status_label=QPushButton(self)
		self.status_label.setGeometry(830,250,350,30)
		self.status_label.setText("Powered by The Scripted Codes")
		self.status_label.setFont(QFont('Arial', 15))
		self.status_label.setStyleSheet("background-color: STEELBLUE")
		self.score_label=QLabel(self)
		self.score_label.move(380,50)
		self.score_label.setText("Score ")
		self.score_label.setFont(QFont('Arial', 15))
		self.score_label.hide()
		self.ques_label=QPushButton(self) 
		self.ques_label.setText("Welcome to iQuizzer")
		self.ques_label.setGeometry(150, 100, 1050, 100)
		self.ques_label.setFont(QFont('Arial', 17))
		self.fixed_label=QLabel(self)
		self.fixed_label.move(850, 50)
		self.fixed_label.setFont(QFont('Arial', 14))
		self.total_questions=int(sheet.cell_value(0, 1))
		self.fixed_label.hide()
		self.fixed_label.setText("Total Questions :  "+str(self.total_questions))
		self.fixed_label.adjustSize()
		self.option_a=QPushButton(self)
		self.option_a.setGeometry(250, 350, 400, 30)
		self.option_a.setFont(QFont('Arial', 12))
		self.option_b=QPushButton(self)
		self.option_b.setGeometry(750, 350, 400, 30)
		self.option_b.setFont(QFont('Arial', 12))
		self.option_c=QPushButton(self)
		self.option_c.setGeometry(250, 500, 400, 30)
		self.option_c.setFont(QFont('Times', 12))
		self.option_d=QPushButton(self)
		self.option_d.setGeometry(750, 500, 400, 30)
		self.option_d.setFont(QFont('Arial', 13))
		self.options_dict=[self.option_a,self.option_b,self.option_c,self.option_d,]
		self.thisdict = {
  				"A":self.option_a,
 				"B":self.option_b,
 				"C":self.option_c,
  				"D":self.option_d
				}
		self.new=QPushButton(self)
		self.new.setGeometry(660,600,100,30)
		self.new.setText("NEXT ")
		self.new.setFont(QFont('Arial',15))
		self.new.hide()
		self.begin=QPushButton(self)
		self.begin.setGeometry(620,540,100,30)
		self.begin.setText("BEGIN ")
		self.begin.setFont(QFont('Arial',15))
		self.restart_butn=QPushButton(self)
		self.restart_butn.setGeometry(560,250,300,30)
		self.restart_butn.setText("Play Once Again ?")
		self.restart_butn.setFont(QFont('Arial',15))
		self.restart_butn.hide()
		self.qposition=2
		self.score=0
		self.new.clicked.connect(self.exam_display)
		self.option_a.clicked.connect(self.user_ans_a)
		self.option_b.clicked.connect(self.user_ans_b)
		self.option_c.clicked.connect(self.user_ans_c)
		self.option_d.clicked.connect(self.user_ans_d)
		self.restart_butn.clicked.connect(self.restart)
		self.begin.clicked.connect(self.initials)
		for x in range(len(self.options_dict)):
			self.options_dict[x].hide()
	def initials(self):
		self.status_label.setGeometry(600,50,150,30)
		self.status_label.setText(" Question No: ")
		self.fixed_label.show()
		self.score_label.show()
		self.score_label.setText("Score")
		self.new.show()
		self.begin.hide()
		self.status_label.setStyleSheet("background-color: None")
		self.scoreupdate=True
		self.option_a.setText("Option A")
		self.option_b.setText("Option B")
		self.option_c.setText("Option C")
		self.option_d.setText("Option D")
		self.ques_label.setText("Question")

		for x in range(len(self.options_dict)):
			self.options_dict[x].show()
			self.options_dict[x].setStyleSheet("background-color:None")
	def exam_display(self):
		global sheet,ques_number
		self.scoreupdate=True

		self.qposition+=1
		self.status_label.setText("  Question "+str(self.qposition-2)+"  ")
		for x in range(len(self.options_dict)):
			self.options_dict[x].setStyleSheet("background-color: None")
		self.ques_label.setText(sheet.cell_value(self.qposition, 0))
		self.ques_label.setFont(QFont('Arial', 15))
		self.option_a.setText(sheet.cell_value(self.qposition, 1))
		self.option_b.setText(sheet.cell_value(self.qposition, 2))
		self.option_c.setText(sheet.cell_value(self.qposition, 3))
		self.option_d.setText(sheet.cell_value(self.qposition, 4))
		self.new.hide()
	def user_ans_a(self):
		self.user_ans='A'
		self.opkey=0
		self.ans_check()
	def user_ans_b(self):
		self.user_ans='B'
		self.opkey=1
		self.ans_check()
	def user_ans_c(self):
		self.user_ans='C'
		self.opkey=2
		self.ans_check()
	def user_ans_d(self):
		self.user_ans='D'
		self.opkey=3
		self.ans_check()
	def ans_check(self):
		try:
			if(self.user_ans ==sheet.cell_value(self.qposition, 5).upper() and self.scoreupdate == True ):
				self.score+=1
				self.scoreupdate=False
				self.thisdict[sheet.cell_value(self.qposition, 5)].setStyleSheet("background-color: Green")	
			else:
				self.options_dict[self.opkey].setStyleSheet("background-color: Red")
				self.scoreupdate=False
				self.thisdict[sheet.cell_value(self.qposition, 5)].setStyleSheet("background-color: Green")
			self.score_label.setText("Your Score : "+str(self.score) +"/"+str(self.qposition-2))
			self.score_label.adjustSize()
			self.new.show()
			if(self.qposition-2==self.total_questions):
				self.restart_butn.show()
				self.new.hide()
				self.status_label.setText("Attempted All Questions")
				self.status_label.adjustSize()
				self.status_label.setStyleSheet("background-color : Brown")

		except:
			self.ques_label.setText("Game not Started / Error")
	def restart(self):
 		self.qposition=2
 		self.score=0
 		self.restart_butn.hide()
 		self.status_label.setStyleSheet("background-color : None")
 		self.initials()
if __name__ == "__main__":
	app=QApplication(sys.argv)
	win=win()
	win.show()
	sys.exit(app.exec_())
