# -*- coding: utf-8 -*-
#

"""
"""
SHOW = True # Show test in GUI-based test launcher

import sys,os
from guidata.dataset.qtwidgets import  DataSetEditGroupBox
from guidata.qt.QtGui import QDialog,QGridLayout,QPushButton,QMainWindow,QStatusBar,QMessageBox
from guidata.dataset.datatypes import DataSet,GetAttrProp, FuncProp, BeginGroup, EndGroup
from guidata.dataset.dataitems import (IntItem,BoolItem,MultipleChoiceItem,FilesOpenItem,StringItem,ChoiceItem,DirectoryItem)
from ppt_v12 import easyPPT
from Webppt_v11 import *

global pyrealpath,singing
global labAtit,Num,outdir
global ImgTYPE,imageDir,width

fun2 = lambda x: os.path.basename(x)
fun1=lambda x:os.path.splitext(x)[0]
func=lambda x:fun2(fun1(x))

class InoutSet(DataSet):
	"""【输出设置】区各项目的定义、及类型设置.

		"""
	def updatedir(self, item, value):
		if not value or not len(value):
			pass
		else:
			self.outprefix = ";".join(map(func, value))
			self.WebTitle = self.outprefix
			self.outpath = os.path.dirname(value[0])
			self.FilesDirName = self.outprefix
			# self.results = u'您选择的文件是：'.encode("gbk")
			# self.results = "\r\n".join(map(fun2, value))


	PPTnames = FilesOpenItem(u" 输入", ("ppt", "pptx"), help=u"选择要处理的PPT文件（必选、同路径下可批量处理）", all_files_first=True).set_prop(
		'display', callback=updatedir)
	PPTnames.set_prop("data",wordwrap=True)
	outpath = DirectoryItem(u"输出路径",
							help=u"(可选)选择存放输出文件夹的位置，默认为输入文件所在路径")  # .set_prop('display')
	WebTitle = StringItem(u"网页标题\n(Title)",
						  help=u"(可选)输入要替换显示的HTM、XML标题，默认为当前处理的PPT文件名\n批处理，对所有文件应用此名称").set_prop('display', active=True)

	outprefix = StringItem(u"文件名称",
						help=u"(可选)输入要替换显示的文件名(应用于当前所有输出文件)，默认为当前处理的PPT文件名\n批处理，对所有输入应用此名称").set_prop("display",active=True,wordwrap=True)
	# results = TextItem(u"已输入文件 ", default=u"已选择输入的PPT文件:\n").set_prop('display', hide=True, active=False)

class ImageSet(DataSet):
	"""【图片设置】区各项目的定义、及类型设置.

		"""
	def update(self, item, value):#更新缩放图名称
		self.ImagesDirName = unicode(value)

	choice=(("png","PNG"),("jpg","JPG"),("gif","GIF"))
	imgtype = ChoiceItem(u"目标格式", choice,help=u"若选择“无”,表示不进行任何图片操作(如:导出原图、缩放图、长图、HTM\XmL)")#.set_prop("display",active=prop2)#.set_pos(col=1)
	#imgtype = ImageChoiceItem(u"图片格式", ["PNG", "JPG"],help=u"选择输出图片格式").vertical(2),callback=update
	raw=BoolItem(u"原图",default=False,help=u"(可选)直接从PPT中导出原尺寸的PNG图片，无缩放\n存放目录为:Slides\\").set_pos(col=0)
	_prop = GetAttrProp("resize")
	#choice = ChoiceItem('Choice', choices).set_prop("display", store=_prop)
	resize = BoolItem(u"缩放图",default=True,help=u"在原尺寸的基础上对宽、高进行按比例缩放,默认勾选").set_pos(col=1).set_prop("display",store=_prop)
	newsize = IntItem(u"重新调整\n宽度", default=709, min=0, help=u"(可选)输出的缩放图宽度，单位dpi,，默认709，上限不得超过原图宽度。",max=2160, slider=True).set_prop("display",active=FuncProp(_prop, lambda x: x ))
	ImagesDirName=StringItem(u"缩放图集名称",default="images",
		help=u"(可选)输入存放调整后图片的目录名，默认images").set_prop('display',active=FuncProp(_prop, lambda x: x))

class FormatSet(DataSet):
	"""【格式设置】区各项目的定义、及类型设置.

	"""
	_prop=GetAttrProp("langue")
	Langue=(("zh-cn",u"zh-cn"),("en_us",u"en_us"),(None,u"无"))#加入发布的中英文选择
	g0=BeginGroup(u"① 网络文件")
	webFormat = MultipleChoiceItem(u"",[ "HTM   "," XML  "],help=u"(可选)默认全选",
								  default=(0,1)).set_prop('display',active=FuncProp(_prop,lambda x:x!=None)).vertical(2).set_pos(col=0)
	# _prop3=GetAttrProp("langue")
	langue=ChoiceItem(u"\r\n发布语言", Langue,help=u"选择要发布的语言类型,若选择无，表示不输出htm、xml文件").set_pos(col=0,colspan=2).set_prop("display")
	_g0=EndGroup(u"网络文件")
	txpdf=MultipleChoiceItem(u"② 文本文件",["PDF "," TXT"],default=(0,),help=u"默认输出pdf").set_pos(col=1)
	singimg=MultipleChoiceItem(u"③ 图片",[u"原图(集)",u"缩放图(集)",u"长图(目标格式)"],default=(0,1,2),help=u"默认输出原图集、缩放图集\n如要生成htm、xml，请保持勾选")#.set_pos(col=1)
	g1=BeginGroup(u"④ PPT(单张)")
	_prop1 = GetAttrProp("singppt")#g3=,"PPT(X)"
	singppt=BoolItem("PPT(X)",default=False,help=u"发布单张幻灯片").set_prop('display',store=_prop1).set_pos(col=0)
	expind=StringItem(u"幻灯片编号",help=u"输入要单张发布成PPT的幻灯片编号,英文逗号分割，连续编号短线相连.1,2,3,5-7,8\n存放目录为Slides2PPT\\",
					  default="1,2,3").set_prop('display',active=FuncProp(_prop1, lambda x: x)).set_pos(col=1)
	_g1=EndGroup("PPT")

pptx = easyPPT()#初始化easyPPT对象，等待使用

class MainWindow(QMainWindow):
	"""u"EasyPPT_Kits.

	主操作面板布局、信号传递/响应、后台文件处理函数定义等
	"""
	args={}
	shortnm=u""
	def __init__(self,parent=None):
		"""定义、初始化面板各元素"""
		super(MainWindow,self).__init__(parent)
		self.setWindowTitle(u"EasyPPT_Kits")
		self.setGeometry(50, 50, 300, 250)
		self.central=QDialog()

		self.groupbox1 = DataSetEditGroupBox(u"【名称设置】",InoutSet,show_button=False)
		self.groupbox2 = DataSetEditGroupBox(u"【图片设置】",ImageSet,show_button=False,comment='')
		self.groupbox3 = DataSetEditGroupBox(u"【格式设置】",FormatSet,show_button=False,comment='')
		pptlist = self.groupbox1.dataset.PPTnames

		self.groupbox2.setEnabled(isinstance(pptlist,list))
		self.groupbox3.setEnabled(isinstance(pptlist,list))
		# button_icon = get_icon("apply.png")
		# quit_icon=get_icon("apply.png")
		self.bimg = QPushButton(u"激活【图片设置】", self)
		self.bimg.setCheckable(True)
		self.bimg.clicked.connect(self.Enableimg)
		self.btn = QPushButton(u"文件详情..", self)
		self.btn.clicked.connect(self.printdetail)#show_button=False
		self.btnok=QPushButton("APPLY",self)
		self.btnok.clicked.connect(self.update_groupboxes)  # show_button=False
		self.statusBar = QStatusBar()
		self.statusBar.showMessage(u"选择输入文件，单击'文件详情'可查看已选文件")
		self.setStatusBar(self.statusBar)
		self.set_ui()
		self.setUpdatesEnabled(True)
		self.groupbox1.updatesEnabled()
		if self.isVisible():
			self.check_input1()

		#self.statusBar.addWidget(self.probar)

	def set_ui(self):
		"""#界面排版布局"""
		layout = QGridLayout()
		layout.addWidget(self.groupbox1,0,0,2,2)
		layout.addWidget(self.bimg, 2, 0,1,2)
		layout.addWidget(self.groupbox2,3,0,1,2)
		layout.addWidget(self.groupbox3,4,0,2,2)
		layout.addWidget(self.btnok, 6, 1, 1, 1)
		layout.addWidget(self.btn,6,0,1,1)
		#layout.addWidget(self.statusBar,7,0,1,2)
		self.central.setLayout(layout)
		self.setCentralWidget(self.central)



	def check_input1(self):
		"""#判断用户选择的文件是够为空.

		"""
		is_ok=True
		pptlist=self.groupbox1.dataset.PPTnames
		if isinstance(pptlist,list) and len(pptlist):

			# if not self.groupbox2.isEnabled():
			# 	self.groupbox2.setEnabled(True)
			if not self.groupbox3.isEnabled():
				self.groupbox3.setEnabled(True)
			self.updatesEnabled()
		else:
			is_ok = False
			QMessageBox.warning(self, self.groupbox1.title(),(u"输入文件不允许为空"))
		return is_ok

	def Enableimg(self):
		"""#【图片设置】区激活设置与否.

		"""
		rec = self.check_input1()
		if rec:
			if self.groupbox2.isEnabled():
				self.groupbox2.setEnabled(False)
				self.bimg.setText(u"激活【图片设置】")
				self.statusBar.showMessage(u"输出文档(格式①或③))中需使用图片，请激活按钮")
			else:
				self.groupbox2.setEnabled(True)
				self.bimg.setText(u"取消【图片设置】")
				self.statusBar.showMessage(u"输出文档(格式②或④)中无需图片，取消激活加快运行速度")
	def printdetail(self):
		"""查看文件详情按钮设置.

		"""
		if self.check_input1():
			QMessageBox.about(self,u"已选",("\n".join(self.groupbox1.dataset.PPTnames)))

	def update_groupboxes(self):
		"""根据用户的输入情况，更新各项目框中获取到的值.

		"""
		refrsf=self.check_input1()
		if refrsf:
			# QMessageBox.about(self, u"已选", _(self.groupbox1.dataset.results))
			self.groupbox1.set()
			if self.groupbox2.isEnabled():
				self.groupbox2.set()

			self.groupbox3.set()
			reply = QMessageBox.question(self, 'Message',u"确认执行程序?",
											QMessageBox.Yes |QMessageBox.No,QMessageBox.Yes)
			if reply == QMessageBox.Yes:
				self.statusBar.clearMessage()
				self.statusBar.showMessage(u"正在收集参数......")

				self.getargs()#####收集用户的参数选择
			else:
				pass

	# args = self.args
	def getargs(self):
		"""获取输入面板的各项目的状态和最终的内容，以供后续处理.

		"""
		grps = [self.groupbox1.dataset, self.groupbox2.dataset, self.groupbox3.dataset]
		MainWindow.args["_PPTnames"], MainWindow.args["_WebTitle"] = grps[0]._PPTnames, grps[0]._WebTitle
		MainWindow.args["_outpath"], MainWindow.args["_outprefix"] = grps[0]._outpath, grps[0]._outprefix
		MainWindow.args["_ImagesDirName"], MainWindow.args["_imgtype"] = grps[1]._ImagesDirName, grps[1]._imgtype
		MainWindow.args["_newsize"]=grps[1]._newsize
		MainWindow.args["_txpdf"], MainWindow.args["_singppt"], MainWindow.args["_expind"] = grps[2]._txpdf, grps[
			2]._singppt, grps[2]._expind
		if self.groupbox2.isEnabled():
			MainWindow.args["_raw"], MainWindow.args["_resize"] = grps[1]._raw, grps[1]._resize
			MainWindow.args["_webFormat"], MainWindow.args["_langue"] = grps[2]._webFormat, grps[2]._langue
			MainWindow.args["_singimg"] = grps[2]._singimg
		else:
			MainWindow.args["_raw"], MainWindow.args["_resize"]=False,False
			MainWindow.args["_webFormat"], MainWindow.args["_langue"] =[],"null"
			MainWindow.args["_singimg"] =[]
		self.statusBar.showMessage(u"分配参数中...")
		self.dispatchargs()###按照输出的文件类型，分配参数####


	def dispatchargs(self):
		"""分配参数，按照【输入设置】，【图片设置】，【格式设置】及处理步骤分配参数值为全局变量.

		"""
		global singing,ImgTYPE,labAtit,Num,width,outdir,imageDir
		self.statusBar.showMessage(u"分配参数2...")
		Num,outdir,imageDir,width = len(MainWindow.args["_PPTnames"]), MainWindow.args["_outpath"], MainWindow.args["_ImagesDirName"], MainWindow.args[
			"_newsize"]
		singing= list(MainWindow.args["_singimg"])
		ImgTYPE= MainWindow.args["_imgtype"].lower()
		func3 = lambda x: [MainWindow.args["_outprefix"],MainWindow.args["_WebTitle"]] if Num == 1 else [func(x)] * 2
		labAtit = map(func3, MainWindow.args["_PPTnames"])  # label and title

		# for key, value in MainWindow.args.items():
		# 	print("key: ", key, "\nValue: ", value)
		self.runmultippt()
		# self.close()

	def runmultippt(self):
		"""处理用户输入的文件，for循环依次处理.

		功能：获取每个文件的绝对路径和文件编号，传给runsingppt(indexppt,pptfilename)进行处理

		"""
		global Num
		for indexppt,pptfilename in enumerate(self.groupbox1.dataset.PPTnames):
			#QCoreApplication.processEvents()
			self.statusBar.showMessage(u"正在处理第..."+str(indexppt+1)+u"个文件...")
			pptfilename=os.path.normpath(pptfilename)#v19修改，保证获取到的文件路径为使用本系统的常规、合法路径-normlization
			self.runsingppt(indexppt,pptfilename)
		QMessageBox.about(self, u"提示信息", (u"运行完毕，共运行" + str(Num) + u"个文件"))
		self.statusBar.showMessage(u"运行完毕，共运行" + str(Num) + u"个文件")
		# self.close()
		
	def runsingppt(self,numb, pptname):
		"""处理单个PPT文件的完整操作步骤，包括：

		STEP 1.pptx.open(filename=pptname, outDir0=outdir,label= Label)--打开ppt文件
		STEP 2.runfileppt()--导出文本型格式【txt/pdf/ppt】
		STEP 3:runimgppt(newtitle)--处理与图片相关的函数
		STEP 4:rundelfinal()---根据用户的选择，删除不需要生成的文件
		"""
		global labAtit
		Label,newtitle = labAtit[numb]
		MainWindow.shortnm = fun2(pptname)
		
		### STEP1: Open PPT FILE
		try:
			pptx.open(filename=pptname, outDir0=outdir,label= Label)
		except:
			#print e
			QMessageBox.warning(self,u"打开文件异常",
			                    (MainWindow.shortnm + u"打开失败\n 请检查文件状态，退出占用文件的相关程序，并重新执行"))
			pptx.closepres()



		#### STEP2: EXTRACT 【TXT/PDF/PPT】
		try:#txt pdf
			self.runfileppt()
		except:
			#print e
			QMessageBox.warning(self,u"TXT/PDF 导出异常", (MainWindow.shortnm + u"文本导出失败\n 重新尝试执行文件"))
			pptx.closepres()


		### STEP3:  DO WORKS RELATED TO IMAGEs【RESIZED_IMAGES/MERGE/HTML/XML】
		try:#image+html/xml
			if self.groupbox2.isEnabled():#如果【图片设置】区域被激活
			
				self.runimgppt(newtitle)##STEP 3-1: EXTRACRT【SLIDES/RESIZEIMAGES
			else:
				pass
		except :
			QMessageBox.warning(self, u"处理图片异常", (MainWindow.shortnm + u"处理图片异常\n 请重试执行"))
			# self.close()
		finally:
			pptx.closepres()#不管是否发生异常，保证文件执行完提取图片的操作后都被安全关闭

		### STEP4：根据用户选择输出的文件格式，自动清除不需要的生成文件
		try:#delete dirs unneeded
			self.rundelfinal()
		except:
				pass
	###文本型文件的操作函数：【txt.pdf.ppt】
	def runfileppt(self):# txttxpdf=[pdf,txt]
		"""导出txt/pdf/ppt文件.

		分别调用ppt_v12文件中easPPT类对象（pptx）的成员函数为：
			1.pptx.saveAs(Format="TXT")--Format为导出格式
			2.pptx.saveAs(Format="PDF")
			3.pptx.slid2PPT(substr=arg)-arg为用户输入的发布单张PPT的编号，list类型
		"""
		if 1 in MainWindow.args["_txpdf"]:  #1：导出TXT，执行ppt_v12中easyPPT的成员函数：saveAs
			self.statusBar.showMessage(MainWindow.shortnm+u"\n : 导出TXT......")
			pptx.saveAs(Format="TXT")
		if 0 in MainWindow.args["_txpdf"]:  #0： 导出PDF，执行ppt_v12中easyPPT的成员函数：saveAs
			self.statusBar.showMessage(MainWindow.shortnm+u"\n :  导出PDF.....")
			pptx.saveAs(Format="PDF")
		if MainWindow.args["_singppt"] and MainWindow.args["_expind"] not in ["", " "]:  # 发布单张PPT，执行成员函数：slid2PPT
			self.statusBar.showMessage(MainWindow.shortnm+u"\n :  导出PPT...")
			pptx.slid2PPT(substr=MainWindow.args["_expind"])


	def runimgppt(self, newtit):
		"""处理与导出图片相关的操作，eg:原图/缩放图/长图/html/xml文件等.

		传入参数：newtit --为操作面板中最终得到的网页标题栏中的内容
		以下步骤都需判断是在何种模式(raw:原图,resize:缩放图)下进行：
		STEP 1:调用ppt_v12文件中easPPT类对象（pptx）的成员函数pngExport：
			1.pptx.pngExport(imgDir=imageDir, newsize=width, imgtype=ImgTYPE.lower(), merge=Merge)
				功能：执行图片导出、缩放、拼接等步骤.
				参数：
					1.imgDir:输出中存放缩放图的文件夹名称，如未更改，为images
					2.newsize:设定输出缩放图的宽度
					3.imgtype:输出图片的格式：gif/jpg/png
					4.merge:取值为True/False,True表示输出长图

		STEP 2:调用Webppt_v11文件中的函数：
			1.webmod(lstdir, locimageDir, Kwarg)
				功能：根据面板输入(language type、htm or xml,raw or resize)，找到图片位置，生成html/xml文件
				参数：
				1.lstdir:list类型，元素为每张需使用图片的在系统中的绝对路径
				2.locimageDir:为存放使用图片的文件夹名称
				3.Kwarg:为发布htm/xml文件传递的相关参数
		"""
		global imageDir,width
		newtitle = newtit
		Merge = 1 if 2 in singing else 0  # 输出要求12代表要拼长图
		if MainWindow.args["_langue"]==None:#对应语言类型选择无，设置webFormat为空，即不发布
			MainWindow.args["_webFormat"]=[]
		if MainWindow.args["_webFormat"]==[]:
			MainWindow.args["_langue"]=None
		Kwarg = {"newtit": newtitle, "outpre": pptx.outfile_prefix,
				 "imgtype": ImgTYPE, "merge": Merge,
				 "langue": MainWindow.args["_langue"], "choice": MainWindow.args["_webFormat"]
				 }

		if MainWindow.args["_resize"] or 1 in singing:  # 输出要求1代输出缩放图，是单图中的原图或者勾选了缩放图模式
			self.statusBar.showMessage(MainWindow.shortnm+u"\n :  导出缩放图......")
			pptx.pngExport(imgDir=imageDir, newsize=width, imgtype=ImgTYPE.lower(), merge=Merge) #导出单张PPT，执行ppt_v12中easyPPT的成员函数：pngExport,传递参数newsize
			self.statusBar.showMessage(MainWindow.shortnm+"\n"+u":  拼接图片......")
			lstdir = pptx.outresizedir2
			locimageDir = imageDir
			if not MainWindow.args["_langue"] or len(MainWindow.args["_webFormat"]):
				self.statusBar.showMessage(MainWindow.shortnm+"\n"+u":  导出HTML/XML......")
				webmod(lstdir, locimageDir, Kwarg)
				
		if MainWindow.args["_raw"] or 0 in singing:  # 输出要求1代表选择原图模式 并且要求输出中选择要生成原图
			self.statusBar.showMessage(MainWindow.shortnm+"\n"+u": 导出原图......")
			pptx.pngExport(imgtype=ImgTYPE, merge=Merge)
			lstdir, rawimageDir = pptx.outslidir2, "Slides"
			Kwarg["outpre"] = pptx.outfile_prefix  # +u'_raw'
			if not MainWindow.args["_langue"] or len(MainWindow.args["_webFormat"]):
				self.statusBar.showMessage(MainWindow.shortnm + "\n" + u": 导出HTML/XML......")
				webmod(lstdir, rawimageDir, Kwarg)


	def rundelfinal(self):
		"""最后一步 删除用户不想输出的中间文件以及删除前进行判断.

		如果用户选择输出htm/xml文件，但不勾选对应模式下的单图输出，则最后依旧存放单图的文件夹
		分别调用ppt_v12文件中easPPT类对象（pptx）的成员函数delslides/delresize 删除原图集/缩放图集
		"""
		delflag = 1 not in MainWindow.args["_singimg"] and 0 not in MainWindow.args["_singimg"] and not len(MainWindow.args["_webFormat"])
		if 0 not in singing and delflag :  # 删除Slides
			pptx.delslides()
		if 1 not in singing and delflag:  # 删除resize
			pptx.delresize()




if __name__ == "__main__":
	#主程序入口
	if getattr(sys,'frozen',False):
		pyrealpath = sys._MEIPASS
	else:
		pyrealpath = os.path.split(os.path.realpath(__file__))[0]
	from guidata.qt.QtGui import QApplication
	app = QApplication(sys.argv)
	window = MainWindow()
	window.show()
	# try:
	app.exec_()#执行用户输入并等待按下apply
