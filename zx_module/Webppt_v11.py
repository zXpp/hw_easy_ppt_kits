# -*- coding: utf-8 -*-
"""
Created on Mon Aug 07 01:25:07 

@author: 'z81022868'
本模块主要实现根据面板输入，将图片发布成htm/xml文件格式。

主函数入口：webmod(lstdir,imgdir,kwarg1={}):
调用子函数：pngs2web(mod,reldirlist=None,outpre=u"",langue=u"",newtit=u""):

XML_Template,HTMLTEMPLATE为本司需要的发布的通用格式，如后期需要其他格式模板，可进行替换。替换后为保证原功能正常使用，需注意：
	1.发布语言类型位置用{LANGUE}替换；如行18,31
	2.网页标题title用{TITLE}替换；如行19,31
	3.查找图片的相对路径用{IMAGE}替换；如行24,31
以上3点均为本函数查找定位、及替换的关键词，要保证程序能够找到，其余部分文字不做要求。
"""
import os
XMLTemplate=u"""<?xml version="1.0" encoding="UTF-8"?>
<!--Arbortext, Inc., 1988-2008, v.4002-->
<!DOCTYPE concept PUBLIC "-//OASIS//DTD DITA Concept//EN"
 "concept.dtd">
<?Pub Inc?>
<concept xml:lang="{LANGUE}">
<title>{TITLE}</title>
<prolog>
<metadata><keywords><keyword>KEYWORD</keyword></keywords></metadata>
</prolog>
<conbody>
<section><image align="center" href="{IMAGE}" placement="noblankline"></image></section><?Pub
Caret?>
</conbody>
</concept>
<?Pub *0000001262?>
"""
HTMLTEMPLTE=u"""<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html lang="{LANGUE}" xml:lang="{LANGUE}" charset="gbk"><head><meta http-equiv="Content-Type" content="text/html; charset=gb2312"><meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7"><title>{TITLE}</title></head><body style="margin: 0px;"><div align="center" font-size="0" display="block"><img src="{IMAGE}" display = "inline" ></img></br></div></body></html>"""

def pngs2web(mod,reldirlist=None,outpre=u"",langue=u"",newtit=u""):
	"""#具体对模板关键词进行定位、替换的函数.

	参数：
		mod:选择的发布格式--htm/xml
		reldirlist:list类型。元素为需使用的图片相对与生成文件的相对路径
		outpre:生成文件的除扩展名外的绝对路径,outpre是加了点的
		newtit:生成文件需显示的新的标题

	"""
	if mod in ["html","HTML","Html","htm"]:
		template=HTMLTEMPLTE
		begflag,endflag,code=u'<img',u'</img></br>','gbk'#pyrealpath=u"D:\\untar\\Panel\\Gudie"f1.read()
	elif mod in ["XML","xml","Xml"]:
		template=XMLTemplate
		begflag,endflag,code=u'<image',u'</image>','UTF-8'
	else:
		return

	content,titbeg,titend,new=template,u'<title>',u'</title>',[]
	matchtag=content[content.find(begflag):content.rfind(endflag)+len(endflag)]#寻找image标签对

	new.extend([matchtag.format(IMAGE=ele) for ele in reldirlist])#生成列中相邻的image标签对
	webout=template.replace(matchtag,u"".join(new))#进行替换

	outname="".join([outpre,'.',mod.lower()])
	out1=open(outname,'wb')
	out1.write(webout.format(LANGUE=langue,TITLE=newtit).encode(code))
	out1.close()
def webmod(lstdir,imgdir,kwarg1={}):
	"""#收集、整理htm/xml的相关参数，传递给pngsweb进行使用.

	"""
	func=lambda x:"".join([u'..\\',imgdir,u"\\",x]) if x.rsplit('.')[-1] in [kwarg1["imgtype"],kwarg1["imgtype"].upper()] else []
	kwarg2={"reldirlist":map(func,os.listdir(lstdir)),
	        "outpre":kwarg1["outpre"],"newtit":kwarg1["newtit"],"langue":kwarg1["langue"]}
	if 0 in kwarg1["choice"]:#html
		#mod,reldirlist=None,outpre=u"",newtit=u""
		pngs2web("htm",**kwarg2)
	if 1 in kwarg1["choice"]:#xml
		pngs2web("xml",**kwarg2)