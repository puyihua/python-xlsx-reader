#!usr/bin/env python
# -*- coding:utf-8 -*-
"""
this module was written for reading xlsx.
It  can return a dict that contains the information whether the col or row is hidden or not.
And it can read the location of pictures and the information of charts
"""
import re
import zipfile
try: 
	import xml.etree.cElementTree as ET
except ImportError: 
	import xml.etree.ElementTree as ET

def transferToABC(num):
	alphbet='ABCDEFGHIJKLMNOPQRSTUVWXYZ'
	num=int(num)
	result=''
	while num>0:
		if num == 26:
			result = 'Z' + result
			break
		else:
			result = alphbet[ num % 26-1] + result
			num /= 26
	return result

class readHidden(object):
	Hidden = {}
	def __init__(self,fileName,sheetNum):
		f = zipfile.ZipFile( fileName , 'r')
		xml = f.read('xl/worksheets/sheet%s.xml'%(sheetNum))
		root = ET.fromstring( xml )
		adict={}
		#root = tree.getroot()
		cols=root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cols")
		col=cols[0].findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}col")
		rows=root.findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheetData")
		row=rows[0].findall("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row")

		for i in col:
			for num in range(int(i.get("min")),int(i.get("max"))+1):
				adict[transferToABC(str(num))] = bool(i.get("hidden"))
		for i in row:
			adict[i.get("r")] = bool(i.get("hidden"))
		self.Hidden = adict



def readPic(fileName,sheetNum):
	f = zipfile.ZipFile( fileName , 'r')
	xml = f.read('xl/drawings/drawing%s.xml'%(sheetNum))
	root = ET.fromstring( xml )
	pics = root.findall("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}twoCellAnchor")
	local=""
	for pic in pics:
		fromCol=pic.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from")\
.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col").text
		fromRow=pic.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from")\
.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row").text
		toCol=pic.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}to")\
.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col").text
		toRow=pic.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}to")\
.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row").text
		[fromCol,fromRow,toCol,toRow] = map(transfer,[fromCol,fromRow,toCol,toRow])
		picfrom=transferToABC(fromCol)+fromRow
		picto=transferToABC(toCol)+toRow
		local+=picfrom+':'+picto+','
	return local

def transfer(string):
	return str(1+int(string))

def readChart(fileName):
	f = zipfile.ZipFile( fileName , 'r')
	pattern = re.compile(r'xl/charts/chart(.+)\.xml')
	templist=[]
	chartlist=[]
	for i in f.namelist():
		if pattern.match(i):
	 		templist.append(i)
	for i in templist:
		xml = f.read(i)
		temp={}
		# get type catalog value
		if re.findall(r"</c:layout><c:(.*?Chart)><c:",xml):
			temp["type"] = re.findall(r"</c:layout><c:(.*?Chart)><c:",xml)[0]
		elif re.findall(r"<c:layout/><c:(.*?Chart)><c:",xml):
			temp["type"] = re.findall(r"<c:layout/><c:(.*?Chart)><c:",xml)[0]
		else:
			temp["type"]= None

		temp["catalog"] = None
		find = re.findall(r"<c:cat>.+?<c:f>(.+?)</c:f>",xml)
		if find:
			temp["catalog"] = ''
			for i in find:
				temp["catalog"] += i + ','

		temp["value"] = None
		find = re.findall(r"<c:val>.+?<c:f>(.+?)</c:f>",xml)
		if find:
			temp["value"] = ''
			for i in find:
				temp["value"] += i + ','

		#get title
		tit=''
		regular = re.findall(r"<a:t>(.+?)</a:t>", xml)
		if regular:
			for i in regular:
				tit+=i
		else:
			tit = None
		temp["title"] = tit
		chartlist.append(temp)
	return chartlist

