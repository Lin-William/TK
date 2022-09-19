#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
    =============================
    Author: Lin.William
    Bio: For fjsdy only, make sure all graphics are in view.
    =============================
"""

import pythoncom
import math,time
import xlrd
import win32com.client


myCAD = win32com.client.Dispatch("AutoCAD.Application")
doc = myCAD.ActiveDocument
doc.Utility.Prompt("Hello! AutoCAD from pywin32.")
msp = doc.ModelSpace


def vtpnt(x, y, z=0):
    """坐标点转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):
    """转化为对象数组"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtfloat(lst):
    """列表转化为浮点数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def vtint(lst):
    """列表转化为整数"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I2, lst)


def vtvariant(lst):
    """列表转化为变体"""
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, lst)


myLayer = myCAD.ActiveDocument.Layers.Add("JustOneLastTime")
myCAD.ActiveDocument.ActiveLayer = myLayer
input = xlrd.open_workbook(r"E:\Stupefy\0.xls")
table = input.sheet_by_name("Sheet1")
myCAD.ActiveDocument.ActiveTextStyle = myCAD.ActiveDocument.TextStyles.Item("HZTXTE")

height_1 = 7.5
height_2 = 4
gap_1 = 20
gap_2 = 10
base_x = 0
base_y = 0

n_0 = 1  # 总平数量
n_1 = int(table.row_values(1)[0])  # 平面数量
n_2 = int(table.row_values(1)[1])  # 纵断数量
abbr = table.row_values(0)[1]  # 村镇名称缩写

# CAD对象	python对象名
# TEXT	AcDbText
# MTEXT	AcDbMText
# POINT	AcDbPoint
# LINE	AcDbLine
# LWPOLYLINE	AcDbPolyline
# ARC	AcDbArc
# CIRCLE	AcDbCircle
# ELLIPSE	AcDbEllipse
# SPLINE	AcDbSpline
# HATCH	AcDbHatch
# MLINE	AcDbMline
# TABLE	AcDbTable


try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    pass
slt = doc.SelectionSets.Add("SS1")
FilterData = vtvariant(["<or", "circle", "lwpolyline", "hatch", "text", "line", "or>"])
FilterType = vtint([-4, 0, 0, 0, 0, 0, -4])
slt.Select(1, vtpnt(0, 1000), vtpnt(594, 1940), FilterType, FilterData)
for obj in slt:
    doc.ObjectIdToObject(obj.ObjectID).copy().Move(vtpnt(0, 1000), vtpnt(0, 0))
time.sleep(1)
# 选取平面部分
try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    pass
slt = doc.SelectionSets.Add("SS1")
FilterData = vtvariant(["<or", "circle", "lwpolyline", "hatch", "text", "line", "or>"])
FilterType = vtint([-4, 0, 0, 0, 0, 0, -4])
slt.Select(1, vtpnt(0, 0), vtpnt(594, 420), FilterType, FilterData)

numberOfRows = 1
numberOfColumns = n_1
numberOfLevels = 1
distanceBwtnRows = 0
distanceBwtnColumns = 594 + gap_1
distanceBwtnLevels = 0
for obj in slt:
    try:
        retObj = doc.ObjectIdToObject(obj.ObjectID).ArrayRectangular(
            numberOfRows,
            numberOfColumns,
            numberOfLevels,
            distanceBwtnRows,
            distanceBwtnColumns,
            distanceBwtnLevels,
        )
    except:
        pass

# 纵断部分

try:
    doc.SelectionSets.Item("SS1").Delete()
except:
    pass
slt = doc.SelectionSets.Add("SS1")
FilterData = vtvariant(["<or", "circle", "lwpolyline", "hatch", "text", "line", "or>"])
FilterType = vtint([-4, 0, 0, 0, 0, 0, -4])
slt.Select(1, vtpnt(0, 520), vtpnt(594, 940), FilterType, FilterData)

for obj in slt:
    try:
        retObj = doc.ObjectIdToObject(obj.ObjectID).ArrayRectangular(
            1, n_2, 1, 0, 594 + gap_2, 0
        )
    except:
        pass

# 笔记
# doc.ObjectIdToObject(obj.ObjectID)  图元ID转化为对应的图元
# 获取选择集中的对象ID转化为对应图元并记录在sltObj
# sltObj=[]
# for obj in slt:
#     sltObj.append(doc.ObjectIdToObject(obj.ObjectID))
# print(sltObj)

# for obj in slt:
#     doc.ObjectIdToObject(obj.ObjectID).Copy().Move(vtpnt(0,0),vtpnt(0,-500))

# 填充独立图名
for i in range(1, n_1 + 1):
    txtString = (
        table.row_values(0)[0] + "配水管线平面图(" + str(i) + "/" + str(n_1) + ")     1:1000"
    )
    x = 225 + base_x + (594 + gap_1) * (i - 1)
    y = 40 + base_y
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_1)

    txtString = table.row_values(0)[0] + "配水管线平面图(" + str(i) + "/" + str(n_1) + ")"
    x = 529 + base_x + (594 + gap_1) * (i - 1)
    y = 24 + base_y
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_2)
    txt.Alignment = 10
    txt.TextAlignmentPoint = insertPnt

    k = str(n_0 + i)
    txtString = "W2022212-C832-" + abbr + "-" + k.zfill(2)
    x = 519 + base_x + (594 + gap_1) * (i - 1)
    y = 14 + base_y
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_2)
    txt.Alignment = 10
    txt.TextAlignmentPoint = insertPnt

for i in range(1, n_2 + 1):
    txtString = (
        table.row_values(0)[0] + "配水管线纵断面图(" + str(i) + "/" + str(n_2) + ")     1:1000"
    )
    x = 225 + base_x + (594 + gap_2) * (i - 1)
    y = 40 + base_y + 520
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_1)

    txtString = table.row_values(0)[0] + "配水管线纵断面图(" + str(i) + "/" + str(n_2) + ")"
    x = 529 + base_x + (594 + gap_2) * (i - 1)
    y = 24 + base_y + 520
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_2)
    txt.Alignment = 10
    txt.TextAlignmentPoint = insertPnt

    k = str(n_0 + n_1 + i)
    txtString = "W2022212-C832-" + abbr + "-" + k.zfill(2)
    x = 519 + base_x + (594 + gap_2) * (i - 1)
    y = 14 + base_y + 520
    insertPnt = vtpnt(x, y)
    txt = myCAD.ActiveDocument.ModelSpace.AddText(txtString, insertPnt, height_2)
    txt.Alignment = 10
    txt.TextAlignmentPoint = insertPnt
