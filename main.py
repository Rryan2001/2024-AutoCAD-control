import pythoncom
import win32com.client as win32
import math
import numpy as np
from pyautocad import Autocad, APoint

# connect to cad graph
acad = Autocad(create_if_not_exists=True)
# acad.prompt() to print in AutoCAD commend panel
acad.prompt("Hello, Autocad from Python")
# acad.doc.Name storges the latest opened CAD file
print(acad.doc.Name)


def vtpnt(x, y, z=0):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtfloat(lst):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def selectObject():
    global polylineCoord
    polylineCoord = []
    wincad = win32.Dispatch("AutoCAD.Application")
    doc = wincad.ActiveDocument
    msp = doc.ModelSpace
    try:
        doc.SelectionSets.Item("SS1").Delete()
    except:
        print("Delete selection failed")
    slt = doc.SelectionSets.Add("SS1")
    doc.Utility.Prompt("请选择多段线，右键结束\n")
    slt.SelectOnScreen()
    if slt.Count == 0:
        doc.Utility.Prompt("未选择对象！\n")
    else:
        entity = slt[0]
        name = entity.EntityName
        for i in range(len(entity.Coordinates)):
            polylineCoord.append(round(entity.Coordinates[i]))
    # Record the locations of the points on the selected polyline
    polylineCoord = [polylineCoord[x:x + 2] for x in range(0, len(polylineCoord), 4)]  #
    # print(type(polylineCoord))
    # print(len(polylineCoord))
    # print(polylineCoord)


# if __name__ == '__main__':
# selectObject()
selectObject()
# basicx=polylineCoord[0]
basicx = []
# basicy=polylineCoord[1]
basicy = []

for i in range(len(polylineCoord)):
    # print(polylineCoord[i][0])
    basicx.append(polylineCoord[i][0])  #
    basicy.append(polylineCoord[i][1])  #

basicx.sort()
basicy.sort()

#print(basicx)
#print(basicy)

# The collection of the coordinates 

# Record the max and min of X and Y coordinates, so that the coordinates of the inner framework can be calculated

framex = [basicx[0] + 500, basicx[-1] - 500]
framey = [basicy[0] + 300, basicy[-1] - 300]

#print(framex)
#print(framey)

point1 = APoint(framex[0], framey[0])  #
point2 = APoint(framex[0], framey[1])
point3 = APoint(framex[1], framey[1])
point4 = APoint(framex[1], framey[0])
# Build the outline of required framework

#
lineObj = acad.model.AddLine(point1, point2)
lineObj = acad.model.AddLine(point2, point3)
lineObj = acad.model.AddLine(point3, point4)
lineObj = acad.model.AddLine(point4, point1)

# draw extending lines on both directions
lineObj = acad.model.AddLine(APoint(basicx[0], framey[0]), point1)
# print('左下画好了')
lineObj = acad.model.AddLine(APoint(basicx[0], framey[1]), point2)
# print('左上画好了')
lineObj = acad.model.AddLine(APoint(basicx[-1], framey[1]), point3)
print(APoint(basicx[-1], framey[1]), point3)
# print('右上画好了')
lineObj = acad.model.AddLine(APoint(basicx[-1], framey[0]), point4)
# print('右下画好了')


# Polyline = acad.model.AddPolyline(framex[0],framey[0],0,framex[0],framey[1],0,framex[1],framey[1],0,framex[1],framey[0],0,framex[0],framey[0],0)


if 1200 <= abs(framey[1] - framey[0]) <= 2400:
    # select the middle point
    x1 = framex[0]

    y1 = 0.5 * (sum(framey))

    x2 = framex[1]

    line_start = APoint(x1, y1)
    line_end = APoint(x2, y1)
    lineObj = acad.model.AddLine(line_start, line_end)

elif 2400 < abs(framey[1] - framey[0]) <= 3600:
    x1 = framex[0]

    y1 = (2 * framey[0] + framey[1]) / 3  # Lower Y value

    y2 = (framey[0] + 2 * framey[1]) / 3  # Higher Y value

    x2 = framex[1]

    line_start1 = APoint(x1, y1)
    line_end1 = APoint(x2, y1)

    line_start2 = APoint(x1, y2)
    line_end2 = APoint(x2, y2)

    # 
    lineObj = acad.model.AddLine(line_start1, line_end1)  #

    lineObj = acad.model.AddLine(line_start2, line_end2)  #

    
elif abs(framey[1] - framey[0]) > 3600:
    x1 = framex[0]

    y1 = (3 * framey[0] + framey[1]) / 4  

    y2 = 0.5 * sum(framey)  

    y3 = (framey[0] + 3 * framey[1]) / 4  

    x2 = framex[1]

    line_start1 = APoint(x1, y1)
    line_end1 = APoint(x2, y1)

    line_start2 = APoint(x1, y2)
    line_end2 = APoint(x2, y2)

    line_start3 = APoint(x1, y3)
    line_end3 = APoint(x2, y3)

   
    lineObj = acad.model.AddLine(line_start1, line_end1)  #

    lineObj = acad.model.AddLine(line_start2, line_end2)  #

    lineObj = acad.model.AddLine(line_start3, line_end3)  #

#####################################################

# Adding annotions to the framework drawn
# 下方长度标注，从左到右，500的间距，模具横向长度，500的间距
ex_labelpoint1 = APoint(framex[0], basicy[0] - 500)  # 标注右侧的点
ex_labelpoint2 = APoint(basicx[0], basicy[0] - 500)  # 标注左侧的点#

text_position1 = APoint(0.5 * (ex_labelpoint1[0] + ex_labelpoint2[0]), basicy[0] - 500)  #
text_position2 = APoint(0.5 * sum(framex), basicy[0] - 500)
text_position3 = APoint(0.5 * (framex[1] + basicx[-1]), basicy[0] - 500)

dim1 = acad.model.AddDimAligned(ex_labelpoint1, ex_labelpoint2, text_position1)
dim1.ArrowheadSize = 30
dim1.TextHeight = 30
dim1.TextGap = 10
dim1.DecimalSeparator = "."

dim2 = acad.model.AddDimAligned(ex_labelpoint2, APoint(framex[1], basicy[0] - 500), text_position2)
dim2.ArrowheadSize = 30
dim2.TextHeight = 30
dim2.TextGap = 10
dim2.DecimalSeparator = "."  #

dim3 = acad.model.AddDimAligned(APoint(framex[1], basicy[0] - 500), APoint(basicx[-1], basicy[0] - 500), text_position3)
dim3.ArrowheadSize = 30
dim3.TextHeight = 30
dim3.TextGap = 10
dim3.DecimalSeparator = "."

# 右侧纵向的标注。此处只会做最上面和最下面的300长度的标注。其他的标注在之后的try 语法中实现

ex_labelpoint3 = APoint(basicx[-1] + 500, basicy[-1])  # 标注上方的点
ex_labelpoint4 = APoint(basicx[-1] + 500, framey[1])  # 标注下方的点

text_position4 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint3[1] + ex_labelpoint4[1]))

text_position5 = APoint((basicx[-1] + 500), 0.5 * (basicy[0] + framey[0]))

dim4 = acad.model.AddDimAligned(ex_labelpoint3, ex_labelpoint4, text_position4)
dim4.ArrowheadSize = 30
dim4.TextHeight = 30
dim4.TextGap = 10
dim4.DecimalSeparator = "."

dim5 = acad.model.AddDimAligned(APoint(basicx[-1] + 500, framey[0]), APoint(basicx[-1] + 500, basicy[0]),
                                text_position5)
dim5.ArrowheadSize = 30
dim5.TextHeight = 30
dim5.TextGap = 10
dim5.DecimalSeparator = "."

# block_name: str = 'SX-PC-BSH$0$SX-PC-SH$0$钢筋桁架_单元俯视'
block_name = 'SX-PC-BSH$0$SX-PC-SH$0$钢筋桁架_单元俯视'
# print(block_name.BlockScaling)
#
#
# block = acad.model.InsertBlock(point1, block_name,1.0, 1.0, 1.0, 0)
# 能用，但是应该需要写for 循环，以及需要知道块参考的长和宽来确定加载的坐标具体位置
# 后四位数字为xyx方向上的scale，以及rotation in radian

for i in range(int(abs(basicx[-1] - basicx[0]) / 200)):
    # block_upper_i = acad.model.InsertBlock(APoint(basicx[0] + i * 200, framey[1]), block_name, 1.0, 1.0, 1.0, 0)#在最上面的长线上加载
    #
    # block_lower_i = acad.model.InsertBlock(APoint(basicx[0] +i*200, framey[0]), block_name, 1.0, 1.0, 1.0, 0)#在最下面的长线上继续加载
   
    block_upper = acad.model.InsertBlock(APoint(basicx[0] + i * 200, framey[1]), block_name, 1.0, 1.0, 1.0,
                                         0)  # 在最上面的长线上加载

    block_lower = acad.model.InsertBlock(APoint(basicx[0] + i * 200, framey[0]), block_name, 1.0, 1.0, 1.0,
                                         0)  # 在最下面的长线上继续加载

    i += 1

block_upper = acad.model.InsertBlock(APoint(basicx[-1], framey[1]), block_name, -1.0, 1.0, 1.0, 0)

block_lower = acad.model.InsertBlock(APoint(basicx[-1], framey[0]), block_name, -1.0, 1.0, 1.0, 0)

for i in range(int(abs((framey[1] - 40) - (framey[0] + 40)) / 200)):
    block_left_i = acad.model.InsertBlock(APoint(framex[0], framey[0] + 40 + i * 200), block_name, 1.0, 1.0, 1.0,
                                          math.pi / 2)

    block_right_i = acad.model.InsertBlock(APoint(framex[1], framey[0] + 40 + i * 200), block_name, 1.0, 1.0, 1.0,
                                           math.pi / 2)

    i += 1

block_left = acad.model.InsertBlock(APoint(framex[0], framey[1] - 40), block_name, 1.0, 1.0, 1.0, -math.pi / 2)

block_right = acad.model.InsertBlock(APoint(framex[1], framey[1] - 40), block_name, 1.0, 1.0, 1.0, -math.pi / 2)


# 尝试判断第一条最低端的框架线是否存在。如果存在，则先画二等分线和底部的距离标注
try:
    y1  #
    for i in range(int(abs((framex[1] - 40) - (framex[0] + 40)) / 200)):
        block_y1_i = acad.model.InsertBlock(APoint(framex[0] + 40 + i * 200, y1), block_name, 1.0, 1.0, 1.0, 0)  #

        i += 1

    block_y1 = acad.model.InsertBlock(APoint(framex[1] - 40, y1), block_name, -1.0, 1.0, 1.0, 0)
    #
    ex_labelpoint5 = APoint(basicx[-1] + 500, y1)

    ex_labelpoint6 = APoint(basicx[-1] + 500, framey[0])

    text_position5 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint5[1] + ex_labelpoint6[1]))

    dim6 = acad.model.AddDimAligned(ex_labelpoint5, ex_labelpoint6, text_position5)
    dim6.ArrowheadSize = 30
    dim6.TextHeight = 30
    dim6.TextGap = 10
    dim6.DecimalSeparator = "."
# 如果中间没用额外的框架线，则直接创建从framey1 到framey0的距离标注
except NameError:
    ex_labelpoint5 = APoint(basicx[-1] + 500, framey[1])

    ex_labelpoint6 = APoint(basicx[-1] + 500, framey[0])

    text_position5 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint5[1] + ex_labelpoint6[1]))

    dim6 = acad.model.AddDimAligned(ex_labelpoint5, ex_labelpoint6, text_position5)
    dim6.ArrowheadSize = 30
    dim6.TextHeight = 30
    dim6.TextGap = 10
    dim6.DecimalSeparator = "."

try:
    y2

    for i in range(int(abs((framex[1] - 40) - (framex[0] + 40)) / 200)):
        block_y2_i = acad.model.InsertBlock(APoint(framex[0] + 40 + i * 200, y2), block_name, 1.0, 1.0, 1.0, 0)  #

        i += 1

    block_y2 = acad.model.InsertBlock(APoint(framex[1] - 40, y2), block_name, -1.0, 1.0, 1.0, 0)

    ex_labelpoint7 = APoint(basicx[-1] + 500, y1)

    ex_labelpoint8 = APoint(basicx[-1] + 500, y2)

    text_position6 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint7[1] + ex_labelpoint8[1]))

    dim7 = acad.model.AddDimAligned(ex_labelpoint7, ex_labelpoint8, text_position6)
    dim7.ArrowheadSize = 30
    dim7.TextHeight = 30
    dim7.TextGap = 10
    dim7.DecimalSeparator = "."
    # print('y2存在')
except NameError:

    ex_labelpoint7 = APoint(basicx[-1] + 500, framey[1])

    ex_labelpoint8 = APoint(basicx[-1] + 500, y1)

    text_position6 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint7[1] + ex_labelpoint8[1]))

    dim7 = acad.model.AddDimAligned(ex_labelpoint7, ex_labelpoint8, text_position6)
    dim7.ArrowheadSize = 30
    dim7.TextHeight = 30
    dim7.TextGap = 10
    dim7.DecimalSeparator = "."

    # print('y2不存在')


try:
    y3

    for i in range(int(abs((framex[1] - 40) - (framex[0] + 40)) / 200)):
        block_y3_i = acad.model.InsertBlock(APoint(framex[0] + 40 + i * 200, y3), block_name, 1.0, 1.0, 1.0, 0)

        i += 1
    block_y3 = acad.model.InsertBlock(APoint(framex[1] - 40, y3), block_name, -1.0, 1.0, 1.0, 0)  #

    ex_labelpoint9 = APoint(basicx[-1] + 500, y2)
    ex_labelpoint10 = APoint(basicx[-1] + 500, y3)
    ex_labelpoint11 = APoint(basicx[-1] + 500, framey[1])

    text_position7 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint9[1] + ex_labelpoint10[1]))
    text_position8 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint10[1] + ex_labelpoint11[1]))

    dim8 = acad.model.AddDimAligned(ex_labelpoint9, ex_labelpoint10, text_position7)
    dim8.ArrowheadSize = 30
    dim8.TextHeight = 30
    dim8.TextGap = 10
    dim8.DecimalSeparator = "."

    dim9 = acad.model.AddDimAligned(ex_labelpoint10, ex_labelpoint11, text_position8)
    dim9.ArrowheadSize = 30
    dim9.TextHeight = 30
    dim9.TextGap = 10
    dim9.DecimalSeparator = "."


except NameError:
    ex_labelpoint9 = APoint(basicx[-1] + 500, y2)

    ex_labelpoint10 = APoint(basicx[-1] + 500, framey[1])

    text_position7 = APoint(basicx[-1] + 500, 0.5 * (ex_labelpoint9[1] + ex_labelpoint10[1]))

    dim8 = acad.model.AddDimAligned(ex_labelpoint9, ex_labelpoint10, text_position7)
    dim8.ArrowheadSize = 30
    dim8.TextHeight = 30
    dim8.TextGap = 10
    dim8.DecimalSeparator = "."

    # print('y3不存在')



