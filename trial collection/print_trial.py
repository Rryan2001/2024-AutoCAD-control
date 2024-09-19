import pythoncom
import win32com.client as win32
import math
import numpy as np
from pyautocad import Autocad, APoint
import win32print
import time

# 打开cad文件
# 自动连接上cad，只要cad是开着的，就创建了一个<pyautocad.api.Autocad> 对象。这个对象连接最近打开的cad文件。
# 如果此时还没有打开cad，将会创建一个新的dwg文件，并自动开启cad软件
acad = Autocad(create_if_not_exists=True)
# acad.prompt() 用来在cad控制台中打印文字
acad.prompt("Hello, Autocad from Python")
# acad.doc.Name储存着cad最近打开的图形名
print(acad.doc.Name)


def vtpnt(x, y, z=0):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))


def vtobj(obj):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, obj)


def vtfloat(lst):
    return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, lst)


def selectObject():
    global polylineCoord
    global entity_ForDelete
    polylineCoord = []
    entity_ForDelete = []
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
        entity_ForDelete = slt[0]
        name = entity_ForDelete.EntityName
        for i in range(len(entity_ForDelete.Coordinates)):
            polylineCoord.append(round(entity_ForDelete.Coordinates[i]))
    # 坐标分组
    polylineCoord = [polylineCoord[x:x + 2] for x in range(0, len(polylineCoord), 4)]  #
    # print(type(polylineCoord))
    # print(len(polylineCoord))
    # print(polylineCoord)


# if __name__ == '__main__':
# selectObject()


#message = "请使用多段线/矩形工具画出打印范围"
#message = '请右键选中范围边界'
selectObject()

basicx = []
# basicy=polylineCoord[1]
basicy = []

for i in range(len(polylineCoord)):
    # print(polylineCoord[i][0])
    basicx.append(polylineCoord[i][0])  #
    basicy.append(polylineCoord[i][1])  #
# polylineCoord 现在是选定的多段线上的点的坐标的合集，每个元素都属于一个点的坐标
basicx.sort()
basicy.sort()


#打印范围为basicx 和 basicy 中的x 与y圈出的范围
entity_ForDelete.Delete()

acaddoc = acad.ActiveDocument
acadmod = acaddoc.ModelSpace
layout = acaddoc.layouts.item('Model')
plot = acaddoc.Plot

# 获取默认打印机
_PRINTER = win32print.GetDefaultPrinter()
_HPRINTER = win32print.OpenPrinter(_PRINTER)


# 打印样式设置函数
'''def PrinterStyleSetting():
    acaddoc.SetVariable('BACKGROUNDPLOT', 0)
    layout.ConfigName = 'RICOH MP C2011'
    layout.StyleSheet = 'monochrome.ctb'
    layout.PlotWithLineweights = False
    layout.CanonicalMediaName = 'A3'
    layout.PlotRotation = 1
    layout.CenterPlot = True
    layout.PlotWithPlotStyles = True
    layout.PlotHidden = False
    print(layout.GetPlotStyleTableNames()[-1])
    layout.PlotType = 4 
    '''


# 默认起始位置和绘图尺寸
DEFAULT_START_POSITION = (basicx[0], basicy[0])
DRAWING_SIZE = (basicx[1]-basicx[0], basicy[1]-basicy[0])
DRAWING_INTEND = 700


# 后台打印类
class BackPrint(object):
    _instance = None

    def __new__(cls, *args, **kw):
        if cls._instance is None:
            cls._instance = super(BackPrint, cls).__new__(cls)
        return cls._instance

    def __init__(self, PositionX, PositionY):
        self.x = PositionX
        self.y = PositionY

    #@staticmethod
    '''def APoint(x, y):
        return win32.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y))'''

    def run(self, Scale=1.0):
        #po1 = self.APoint(self.x * Scale - 1, self.y * Scale)
        #po2 = self.APoint(self.x * Scale - 1 + DRAWING_SIZE[0], self.y * Scale + DRAWING_SIZE[1])
        #layout.SetWindowToPlot(po1, po2)
        layout.SetWindowToPlot(basicx[1]-basicx[0], basicy[1]-basicy[0])

        #PrinterStyleSetting()
        plot.PlotToDevice()


# 打印任务类
class PrintTask:
    def __init__(self, maxPrintPositionArray, startPosition=(DEFAULT_START_POSITION[0], DEFAULT_START_POSITION[1])):
        self._PrinterStatus = 'Waiting'
        self.maxPrintPositionArray = maxPrintPositionArray
        self.printBasePointArray = []
        self.taskPoint = startPosition
        self.PrintingTaskNumber = 0

    def runtask(self):
        if not self.printBasePointArray:
            self.printBasePointArray = self.generalPrintBasePointArray(self.maxPrintPositionArray)

        for position in self.printBasePointArray:
            self.taskPoint = position
            current_task = BackPrint(*position)
            current_task.run()

            self.PrintingTaskNumber = len(win32print.EnumJobs(_HPRINTER, 0, -1, 1))

            while self.PrintingTaskNumber >= 5:
                time.sleep(1)
                self.PrintingTaskNumber = len(win32print.EnumJobs(_HPRINTER, 0, -1, 1))
            time.sleep(1)

    def ResumeTask(self):
        pass

    def generalPrintBasePointArray(self, maxPrintPositionArray):
        printBasePointArray = []
        next_drawing_xORy_intend = DRAWING_INTEND

        current_x = int((self.taskPoint[0] - 4) / DRAWING_INTEND) * DRAWING_INTEND + DEFAULT_START_POSITION[0]
        current_y = int((self.taskPoint[1] - 4) / DRAWING_INTEND) * DRAWING_INTEND + DEFAULT_START_POSITION[1]

        for position in maxPrintPositionArray:
            while current_x <= position + DEFAULT_START_POSITION[0]:
                printBasePointArray.append((current_x, current_y))
                current_x += next_drawing_xORy_intend
            current_x = DEFAULT_START_POSITION[0]
            current_y += next_drawing_xORy_intend
        return printBasePointArray

    def getTaskNumber(self):
        TaskNumber = self.PrintingTaskNumber
        try:
            TaskNumber = len(win32print.EnumJobs(_HPRINTER, 0, -1, 1))
            return TaskNumber
        except Exception as e:
            return TaskNumber


if __name__ == '__main__':
    #task = PrintTask([27895, ], (6194, 4))
    task = PrintTask([27895, ])
    task.runtask()


