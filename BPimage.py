# -*-coding: utf-8 -*

import win32print, sys
import win32ui, win32gui
import win32con, pywintypes
import os, logging, pickle
import misc
from misc import loggerName
from PIL import Image, ImageWin
from math import ceil
#
logger = logging.getLogger(loggerName)
logging.basicConfig(level=logging.INFO)
#
# Constants for GetDeviceCaps
#
#
win32con.DMPAPER_A1 = 271
win32con.DMPAPER_A1_OVERSIZE = 621
win32con.DMPAPER_A2_OVERSIZE = 620

# HORZRES / VERTRES = printable area
#
HORZRES = 8
VERTRES = 10
#
# LOGPIXELS = dots per inch
#
LOGPIXELSX = 88
LOGPIXELSY = 90
#
# PHYSICALWIDTH/HEIGHT = total area
#
PHYSICALWIDTH = 110
PHYSICALHEIGHT = 111
#
# PHYSICALOFFSETX/Y = left / top margin
#
PHYSICALOFFSETX = 112
PHYSICALOFFSETY = 113


KOMPAS_EXTENSIONS = (".cdw", ".spw")
paperFormats = ("A0", "A1", "A2", "A3", "A4", "A5")
settingsFileName = "settings.pkl"
try:
    with open(settingsFileName, 'rb') as fileObject:
        printersByPaperFormat = pickle.load(fileObject)
        isSettingsLoaded = True
except IOError:
    isSettingsLoaded = False
    printersByPaperFormat = {
            "A0": None,
            "A1": None,
            "A2": None,
            "A3": None,
            "A4": None,
            "A5": None
            }



prdict = None


def build_dict():
    global prdict
    lst = win32print.EnumPrinters(
        win32print.PRINTER_ENUM_CONNECTIONS
        + win32print.PRINTER_ENUM_LOCAL)
    prdict = {}
    for flags, description, name, comment in lst:
        prdict[name] = {}
        prdict[name]["flags"] = flags
        prdict[name]["description"] = description
        prdict[name]["comment"] = comment


def listprinters():
    dft = win32print.GetDefaultPrinter()
    if prdict is None:
        build_dict()
    keys = prdict.keys()
    rc = [ dft ]
    for k in keys:
        if k != dft:
            rc.append(k)
    return rc


def desc(name):
    if prdict == None:
        listprinters()
    return prdict[name]


def getImageSizeInMM(img):
    """
    Функция определяет физический размер требуемый для печати изходя из DPI и размера изображения в пикселях.
    :param img: объект, получаемый из PIL.Image
    :return: list(width, height) размер изображения в милиметрах
    """
    dpi = img.info['dpi']
    return list(map(lambda size, lDPI: int(ceil(float(size) / lDPI * 25.4)), img.size, dpi))


def getImagePaperFormat(img):
    """
    Функция определяет формат бумаги, требуемый для печати в полный размер. Исходные данные для определения формата
    бумаги - размер изображения в пикселях и количество точек на дюйм(DPI).
    :param img: объект, получаемый из PIL.Image
    :return: формат бумаги в формате A0,A1,A2,A3,A4. None при неудаче при определении формата
    """
    A = [(841, 1189)]
    OFFSET = 25
    for x in range(1, 5):
        prvHeight = A[x - 1][1]
        prvWidth = A[x - 1][0]
        A.append((prvHeight / 2, prvWidth))
        # В результате в массиве A хранятся размеры листов от A0 (под адресом A[0]) до A5 (под адресом A[5])
        # Для ускорения работы программы можно записать уже посчитанные значения

    imgSizeMM = getImageSizeInMM(img)
    if imgSizeMM[0] > imgSizeMM[1]:
        imgSizeMM.reverse()
    for pStandard in A:
            delta = list(map(lambda imgSize, pSize: abs(imgSize - pSize) < OFFSET, imgSizeMM, pStandard))
            if delta.count(True) == len(delta):
                imgSizeMM = pStandard
                break
    try:
        paperFormat = 'A%i' % A.index(tuple(imgSizeMM))
    except ValueError:
        return None
    return paperFormat


def printImage(printer_name, bmp, paperSize=win32con.DMPAPER_A4, tray=None, jobTitle="Untitled"):

    """
    Функция печати изображения по входным параметрам:
    :param printer_name: имя принтера в windows
    :param img:
    :param paperSize:
    :param tray:
    """
    PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
    hprinter = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
    properties = win32print.GetPrinter(hprinter, 2)
    pr = properties['pDevMode']
    if (bmp.size[1] < bmp.size[0]):
        pr.Orientation = pr.Orientation = win32con.DMORIENT_LANDSCAPE # 2
    else:
        pr.Orientation = pr.Orientation = win32con.DMORIENT_PORTRAIT # 1
    pr.PaperSize = paperSize
    # if tray:
    #     pr.DefaultSource = tray
    properties['pDevMode'] = pr
    try:
        win32print.SetPrinter(hprinter, 2, properties, 0)
    except Exception:
        pass # TODO выдача ошибки
        sys.exit()
    #
    # You can only write a Device-independent bitmap
    #  directly to a Windows device context; therefore
    #  we need (for ease) to use the Python Imaging
    #  Library to manipulate the image.
    #
    # Create a device context from a named printer
    #  and assess the printable size of the paper.
    #
    gDC = win32gui.CreateDC("WINSPOOL", printer_name, pr)
    hDC = win32ui.CreateDCFromHandle(gDC)
    printable_area = hDC.GetDeviceCaps(HORZRES), hDC.GetDeviceCaps(VERTRES)
    printer_size = hDC.GetDeviceCaps(PHYSICALWIDTH), hDC.GetDeviceCaps(PHYSICALHEIGHT)
    printer_margins = hDC.GetDeviceCaps(PHYSICALOFFSETX), hDC.GetDeviceCaps(PHYSICALOFFSETY)
    #
    #  Work out how much to multiply
    #  each pixel by to get it as big as possible on
    #  the page without distorting.
    #
    ratios = [1.0 * printable_area[0] / bmp.size[0], 1.0 * printable_area[1] / bmp.size[1]]
    scale = min (ratios) * 0.99
    #
    # Start the print job, and draw the bitmap to
    #  the printer device at the scaled size.
    #
    hDC.StartDoc(jobTitle)
    hDC.StartPage()
    # bmp.mode = "RGB"
    dib = ImageWin.Dib(bmp, bmp.size)
    scaled_width, scaled_height = [int(scale * i) for i in bmp.size]
    x1 = int((printer_size[0] - scaled_width) / 2) - printer_margins[0]
    y1 = int((printer_size[1] - scaled_height) / 2) - printer_margins[1]
    x2 = x1 + scaled_width
    y2 = y1 + scaled_height
    dib.draw(hDC.GetHandleOutput(), (x1, y1, x2, y2))

    hDC.EndPage()
    hDC.EndDoc()
    hDC.DeleteDC()

# эта функция уже не нужна, но пока просто заккоментирую, вдруг пригодится
# def getPrinterByPaperFormat(max_PaperFormat):
#     for printer in PRINTERS.items():
#         if max_PaperFormat in printer[1]['PaperFormats']:
#             return PRINTERS[printer[0]]
#     return {} # TODO изменить логику


def autoPrintImage(file_name):
    try:
        img = Image.open(file_name)
    except IOError:
        pass
        return False # TODO выдать лог с ошибкой
    frm = getImagePaperFormat(img)
    printer = printersByPaperFormat[frm]
    if printer is None:
        return False
    curFile = os.path.split(file_name)[1]
    frm = getattr(win32con, "DMPAPER_" + frm)
    printImage(printer, img, frm, tray=None, jobTitle=curFile)
    return True
