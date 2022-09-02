# -*-coding: utf-8 -*-
# from win32com.client import Dispatch
from win32com.universal import com_error
import os, sys, logging
from misc import loggerName, KompasExt

logger = logging.getLogger(loggerName)

OUTPUT_FORMAT = ".jpg"


class RastererException(Exception):
    pass


class FileTypeException(RastererException):
    pass


class PathException(RastererException):
    pass


def rasterKompasFile(kompasObject, inputFilePath, outputFilePath):
    """
    Функция конвертирования файла КОМПАС в формат jpg c DPI 300
    :param inputFile:
    :param outputFilePath:
    :raise:
    """
    extension = os.path.splitext(inputFilePath)[1]
    docType = KompasExt.get(extension, None)
    if not docType:
        raise FileTypeException("Extension %s not supported" % extension)
    if not os.path.exists(inputFilePath):
        raise PathException("file '%s' don't exists" % inputFilePath)
    outputPath = os.path.split(outputFilePath)[0]
    if not os.path.isdir(outputPath):
        raise PathException("path '%s' don't exists" % outputPath)

    doc = getattr(kompasObject, docType)
    doc.ksOpenDocument(inputFilePath, 1) # (путь к документу, режим открытия)
    # 0 – видимый режим
    # 1 – невидимый режим
    # 3 – видимый режим без синхронизации со сборочным чертежом
    # 4 – невидимый режим без синхронизации со сборочным чертежом

    RasterFormatParam = doc.RasterFormatParam() # доступ к интерфейсу параметров
    RasterFormatParam.Init() # обнуляет значения всех свойств интерфейса

    # RasterFormatParam.Format = 0  # bmp - слишком большой размер изображений
    # RasterFormatParam.Format = 1  # gif - не тестировалось
    RasterFormatParam.Format = 2  # jpg - работает
    # RasterFormatParam.Format = 3  # png - не тестировалось
    # RasterFormatParam.Format = 4  # tiff - не работает с многостраничными документами
    # RasterFormatParam.Format = 5  # tga - ?
    # RasterFormatParam.Format = 6  # pcx - ?
    # RasterFormatParam.Format = 16  # wmf (не поддерживается) - ?
    # RasterFormatParam.Format = 17  # emf - ?

    # параметры .ExtResolution и .ColorBPP не менять, иначе изображение будет неверно распечатано
    # параметр .ColorType менять можно, допустимые параметры: 1 (ч/б изображение), 2, 4, 8, 16, 24, 32
    bitDepth = 24 
    RasterFormatParam.ColorType = 1 # глубина цвета при растеризации документа
    RasterFormatParam.ColorBPP = bitDepth # глубина цвета сохранённого растеризованного изображения
    RasterFormatParam.ExtResolution = 300 # dpi изображения

    # RasterFormatParam.MultiPageOutput = False # как показывает практика, данный парамер нифига не работает
    # .multiPageOutput – признак сохранения листов документа в одном файле.
    # Если значение данного свойства равно true, то все листы документа сохраняются в одном файле.
    # Если же значение этого свойства равно false, то листы сохраняются в отдельных файлах.
    # Данное свойство используется только для формата TIFF.
    # Для других форматов листы сохраняются в отдельные файлы всегда.

    RasterFormatParam.OnlyThinLine = False # признак вывода в тонких линиях.
    # Если значение этого свойства равно true, то содержимое документа выводится только в тонких линиях.
    # Если же значение этого свойства равно false, то при выводе документа используются линии, установленные для объектов.

    RasterFormatParam.RangeIndex = 0  # all pages
    # 0 – все листы;
    # 1 – нечетные листы;
    # 2 – четные листы.

    isSuccess = doc.SaveAsToRasterFormat(outputFilePath, RasterFormatParam)
    return isSuccess
    # doc.ksCloseDocument()


# def __rasterKompasFile(doc, outputFilePath, RasterFormatParam):
#     success = doc.SaveAsToRasterFormat(outputFilePath, RasterFormatParam)
#     if not success:
#         raise RastererException('Error during rasterization of %s' % outputFilePath)