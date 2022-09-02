# -*-coding: utf-8 -*-
import os, sys, logging


loggerName = "console"
logger = logging.getLogger(loggerName)
logging.basicConfig(level=logging.INFO)

KompasExt = {".cdw": "Document2D",
             ".spw": "SpcDocument"}


class KompasException(Exception):
    pass


class FileTypeException(Exception):
    pass


class PathException(Exception):
    pass


def isDirExists(path):
    if os.path.isdir(path):
        return True
    else:
        return False


def getFileName(filePath):
    fileNameExt = os.path.basename(filePath)
    return os.path.splitext(fileNameExt)[0]


def getFileList(path):
    files = []
    for fileOrDir in os.listdir(path):
        if os.path.isfile(path + '\\' + fileOrDir):
            files.append(path + '\\' + fileOrDir)
    return files


def splitDirToFolders(path):
    allparts = []
    while 1:
        parts = os.path.split(path)
        if parts[0] == path:  # sentinel for absolute paths
            allparts.insert(0, parts[0])
            break
        elif parts[1] == path: # sentinel for relative paths
            allparts.insert(0, parts[1])
            break
        else:
            path = parts[0]
            allparts.insert(0, parts[1])
    return allparts


def splitLongMessage(startMessage, path, endMessage=None, lvl=logging.INFO):
    logger.log(lvl, startMessage)
    pathSplitted = splitDirToFolders(path)
    pathSplitted[0] = pathSplitted[0][:-1]
    logMessage = pathSplitted[0]
    del pathSplitted[0]
    for folder in pathSplitted:
        if len(logMessage) + len('\\') + len(folder) > 32:
            logMessage += '\\'
            logger.log(lvl, logMessage)
            logMessage = ''
        logMessage += '\\' + folder
    if logMessage != '':
        if logMessage[-1] == ':':
            logMessage += '\\'
        logger.log(lvl, logMessage)
    if endMessage is not None:
        logger.log(lvl, endMessage)


def checkWritePath(path):
    try:
        open(path + "/dummy", "w")
    except IOError:  # если не удалось записать файл в папку
        return False
    else:
        try:
            os.remove(path + "/dummy")
        except (IOError, WindowsError):
            pass
    return True


def getPageCount(inputFilePath, Kompas=None):
    if Kompas is None:
        raise KompasException("A COM Kompas object is not given")
    if not os.path.exists(inputFilePath):
        raise PathException("file '%s' don't exists" % inputFilePath)
    extension = os.path.splitext(inputFilePath)[1]
    docType = KompasExt.get(extension, None)
    if not docType:
        raise FileTypeException("Extension %s not supported" % extension)

    doc = getattr(Kompas, docType)
    doc.ksOpenDocument(inputFilePath, 1)
    if docType == "Document2D":
        pageCount = doc.ksGetDocumentPagesCount()
    elif docType == "SpcDocument":
        pageCount = doc.ksGetSpcDocumentPagesCount()
    doc.ksCloseDocument()
    return pageCount