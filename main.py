from tkinter import *
from tkinter import ttk
from tkinter import Menu
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText
from win32com.client import DispatchEx
import pickle, threading, queue, logging, tempfile, sys, os
import rasterer, BPimage, misc
from BPimage import listprinters, paperFormats, settingsFileName, isSettingsLoaded


logger = logging.getLogger(misc.loggerName)
logging.basicConfig(level=logging.INFO)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

Logo = resource_path("app.ico")


# TODO
# баги
    # DONE отсутствие компаса
    # DONE папка readonly
    # DONE выбор папки без файлов
    # DONE отсутствие нужного принтера


class QueueHandler(logging.Handler):
    """Class to send logging records to a queue
    It can be used from different threads
    The ConsoleUi class polls this queue to display records in a ScrolledText widget
    """
    # Example from Moshe Kaplan: https://gist.github.com/moshekaplan/c425f861de7bbf28ef06
    # (https://stackoverflow.com/questions/13318742/python-logging-to-tkinter-text-widget) is not thread safe!
    # See https://stackoverflow.com/questions/43909849/tkinter-python-crashes-on-new-thread-trying-to-log-on-main-thread

    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue

    def emit(self, record):
        self.log_queue.put(record)


class ConsoleUi:
    """Poll messages from a logging queue and display them in a scrolled text widget"""

    def __init__(self, frame):
        self.frame = frame
        # Create a ScrolledText wdiget
        self.scrolled_text = ScrolledText(frame, state='disabled', height=8, width=0)
        self.scrolled_text.grid(row=0, column=0, sticky=(N, S, W, E))
        self.scrolled_text.configure(font='TkFixedFont')
        self.scrolled_text.tag_config('INFO', foreground='black')
        self.scrolled_text.tag_config('DEBUG', foreground='gray')
        self.scrolled_text.tag_config('WARNING', foreground='DarkOrange2')
        self.scrolled_text.tag_config('ERROR', foreground='red')
        self.scrolled_text.tag_config('CRITICAL', foreground='red', underline=1)
        # Create a logging handler using a queue
        self.log_queue = queue.Queue()
        self.queue_handler = QueueHandler(self.log_queue)
        # formatter = logging.Formatter('%(asctime)s: %(message)s')
        # self.queue_handler.setFormatter(formatter)
        logger.addHandler(self.queue_handler)
        # Start polling messages from the queue
        self.frame.after(100, self.poll_log_queue)

    def display(self, record):
        msg = self.queue_handler.format(record)
        position = self.scrolled_text.vbar.get()
        self.scrolled_text.configure(state='normal')
        self.scrolled_text.insert(END, msg + '\n', record.levelname)
        self.scrolled_text.configure(state='disabled')
        # Autoscroll to the bottom
        if (position[1] == 1.0):
            self.scrolled_text.yview(END)

    def poll_log_queue(self):
        # Check every 100ms if there is a new message in the queue to display
        while True:
            try:
                record = self.log_queue.get(block=False)
            except queue.Empty:
                break
            else:
                self.display(record)
        self.frame.after(100, self.poll_log_queue)


def rasterAndPrint(inputFilePath, outputImagePath=None, Kompas=None):
    if outputImagePath is None:
        outputImagePath = os.path.join(os.environ['temp'], misc.getFileName(outputImagePath) + rasterer.OUTPUT_FORMAT)
    logger.log(logging.INFO, '>>>Растеризация\n%s' % os.path.split(inputFilePath)[1])
    isSuccess = rasterer.rasterKompasFile(Kompas, inputFilePath, outputImagePath)  # функция растеризации файла компаса
    if isSuccess == False:
        logger.log(logging.ERROR, ">>>Ошибка при растеризации файла\n%s" % os.path.split(inputFilePath)[1])
        return False
    pageCount = misc.getPageCount(inputFilePath, Kompas)
    if pageCount > 1: # если многостраничный документ
        successCount = pageCount
        for page in range(1, pageCount + 1):
            # вставляем в название файла номер страницы (%name%(%page%).jpg)
            fileWithPageNumber = outputImagePath[:-4] + '(' + str(page) + ')' + outputImagePath[-4:]
            logger.log(logging.INFO, '>>>Печать файла\n%s' % os.path.split(fileWithPageNumber)[1])
            isSuccess = BPimage.autoPrintImage(fileWithPageNumber)
            # os.remove(outputImagePath[:-4] + '(' + str(page) + ')' + outputImagePath[-4:])
            if isSuccess == False:
                logger.log(logging.ERROR, ">>>Ошибка при печати файла\n%s" % os.path.split(fileWithPageNumber)[1])
                successCount -= 1
                pass
        if successCount == 0:
            return False
    else:
        logger.log(logging.INFO, '>>>Печать файла\n%s' % os.path.split(outputImagePath)[1])
        isSuccess = BPimage.autoPrintImage(outputImagePath)  # функция печати файла компаса
        # os.remove(outputImagePath)
        if isSuccess == False:
            logger.log(logging.ERROR, ">>>Ошибка при печати файла\n%s" % os.path.split(outputImagePath)[1])
            return False
    return True


def runBatchPrint(path):

    printButton.config(state='disabled')
    if path == "Укажите путь к папке":
        printButton.config(state='normal')
        return False
    elif misc.isDirExists(path):
        fileList = misc.getFileList(path)
        KompasExtExists = False
        for filePath in fileList:
            if os.path.splitext(filePath)[1] in rasterer.KompasExt.keys():
                KompasExtExists = True
                break
        if KompasExtExists:
            pass
        else:
            misc.splitLongMessage('>>>В данной директории\nотсутствуют файлы компаса', path, lvl=logging.ERROR)
            printButton.config(state='normal')
            return False
    else:
        misc.splitLongMessage('>>>Несуществующая директория', path, lvl=logging.ERROR)
        printButton.config(state='normal')
        return False

    try:
        Kompas = DispatchEx("KOMPAS.Application.5")
        logger.log(logging.INFO, ">>>Запущен Компас")
    except:
        logger.log(logging.ERROR, ">>>Не удалось запустить Компас")
        printButton.config(state='normal')
        return False

    if misc.checkWritePath(path):
        tempDir = path
    else:
        tempDir = os.environ['temp']

    with tempfile.TemporaryDirectory(dir=tempDir, prefix="__kompas-batchPrint_") as outputFolder:
        misc.splitLongMessage('>>>Создана временная директория', outputFolder)
        for filePath in fileList:
            if os.path.splitext(filePath)[1] in rasterer.KompasExt.keys(): # обрабатываются только файлы с соответствующими расширениями
                outputFilePath = os.path.join(outputFolder, misc.getFileName(filePath) + rasterer.OUTPUT_FORMAT)
                # isSuccess = 
                rasterAndPrint(filePath, outputFilePath, Kompas)
                # if isSuccess == False:
                #     printButton.config(state='normal')
                #     return False

    logger.log(logging.INFO, '>>>Временная директория удалена')
    if not Kompas.Visible:
        logger.log(logging.INFO, '>>>Компас закрыт')
        Kompas.Quit()
    printButton.config(state='normal')
    return True


def exploreButtonClicked():
    folderPath = filedialog.askdirectory()
    folderPath = folderPath.replace('/', '\\')
    folderPathBox.delete(0, END)
    folderPathBox.insert(0, folderPath)


class printThread(threading.Thread):
   def __init__(self):
      threading.Thread.__init__(self)
   def run(self):
      runBatchPrint(folderPath.get())


def printButtonClicked():
    thread1 = printThread()
    thread1.start()


def saveSettingsButtonClicked():
    for i in range (len(paperFormats)):
        variable = paperFormatComboboxVar[i].get()
        if variable == "Принтер не выбран":
            BPimage.printersByPaperFormat['A'+str(i)] = None
        else:
            BPimage.printersByPaperFormat['A'+str(i)] = variable
    with open(settingsFileName, 'wb') as fileObject:
        pickle.dump(BPimage.printersByPaperFormat, fileObject)
    logger.log(logging.INFO, ">>>Настройки сохранены в файл\n%s" % settingsFileName)


def pathBoxTriggered(pathBoxString, trigger):
    if trigger == "focusin":
        if pathBoxString == "Укажите путь к папке":
            folderPathBox.delete(0, END)
            folderPathBox.config(fg="black", validate="all")
    elif trigger == "focusout":
        if pathBoxString == '':
            folderPathBox.insert(0, "Укажите путь к папке")
            folderPathBox.config(fg="gray", validate="all")
        else:
            folderPathBox.config(fg="black", validate="all")
    else:
        if pathBoxString != "Укажите путь к папке":
            folderPathBox.config(fg="black", validate="all")
    return True


# def tabSwitched(event):
#     print("tabSwitched")

window = Tk()
window.title("Пакетная печать Компас")
try:
    window.iconbitmap(Logo)
except:
    pass
# window.geometry('350x200')
window.resizable(FALSE, FALSE)

# tabs configure
tab_control = ttk.Notebook(window)
printTab = ttk.Frame(tab_control)
printersTab = ttk.Frame(tab_control)
# helpTab = ttk.Frame(tab_control)
tab_control.add(printTab, text="Печать")
tab_control.add(printersTab, text="Принтеры")
# tab_control.add(helpTab, text="Справка")
tab_control.pack(expand=1, fill='both')
# tab_control.bind("<<NotebookTabChanged>>", tabSwitched)


# 1 tab
printTab.grid_rowconfigure(0, minsize=5)
printTab.grid_columnconfigure(0, minsize=20)

exploreButton = Button(printTab, text="Обзор", command=exploreButtonClicked)
exploreButton.grid(column=1, row=1)

folderPath = StringVar()
validate = printTab.register(pathBoxTriggered)
folderPathBox = Entry(printTab, width=40, textvariable=folderPath, validate="all", validatecommand=(validate, '%P', '%V'))
folderPathBox.grid(column=2, row=1, columnspan=1, sticky='ns')
folderPathBox.insert(0, "Укажите путь к папке")
folderPathBox.config(fg="gray")
# folderPathBox.config(state="readonly", readonlybackground="white")

printTab.grid_rowconfigure(2, minsize=40)
printButton = Button(printTab, text="Печать", command=printButtonClicked)
printButton.grid(column=1, row=2, columnspan=2, sticky='nswe')

console_frame = ttk.Labelframe(printTab, text="Console")
console_frame.grid(column=1, row=3, columnspan=2, sticky='nswe')
console_frame.columnconfigure(0, weight=1)
console_frame.rowconfigure(0, weight=1)
console = ConsoleUi(console_frame)

printTab.grid_rowconfigure(4, minsize=10)
printTab.grid_columnconfigure(3, minsize=20)


# 2 tab
printersTab.grid_rowconfigure(0, minsize=5)
printersTab.grid_columnconfigure(0, minsize=20)

endRow = 1
paperFormatComboboxVar = [StringVar() for _ in range(len(paperFormats))] # .set("Принтер не выбран")
for paper, row in zip(paperFormats, range(1, len(paperFormats)+1)):
    paperFormatLabel = ttk.Label(printersTab, text=paper)
    paperFormatLabel.config(font=("TkDefaultFont", 12))
    paperFormatLabel.grid(column=1, row=row, padx=5, pady=5)

    paperFormatCombobox = ttk.Combobox(printersTab, textvariable=paperFormatComboboxVar[row-1], width=35)
    paperFormatCombobox['values'] = ["Принтер не выбран"] + listprinters()
    if BPimage.printersByPaperFormat[paper] is not None:
        paperFormatCombobox.current(paperFormatCombobox['values'].index(BPimage.printersByPaperFormat[paper]))
    else:
        paperFormatCombobox.current(0)
    paperFormatCombobox.config(state="readonly")
    paperFormatCombobox.grid(column=2, row=row, padx=5, pady=5, sticky='wes')
    endRow += 1

saveSettingsButton = Button(printersTab, text="Сохранить настройки", command=saveSettingsButtonClicked)
saveSettingsButton.grid(row=endRow, column=1, columnspan=2, padx=5, pady=5)
printersTab.grid_rowconfigure(endRow, minsize=20)
endRow += 1

printersTab.grid_rowconfigure(endRow, minsize=5)
printersTab.grid_columnconfigure(3, minsize=20)

if isSettingsLoaded == True:
    logger.log(logging.INFO, ">>>Настройки загружены из файла\n%s" % settingsFileName)
else:
    logger.log(logging.WARNING, ">>>Не обнаружен файл настроек\n%s" % settingsFileName)
    logger.log(logging.WARNING, ">>>Применены настройки\nпо умолчанию")


window.mainloop()