import PyPDF2
import subprocess
import glob
import tempfile
import os, re, sys
import platform
from threads import Worker, WorkerSignals
from MainWindow import *
from PyQt5.QtWidgets import QApplication, QMainWindow, QMessageBox, QFileDialog, QCheckBox
from datetime import datetime

# Global variables
pdf_file = ""
# What OS are we using?
osName = platform.system()
match osName:
  case 'Linux':
    True
  case 'Windows':
    os.altsep = '/'
    import win32print
  case _:
    show_msg(QMessageBox.Critical,f"Sorry, couldn't figure out which OS you're using. platform.system() reports >{osName}<")
    exit()

def empty(str):
  if re.search(r'^\s*$',str):
    return True
  return False

def thread_complete():
  """ What gets called when the thread has finished """
  # All done. Clear the form fields and tell the user
  show_msg(QMessageBox.Information,"All done!")
  ui.txt_pdf_file.setText("")
  ui.txt_num_pages.setText("")

def process_results():
  """ Process the results of the thread """
  True

def log(msg=""):
  myPath = os.path.dirname(sys.argv[0])
  myBasename = os.path.splitext(os.path.basename(sys.argv[0]))[0]
  with(open(os.path.join(myPath,f"{myBasename}.log"), "a")) as lh:
    print(datetime.now().strftime("%a %d %b %Y %H:%M:%S ") + msg, file=lh)

def show_msg(msg_type, msg):
  msg_box = QMessageBox()
  msg_box.setWindowTitle("")
  msg_box.setText(msg)
  msg_box.setIcon(msg_type)
  msg_box.exec_()

def progress_fn(msg):
  """ 
  The function that updates (perhaps periodically) a status display. And before you ask
  I'm not using a progress bar because the MFP doesn't tell us how much data is coming so
  a progress bar is pointless.
  """
  ui.statusbar.showMessage(msg)

def file_select():
  """
  Open file sselect dialog
  """
  global pdf_file
  pdf_file, filter = QFileDialog.getOpenFileName(MainWindow, "Select PDF file to process", '.', "PDF (*.pdf)")
  ui.txt_pdf_file.setText(pdf_file)

def GoClicked():
  # global osName
        
  msg = ""
  if empty(ui.txt_pdf_file.text()):
    msg += "You must select a PDF file to process.\n"
  if empty(ui.txt_num_pages.text()) or ui.txt_num_pages.text() == "0":
    msg += "You must specify the number of pages per file.\n"
  if empty(ui.comboBox.currentText()):
    msg += "You must specify the printer to use.\n"

  if len(msg) > 0:
    show_msg(QMessageBox.Critical,msg)
    return

  pdf_file_name = ui.txt_pdf_file.text()
  ppf = int(ui.txt_num_pages.text())
  printer = ui.comboBox.currentText()

    # Start a new thread to do the actual download
  pdf_thread = Worker(process_PDF, pdf_file_name=pdf_file, ppf=ppf, printer=printer)
  pdf_thread.signals.result.connect(process_results)
  pdf_thread.signals.progress.connect(progress_fn)
  pdf_thread.signals.finished.connect(thread_complete)
  threadpool.start(pdf_thread)

def process_PDF(progress_callback, pdf_file_name, ppf, printer):
  """
  This is where we split the specified file and print them as seperate jobs.
  """
  prn_cmd = ""
  progress_callback.emit(f"Spliting {pdf_file_name} into {ppf}-page files")
  log(f"Spliting {pdf_file_name} into {ppf}-page files")
  
  # Start with a temp directory
  with tempfile.TemporaryDirectory() as output_dir:
    # Open the PDF file in read-binary mode
    with open(pdf_file_name, 'rb') as pdf_file:
      # Create a PDF reader object
      pdf_reader = PyPDF2.PdfReader(pdf_file)

      # Get the total number of pages
      total_pages = len(pdf_reader.pages)

      # Split the PDF file into {ppf}-page files
      for i in range(0, total_pages, ppf):
        # Create a new PDF writer object
        pdf_writer = PyPDF2.PdfWriter()

        # Add the 3-page range to the new PDF file
        for j in range(i, min(i+ppf, total_pages)):
          page = pdf_reader.pages[j]
          pdf_writer.add_page(page)

        # Save the new PDF file with the 3-page range
        with open(f'{output_dir}/output_file_{i//ppf+1}.pdf', 'wb') as output_file:
          pdf_writer.write(output_file)

    progress_callback.emit(f"Sending files to printer {printer}")
    log(f"Sending files to printer {printer}")
    file_pattern = "output_file*.pdf"
    matching_files = glob.glob(os.path.join(output_dir,file_pattern))
    

    for file in matching_files:
      # Print the file
      match osName:
        case "Linux":
          prn_cmd = f"lpr -P'{printer}' {file}"
        case "Windows":
          prn_cmd = f'pdftoprinter {file} "{printer}"'
          # if re.match(r'\\\\', printer):
          #   prn_cmd = f'type {file} >"{printer}"'
          # else:
          #   prn_cmd = f'print /D:"{printer}" {file}'

      log(prn_cmd)
      print(prn_cmd)
      output = subprocess.run(prn_cmd, shell=True, text=True, capture_output=True)
      if (not empty(output.stderr)):
        progress_callback.emit(f"Problem executing {prn_cmd}: {output.stderr}")
        log(f"Problem executing {prn_cmd}: {output.stderr}")

########### Main #############
if __name__ == '__main__':

  # Set up display window
  app = QApplication(sys.argv)
  MainWindow = QMainWindow()
  ui = Ui_MainWindow()
  ui.setupUi(MainWindow)
  screen = app.primaryScreen()
  x = int((screen.size().width() - 760) / 2)
  y = int((screen.size().height() - 312) / 2)
  MainWindow.setWindowTitle('splitPDF')
  MainWindow.move(x, y)

  # Retrieve list of available printers. Access detailes with printers[n]["pPrinterName"]
  printers = []
  if osName == "Linux":
    lpstat = subprocess.run(['lpstat', '-a'], capture_output = True, text=True) # List printers that are accepting jobs
    printers = [x.split()[0] for x in lpstat.stdout.splitlines()]
  elif osName == "Windows":
    # first see if pdftoprinter is available
    import shutil
    pdftoprinter_path = shutil.which("pdftoprinter.exe")
    if pdftoprinter_path is None:
      log("Can't find PDFToPrinter.exe on the PATH. Please install it then try again.")
      show_msg(QMessageBox.Critical,"Can't find PDFToPrinter.exe on the PATH. Please install it then try again.")
      sys.exit()

    printer_objs = win32print.EnumPrinters(
      (win32print.PRINTER_ENUM_CONNECTIONS | win32print.PRINTER_ENUM_LOCAL),  # Enumerate network printers
      None,  # No printer name filter
      2  # Return detailed printer information
    )
    printers = [x["pPrinterName"] for x in printer_objs]

  # for p in printers:
  ui.comboBox.addItems(printers)


  # pressing "Enter" key while in any of the fields can start the process
  ui.btn_selectPDF.clicked.connect(file_select)
  ui.btn_Go.clicked.connect(GoClicked)
  ui.txt_pdf_file.returnPressed.connect(GoClicked)
  ui.txt_num_pages.returnPressed.connect(GoClicked)

# Set up threadpool
  threadpool = QtCore.QThreadPool()
  threadpool.setMaxThreadCount(2)

  MainWindow.show()
  sys.exit(app.exec_())
