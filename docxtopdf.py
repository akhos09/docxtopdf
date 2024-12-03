import subprocess
import tkinter
from tkinter import filedialog as fd
from tkinter import Tk
from spire.doc import *
from spire.doc.common import *
def check():
    try:
        import spire.doc
        convert()
    except ImportError as e:
        print('---------------------------------------------------------------------')
        print('Module named "Spire.Doc" not found. Please install Spire.Doc (pip install Spire.Doc)')
        print('---------------------------------------------------------------------')
        option = str(input('Do you want to install it? (y/n): '))
        if option == 'y':
            act = 'python.exe -m pip install --upgrade pip'
            print('Updating pip...')
            proc1 = subprocess.getoutput(["powershell", "-command", f"{act}"])
            print(proc1)
            command = 'pip install Spire.Doc'
            print ('-----------------------------')
            print('Installing the module pick...')
            proc2 = subprocess.getoutput(["powershell", "-command", f"{command}"])
            print(proc2)
            print('Installation completed. Execute the script again with Spire.Doc installed.')
        else:
            print('Exiting...')
#----------------------------------------------------------------------------------------------------------------------------------------------------
def select_folder():
    print('Please select the directory where the .docx files are located: ')
    root = tkinter.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_selected = fd.askdirectory()
    root.destroy()
    return folder_selected

def select_file():
    filetypes = [('Docx files', '*.docx'), ('All Files', '*.*')]
    print('Please select the .docx file do you want to convert: ')
    root = tkinter.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    file_path = fd.askopenfilename(title='Select a docx file', filetypes=filetypes)
    root.destroy()
    return file_path
    
def convert_dir():
    folder_selected = select_folder()
    for i in select_folder(folder_selected):    
        print(f'i')
check()
