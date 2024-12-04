import subprocess
import os
import tkinter
from tkinter import filedialog as fd
from tkinter import Tk
from spire.doc.common import *
from spire.doc import *
#----------------------------------------------------------------------------------------------------------------------------------------------------
def check():
    try:
        import spire.doc.common
        import spire.doc
        case()
    except ImportError as e:
            print('---------------------------------------------------------------------------------------------')
            print('Module named "spire.doc" not found. Please install spire.doc (pip install spire.doc)')
            print('---------------------------------------------------------------------------------------------')
            
            option = str(input('Do you want to install it? (y/n): '))
            if option == 'y':
                act = 'python.exe -m pip install --upgrade pip'
                print('Updating pip...')
                proc1 = subprocess.getoutput(["powershell", "-command", f"{act}"])
                print(proc1)
                command = 'pip install spire.doc'
                print ('-----------------------------')
                print('Installing the module spire.doc...')
                proc2 = subprocess.getoutput(["powershell", "-command", f"{command}"])
                print(proc2)
                print('Installation completed. Execute the script again with spire.doc installed.')
            else:
                print('Exiting...')
#----------------------------------------------------------------------------------------------------------------------------------------------------
def case():
    option = int(input('Do you want to convert one file or multiple files? (1 for one file / 2 for multiple files): '))
    if option == 1:
        convert_singfile()
    elif option == 2:
        convert_dir()
    else:
        print('Please enter a correct option: ')
        case()
#----------------------------------------------------------------------------------------------------------------------------------------------------
def select_folder():
    print('Please select the directory where the .docx files are located: ')
    root = tkinter.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    folder_selected = fd.askdirectory()
    root.destroy()
    return folder_selected
#----------------------------------------------------------------------------------------------------------------------------------------------------
def select_file():
    filetypes = [('Docx files', '*.docx'), ('All Files', '*.*')]
    print('Please select the .docx file do you want to convert: ')
    root = tkinter.Tk()
    root.withdraw()
    root.wm_attributes('-topmost', 1)
    file_path = fd.askopenfilename(title='Select a docx file', filetypes=filetypes)
    root.destroy()
    return file_path
#----------------------------------------------------------------------------------------------------------------------------------------------      
def convert_singfile():
    file_path = select_file()
    name_pdf = str(input('Enter the name for the pdf file (without the .pdf extension): '))
    dest_path = f'./pdf_converted/{name_pdf}.pdf'
    document = Document()
    document.LoadFromFile(f"{file_path}")
    document.SaveToFile(f"{dest_path}", FileFormat.PDF)
    print(f'PDF file has been saved in {dest_path}')
#---------------------------------------------------------------------------------------------------------------------------------------------------
def convert_dir():
    folder_path = select_folder()
    folder_filtered = [file for file in os.listdir(folder_path) if file.endswith(".docx")]
    for i in folder_filtered:
        document = Document()
        document.LoadFromFile(os.path.join(folder_path, i))
        base_name = os.path.splitext(i)[0]
        dest_path = f'./pdf_converted/{base_name}.pdf'
        document.SaveToFile(f"{dest_path}", FileFormat.PDF)
        print(f'PDF file {base_name}.pdf has been saved in the folder pdf_converted.')
        
check()

#made by @akhos09 github
