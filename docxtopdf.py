import subprocess
import tkinter
from tkinter import filedialog as fd
from tkinter import Tk
import aspose.words as aw
#----------------------------------------------------------------------------------------------------------------------------------------------------
def check():
    try:
        import docx2pdf
        case()
    except ImportError as e:
            print('---------------------------------------------------------------------')
            print('Module named "docx2pdf" not found. Please install docx2pdf (pip install docx2pdf)')
            print('---------------------------------------------------------------------')
            option = str(input('Do you want to install it? (y/n): '))
            if option == 'y':
                act = 'python.exe -m pip install --upgrade pip'
                print('Updating pip...')
                proc1 = subprocess.getoutput(["powershell", "-command", f"{act}"])
                print(proc1)
                command = 'pip install docx2pdf'
                print ('-----------------------------')
                print('Installing the module docx2pdf...')
                proc2 = subprocess.getoutput(["powershell", "-command", f"{command}"])
                print(proc2)
                print('Installation completed. Execute the script again with docx2pdf installed.')
            else:
                print('Exiting...')
#----------------------------------------------------------------------------------------------------------------------------------------------------
def case():
    option = int(input('Do you want to convert one file or multiple files? (1 for one file/2 for multiple files): '))
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
#---------------------------------------------------------------------------------------------------------------------------------------------------
def convert_dir():
    folder_selected = select_folder()

#----------------------------------------------------------------------------------------------------------------------------------------------      
def convert_singfile():
    file_path = select_file()
    name_pdf = str(input('Enter the name for the pdf file: '))
    dest_path = f'./pdf_converted/{name_pdf}.pdf'
    doc = aw.Document(f"{file_path}")
    doc.save(f"{dest_path}")
    print(f'PDF file has been saved in {dest_path}')
check()
