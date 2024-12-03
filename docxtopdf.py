import subprocess
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
def convert():
    print('hola')
check()
