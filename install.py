
# Standard Library Dependencies
import os                             # Used for path validation
import sys                            # Used to access info about system
import logging                        # Used for optional logging details               
import traceback                      # Used to log error details on raise
import subprocess                     # Used to invoke any necessary binaries
from shutil import copyfile           # Used to copy files between directories

# Third-Party Dependencies
import winshell                       # Allows execution of winshell functions
import requests                       # Used to download any necessary binary files
from tqdm import tqdm                 # Used to generate a progress bar on donwloads
from elevate import elevate           # Forces sudo through a gui on linux & MacOS systems
from win32com.client import Dispatch  # Instantiate COM objects to dispatch any tasks through

# Setting up Constants

#  OS name booleans
mac = True if sys.platform == "darwin" else False
windows = True if os.name == "nt" else False
linux = True if sys.platform == "linux" or sys.platform == "linux2" else False

# Instalation and download folders
DOWNLOAD_FOLDER = f"{os.getenv('USERPROFILE')}\\Downloads" if windows else f"{os.getenv('HOME')}/Downloads"
DESKTOP_FOLDER = f"{os.getenv('USERPROFILE')}\\Desktop" if windows else f"{os.getenv('HOME')}/Desktop" 
DOCUMENTS_FOLDER = f"{os.getenv('USERPROFILE')}\\Documents" if windows else f"{os.getenv('HOME')}/Documents" 

# Executable paths
if windows:
    PIP_EXECUTABLE = os.path.realpath(f"{os.environ['ProgramFiles']}\\Python38\\Scripts\\pip.exe")
    JUPYTER_EXECUTABLE = os.path.realpath(f"{os.environ['ProgramFiles']}\\Python38\\Scripts\\jupyter.exe")
    JUPYTER_LAB_EXECUTABLE = os.path.realpath(f"{os.environ['ProgramFiles']}\\Python38\\Scripts\\jupyter-lab.exe")
else:
    elevate(show_console=False) # Displays a popup window to give script sudo access
    PIP_EXECUTABLE = "python3.8"
    JUPYTER_EXECUTABLE = "jupyter"       # Don't know enough about macos to make version specific
    JUPYTER_LAB_EXECUTABLE = "jupyter-lab" # Don't know enough about macos to make version specific


def _download(name, url, extension) -> str:
    """Downloads binaries from remote sources"""
    file_path = f"{DOWNLOAD_FOLDER}{os.sep}{name}{extension}"

    if os.path.exists(file_path): # If file already exists
        logging.info(f"File {file_path} already downloaded")
        return file_path

    logging.info(f"Downloading {name}")
    logging.info("Starting binary download")

    # Setting up necessary download variables
    file_stream = requests.get(url, stream=True) # The open http request for the file
    chunk_size = 1024 # Setting the progress bar chunk size to measure in kb
    total_length = int(file_stream.headers.get('content-length')) # Getting file size

    # Setting up the download progress bar
    progress_bar = tqdm(total=total_length, unit='iB', unit_scale=True)
    progress_bar.set_description(f"Download progress for {name}:")

    # Write the incoming data stream to a file and update progress bar as it downloads
    with open(file_path, 'wb') as download_file: 
        for chunk in file_stream.iter_content(chunk_size): 
            if chunk:
                progress_bar.update(len(chunk))
                download_file.write(chunk)
    progress_bar.close()

    return file_path

def _install(path, args):
    """Install executable files with provided args"""
    print(f"Installing {path}")
    logging.debug(f"Installing {path}")
    logging.debug("Installing: " + str([path, *args]))
    subprocess.call([path, *args], shell=True)


def step_1():
    """install python 3.8.6"""
    print("Entering Step 1; Install Python 3.8.6")
    logging.debug("Entering Step 1; Install Python 3.8.6")
    if windows:
        if os.path.exists(PIP_EXECUTABLE):
            logging.debug("Python and pip already isntalled, skipping python installation")
        else:
            exc_path = _download("python-installer", "https://www.python.org/ftp/python/3.8.6/python-3.8.6-amd64.exe", ".exe") 
            _install(exc_path, ["/quiet", "PrependPath=1", "InstallAllUsers=1"])
            os.remove(exc_path)
    elif linux:
        ...
    elif mac:
        # TODO: add check for if python and pip are already installed
        exc_path = _download("python-installer", "https://www.python.org/ftp/python/3.8.6/python-3.8.6-macosx10.9.pkg", ".pkg")
        subprocess.call(["installer", "-pkg", exc_path, "-target" "~"]) # Install python 3.8.6

def step_2():
    """Install NodeJS"""
    print("Entering Step 2; Install NodeJS")
    logging.debug("Entering Step 2; Install NodeJS")
    if windows:
        if os.path.exists(os.path.realpath(f"{os.environ['ProgramFiles']}\\nodejs")):
            logging.debug("NPM and nodeJS already installed, skipping installation")
        else:
            exc_path = _download("node", "https://nodejs.org/dist/v12.19.0/node-v12.19.0-x64.msi", ".msi")
            # Can't use _install() because need to use msi tools
            subprocess.Popen(f'msiexec.exe /i {exc_path} ADDLOCAL=\"DocumentationShortcuts,EnvironmentPathNode,EnvironmentPathNpmModules,npm,NodeRuntime,EnvironmentPath\" /passive',
            universal_newlines=True, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()

            from tkinter import Tk, Label
            from tkinter import messagebox 
            
            root = Tk() 
            root.geometry("500x500") 


            details = Label(root, text ='Please close this window and re-run install.exe to finalize installation', font = "50")  
            details.pack()
            root.mainloop()

            os.remove(exc_path)
            sys.exit()
    elif linux:
        ...
    elif mac:
        try:# TODO: Validate this works
            subprocess.call(["npm"]) 
            logging.debug("NPM and nodeJS already installed, skipping installation")
        except FileNotFoundError:
            exc_path = _download("node", "https://nodejs.org/dist/v12.19.0/node-v12.19.0.pkg", ".pkg")
            subprocess.call(["installer", "-pkg", exc_path, "-target" "~"]) # Install nodejs

def step_3_to_6():
    """Install pip packages; jupyterlab, ipywidgets, ipycanvas, ipyevents, spark"""
    print("Entering Steps 3-6; Install Python and Jupyterlab Packages")
    logging.debug("Entering Steps 3-6; Install Python and Jupyterlab Packages")


    logging.info("Installing pip packages")
    for package in ["jupyterlab", "ipywidgets", "ipycanvas", "ipyevents"]:
        logging.debug(f"Installing pip package {package} with pip executable {PIP_EXECUTABLE}")
        subprocess.call([PIP_EXECUTABLE, "install", package], shell=True)

    logging.debug("Installing ipywidgets")
    subprocess.call([JUPYTER_EXECUTABLE, "labextension", "install", f"@jupyter-widgets/jupyterlab-manager"], shell=True)

    logging.info("Installing jupyter packages")
    for jupyter_package in ["ipycanvas", "ipyevents"]:
        logging.debug(f"Installing JupyterLab package {jupyter_package} with jupyter executable {JUPYTER_EXECUTABLE}")
        subprocess.call([JUPYTER_EXECUTABLE, "labextension", "install", "@jupyter-widgets/jupyterlab-manager", jupyter_package], shell=True)

    subprocess.call([JUPYTER_EXECUTABLE, "lab", "build"])
    logging.debug("Finished installing all pip and jupyter packages")

def step_7():
    """Install the spark package for use in jupyterlab"""
    print("Entering Step 7; Install Spark")
    logging.debug("Entering Step 7; Install Spark")
    subprocess.call([PIP_EXECUTABLE, "install", "schulich-ignite"], shell=True)


def step_8_to_9():
    """Create a folder in the documents folder called ignite_notebooks with a default notebook called ignite.ipynb"""
    print("Entering Step 8-9; Setup Default ignite folder and notebook")
    logging.debug("Entering Step 8-9; Setup Default ignite folder and notebook")

    # Create default notebook folder
    if not os.path.exists(f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks"):
        logging.debug("No default notebook folder found, initializing")
        os.mkdir(f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks")
    else:
        logging.debug("Default notebook folder found, skipping creation")
    
    # Create default notebook file
    if not os.path.exists(f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks{os.sep}ignite.ipynb"):
        logging.debug("No default notebook found, initializing")
        template_file = _download("ignite", "https://raw.githubusercontent.com/Descent098/installation-script/master/ignite.ipynb", ".ipynb")
        copyfile(template_file,f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks{os.sep}ignite.ipynb")
    else:
        logging.debug("Default notebook found, skipping creation")

    subprocess.call([JUPYTER_EXECUTABLE, "trust", f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks{os.sep}ignite.ipynb"])

def step_10():
    """Adds an ignite icon to the desktop for easy launching"""
    print("Entering Step 10; Setup Desktop Shortcut")
    logging.debug("Entering Step 10; Setup Desktop Shortcut")
    # TODO: Change path to Schulich Ingnite repo
    icon_file = _download("ignite", "https://raw.githubusercontent.com/Descent098/installation-script/master/ignite.ico", "ico")
    copyfile(icon_file,f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks{os.sep}ignite.ico")
    if windows:

        logging.debug("Setting up shortcut attributes")
        path = os.path.join(DESKTOP_FOLDER, "Ignite.lnk")  # Setup path for shortcut
        target = JUPYTER_LAB_EXECUTABLE  # Setup path for target executable
        wDir = f"{DOCUMENTS_FOLDER}\\ignite_notebooks"  # Set directory to run executeable from
        icon = f"{DOCUMENTS_FOLDER}{os.sep}ignite_notebooks{os.sep}ignite.ico"  # Set icon path

        logging.debug("Grabbing windows shell scripts")
        shell = Dispatch('WScript.Shell')  # Grab the WScript shell function to build shortcuts

        logging.debug("Creating shortcut template")
        shortcut = shell.CreateShortCut(path)  # Begin creating shortcut objects
        # Add previously built variables to shortcut object

        logging.debug("Writing shortcut attributes to object")
        shortcut.Targetpath = target
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon

        logging.debug("Flushing shortcut")
        shortcut.save() # Flush shortcut to the desktop


class SysLogger:
    # Adapted from https://stackoverflow.com/a/31688396/11602400
    def __init__(self, level):
        self.level = level # Define the log level for logs from stream

    def write(self, message):
        """Writes the message to the instance

        Parameters
        ----------
        message : str
            The incomming message from the stream
        """
        if message != '\n':
            self.level(message)

    def flush(self):
        """Flush all messages to the logs"""
        self.level(sys.stderr)

def main():
    logging.basicConfig(
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)-8s %(message)s",
        datefmt="%y-%m-%d %H:%M:%S",
        handlers=[logging.FileHandler("ignite_install.log"),
                        logging.StreamHandler()] # TODO: Remove on launch
        )
        
    logger = logging.getLogger(__name__)
    sys.stdout = SysLogger(logger.debug)
    sys.stderr = SysLogger(logger.warning)

    
    try:

        step_1()
        step_2()
        step_3_to_6()
        step_7()
        step_8_to_9()
        step_10()

    except Exception as identifier:
        logging.error(f"{identifier}: {traceback.format_exc()}")
        sys.exit()

if __name__ == "__main__":
    main()
