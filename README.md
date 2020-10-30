**CURRENTLY WIP**

# Ignite Installer

This is the installer script for setting up a local version of the ignite Jupyter-lab stack. Most people will just want to install via the [releases](https://github.com/Schulich-Ignite/installation-script/releases). 

## Steps

These are the current steps the installer takes, each step is labeled by a function called ```step_x``` where x is the step number.

1. install [python](https://www.python.org/)
2. install [nodejs](https://nodejs.org/en/) (windows ```msiexec.exe /i node-v12.19.0-x64.msi TARGETDIR="C:\Program Files\NodeJS" /quiet /promptrestart /l* install.log```)
3. Install [jupyterlab](https://jupyterlab.readthedocs.io/en/stable/): ```pip install jupyterlab```
4. Install [ipywidgets](https://ipywidgets.readthedocs.io/en/latest/): ```pip install ipywidgets && jupyter labextension install @jupyter-widgets/jupyterlab-manager@2.0```
5. Install [ipycanvas](https://ipycanvas.readthedocs.io/en/latest/index.html): ```pip install ipycanvas && jupyter labextension install @jupyter-widgets/jupyterlab-manager ipycanvas```
6. Install [ipyevents](https://github.com/mwcraig/ipyevents): ```pip install ipyevents && jupyter labextension install @jupyter-widgets/jupyterlab-manager ipyevents```
7. Install [git](https://git-scm.com) and [spark](https://github.com/Schulich-Ignite/spark): ```pip install git+https://github.com/Schulich-Ignite/spark```
8. Make an ```~/Documents/ignite_notebooks``` directory in the documents folder
9. Copy ```ignite.ipynb``` to ```~/Documents/ignite_notebooks```
10. Create shortcut on desktop ```cd ~/Documents/ignite_notebooks && jupyter-lab```

## Development guide

See CONTRIBUTING.md for development details

## Creating binary distributions

Below are the details for creating/updating binary distributions available in [releases](https://github.com/Schulich-Ignite/installation-script/releases)

### Install prerequisites

1. Install python3.6+
2. Run ```pip install -r requirements.txt```

### Windows

1. Run ```pyinstaller install_win.spec```
2. Binary is available at ```dist/install.exe```


## Known issues
- If you have jupyter or jupyter-lab installed on a python version other than 3.8.6 please uninstall them before running the script
- If you have anaconda installed, please uninstall it. It locks the installation of jupyter and doesn't allow other python versions to install it.

