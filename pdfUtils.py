#-------------------------------------------------------------------------------
# Name:        module1
# Purpose:
#
# Author:      tejad
#
# Created:     12/09/2017
# Copyright:   (c) tejad 2017
# Licence:     <your licence>
#-------------------------------------------------------------------------------
__author__ = "Miguel Tejada"
__version__ = "0.1"
__email__ = "tejada.miguel@gmail.com"
__license__ = "tejada.miguel@gmail.com"
__versionHistory__ = [
    ["0.0", "170912", "MTEJADA", "START"],
    ["0.1", "231114", "MTEJADA", "Actualizacion python311"]]

import time
import subprocess
import winreg

import PyPDF2
import os,sys
import win32com.client as win32
import win32con
import win32gui
import win32ui
from os.path import join
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
elif __file__:
    application_path = os.path.dirname(__file__)
poppler_path = os.path.join(application_path,'poppler','Library','bin')
if not os.path.exists(poppler_path):
    win32ui.MessageBox("Missing poppler path!!! %s" % poppler_path, "Poppler Path not found")
import pdf2image # should install python-poppler

### ATENCION: WINDOWS POR DEFECTO DESACTIVA EN EL REGISTRO LA OPCION DEL MENU CONTEXTTUAL CON MAS DE 15 ITEMS. HAY QUE CAMBIARLO A MANO

def explorer_fileselection(ext=None):
    clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}'
    shellwindows = win32.Dispatch(clsid)
    files = []
    try:
        for window in range(shellwindows.Count):
            window_URL = shellwindows[window].LocationURL
            if not window_URL.startswith('file'):
                continue
            # window_dir = window_URL.split("///")[1].replace("/", "\\")
            if True: #window_dir == working_dir:
                selected_files = shellwindows[window].Document.SelectedItems()
                for ifile in range(selected_files.Count):
                    nfile = selected_files.Item(ifile).Path
                    if ext is None or nfile.endswith(ext):
                        files.append(nfile)
    except:
        win32ui.MessageBox("Close PDF Utils!", "Error")
    del shellwindows
    return files

def Merge(filelist):
    merger = PyPDF2.PdfMerger(strict=False)
    for fname in filelist:
        merger.append(fname)
    carpeta = join(os.path.dirname(filelist[0]),'merged.pdf')
    counter = 0
    while os.path.exists(carpeta):
        counter += 1
        carpeta = join(os.path.dirname(filelist[0]), 'merged_v{}'.format(counter) + ".pdf")
    merger.write(carpeta)

def Split(filelist):
    if isinstance(filelist,str):
        filelist = [filelist]
    for afile in filelist:
        infile = PyPDF2.PdfReader(afile, strict=False)
        for i in range(infile.getNumPages()):
            p = infile.getPage(i)
            outfile = PyPDF2.PdfFileWriter()
            outfile.addPage(p)
            name = os.path.splitext(afile)[0] + "_page" + str(i) + ".pdf"
            with open(name,'wb') as f:
                outfile.write(f)

def SplitToPNG(filelist):

    if isinstance(filelist,str):
        filelist = [filelist]
    for afile in filelist:
        pages = pdf2image.convert_from_path(afile, poppler_path=poppler_path)
        for i in range(len(pages)):
            name = str(os.path.splitext(afile)[0] + "_page" + str(i) + ".png")
            pages[i].save(name)

def main():
    if len(sys.argv) < 3:
        return
    op_type = sys.argv[1]
    files = [x for x in sys.argv[2:] if os.path.exists(x) and x.endswith('.pdf')]
    # files = explorer_fileselection(ext="pdf")
    if not files:
        return
    dirname = os.path.dirname(files[0])
    if not all([os.path.dirname(f)==dirname for f in files]):
        win32ui.MessageBox("Multiple windows opened with \nfiles selected", "PDFUTIL ERROR")
        return

    args = sys.argv
    if op_type == 'merge' in args:
        Merge(files)
    elif 'split' in args:
        Split(files)
    elif 'splitpng' in args:
        SplitToPNG(files)

def get_number_of_instances():
    _wmi = win32.GetObject('winmgmts:')
    processes = _wmi.ExecQuery('select * from win32_process')
    prog_ids = {}
    for x in processes:
        if x.Name != 'pdfUtils.exe':
            continue
        prog_ids[x.ProcessId] = x.ParentProcessId
    for k, v in list(prog_ids.items()):
        if v in prog_ids.keys():
            prog_ids.pop(k)
    del processes
    del _wmi
    return len(prog_ids)


def queryValue(key, name):
    value, type_id = winreg.QueryValueEx(key, name)
    return value


def show(key):
    for i in range(1024):
        try:
            n, v, t = winreg.EnumValue(key, i)
            print
            '%s=%s' % (n, v)
        except EnvironmentError:
            break


def set_environ_actif(actif):
    try:
        path = r'Environment'
        reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(reg, path, 0, winreg.KEY_ALL_ACCESS)
        name = "PDFUTIL"
        value = "1" if actif else "0"
        if name.upper() == 'PATH':
            value = queryValue(key, name) + ';' + value
        if value:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, value)
        else:
            winreg.DeleteValue(key, name)

    except Exception as e:
        print(e)
    finally:
        winreg.CloseKey(key)
        winreg.CloseKey(reg)


if __name__ == '__main__':
    main()