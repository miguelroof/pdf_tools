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

import PyPDF2
import os,sys
from os.path import join
import win32com.client as win32
import win32ui, win32gui
# from PythonMagick import *
# import wand

### ATENCION: WINDOWS POR DEFECTO DESACTIVA EN EL REGISTRO LA OPCION DEL MENU CONTEXTTUAL CON MAS DE 15 ITEMS. HAY QUE CAMBIARLO A MANO

os.environ['MAGICK_HOME'] = os.path.abspath('.')
def explorer_fileselection(ext=None):
    working_dir = os.getcwd()
    clsid = '{9BA05972-F6A8-11CF-A442-00A0C90A8F39}' #Valid for IE as well!
    shellwindows = win32.Dispatch(clsid)
    files = []
    try:
        for window in range(shellwindows.Count):
            window_URL = shellwindows[window].LocationURL
            if not window_URL.startswith('file'): continue
            window_dir = window_URL.split("///")[1].replace("/", "\\")
            if True: #window_dir == working_dir:
                selected_files = shellwindows[window].Document.SelectedItems()
                for ifile in range(selected_files.Count):
                    nfile = selected_files.Item(ifile).Path
                    if ext is None or nfile.endswith(ext):
                        files.append(nfile)
    except:   #Ugh, need a better way to handle this one
        win32ui.MessageBox("Close IE!", "Error")
    del shellwindows
    return files

def Merge(filelist):
    merger = PyPDF2.PdfFileMerger()
    for fname in filelist:
        merger.append(PyPDF2.PdfFileReader(open(fname,'rb')))
    carpeta = join(os.path.dirname(filelist[0]),'merged.pdf')
    merger.write(carpeta)

def Split(filelist):
    if isinstance(filelist,str):
        filelist = [filelist]
    for afile in filelist:
        infile = PyPDF2.PdfFileReader(open(afile,'rb'))
        for i in xrange(infile.getNumPages()):
            p = infile.getPage(i)
            outfile = PyPDF2.PdfFileWriter()
            outfile.addPage(p)
            name = os.path.splitext(afile)[0] + "_page" + str(i) + ".pdf"
            with open(name,'wb') as f:
                outfile.write(f)

def SplitToPNG(filelist):
    oldenv = os.environ['MAGICK_HOME']
    os.environ['MAGICK_HOME'] = os.path.abspath('.')
    if isinstance(filelist,str):
        filelist = [filelist]
    for afile in filelist:
        infile = PyPDF2.PdfFileReader(open(afile,'rb'))
        for i in xrange(infile.getNumPages()):
            img = Image()
            img.density("300")
            img.read(str(afile)+'[' + str(i) + ']')
            name = str(os.path.splitext(afile)[0] + "_page" + str(i) + ".png")
            img.write(name)
            del img
        del infile
    os.environ['MAGICK_HOME'] = oldenv

def main():
    args = sys.argv
    if 'merge' in args:
        files = explorer_fileselection(ext="pdf")
        if not files:
            win32ui.MessageBox("Choose Files to merge!", "MERGE ERROR"); return
        dirname = os.path.dirname(files[0])
        if not all([os.path.dirname(f)==dirname for f in files]):
            win32ui.MessageBox("Multiple windows opened with \nfiles selected", "MERGE ERROR")
            return
        Merge(files)
    elif 'split' in args:
        files = explorer_fileselection(ext="pdf")
        if not files:
            win32ui.MessageBox("Choose Files to split!", "SPLIT ERROR"); return
        dirname = os.path.dirname(files[0])
        if not all([os.path.dirname(f)==dirname for f in files]):
            win32ui.MessageBox("Multiple windows opened with \nfiles selected", "SPLIT ERROR")
            return
        Split(files)
    elif 'splitpng' in args:
        files = explorer_fileselection(ext="pdf")
        if not files:
            win32ui.MessageBox("Choose Files to split!", "SPLIT ERROR"); return
        dirname = os.path.dirname(files[0])
        if not all([os.path.dirname(f)==dirname for f in files]):
            win32ui.MessageBox("Multiple windows opened with \nfiles selected", "SPLIT ERROR")
            return
        SplitToPNG(files)

if __name__ == '__main__':
    # main()
    tlist = explorer_fileselection('pdf')
    print(len(tlist))
##    SplitToPNG(tlist)

