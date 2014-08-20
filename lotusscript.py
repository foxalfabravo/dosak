#!/usr/bin/python 
# encoding: utf-8
'''
Domino & Lotusscript toolbox 
'''

import win32com.client
import os

from win32com.client import makepy
from toolbox import which
from toolbox import setIniKey
from toolbox import myLog
from toolbox import checkProcessRunning

def SetupNotesClient(verbose):
    makepy.GenerateFromTypeLibSpec('Lotus Domino Objects',bForDemand = 0, verboseLevel=verbose)
    makepy.GenerateFromTypeLibSpec('IBM Notes Automation Classes',bForDemand = 0, verboseLevel=verbose)
    notesSession = win32com.client.Dispatch('Lotus.NotesSession')
    notesSession.Initialize("")
    return notesSession
   
def FindDesigner():
    path2Notes = which("Designer.exe")
    if path2Notes is None:
        print "Designer not found - Please check your PATH and/or your installation"
    return path2Notes 

def FindNotes():
    path2Notes = which("Notes.exe")
    if path2Notes is None:
        print "Notes not found - Please check your PATH and/or your installation"
    return path2Notes 
        
def EnableHeadlessDesigner(path2notes):
    notesIniFile= os.path.join(path2notes, "notes.ini")
    setIniKey(notesIniFile, "Notes","DESIGNER_AUTO_ENABLED","true") 
    setIniKey(notesIniFile, "Notes","CREATE_R85_DATABASES","1")

def SwitchUser(sIdFile):
    path2notes, desName = os.path.split(FindNotes())
    notesIniFile= os.path.join(path2notes, "notes.ini")
    myLog("Switching to ID file %s" % sIdFile)
    setIniKey(notesIniFile, "Notes","KeyfileName",sIdFile)
     
       
def iterateDocuments(docs):
    doc = docs.getFirstDocument()
    while doc:
        yield doc
        doc = docs.getNextDocument(doc)
        
def isNotesRunning():
    return checkProcessRunning("notes2")

def isDesignerRunning():
    return checkProcessRunning("notes2")

def GetTitle(comNotesDocument):
    if (comNotesDocument is None):
        return ""
    else:
        item = comNotesDocument.GetFirstItem("$TITLE")
        if not(item is None):
            title = item.Text
        else:
            title = "no title"
        return title.encode('ascii', 'ignore')