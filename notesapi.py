#!/usr/bin/python 
# encoding: utf-8
'''
NotesApi (nnotes.dll) incomplete wrapper 
'''

import sys
from ctypes import WinDLL,c_short,c_int,c_char_p,create_string_buffer,POINTER,byref

NOTES_NOERROR   = 0

class NotesApiWrapper:
    '''
    classdocs
    '''
    bInitialized = False
 
    def __init__(self):
        self.notesDll = WinDLL("nnotes.dll")
        self.bInitialized = True

            
    def NotesInitExtended(self,argv):
        self.notesDll.NotesInitExtended.argtypes =  [c_int,POINTER(c_char_p)]
        self.notesDll.NotesInitExtended.restype =  c_short

        argv_type = c_char_p * len(argv)
        myArgv = argv_type(*argv)
        myArgc = c_int(len(argv))
        
        result = self.notesDll.NotesInitExtended(myArgc,myArgv)
        if (result <> NOTES_NOERROR):
            print "NotesInitExtended Result " + hex(result)
        return result
    
    def NotesTerm(self):
        self.notesDll.NotesTerm
 
    def NotesGetErrorString (self,error):
        self.notesDll.OSLoadString.argtypes = [c_int, c_short,c_char_p,c_short]
        self.notesDll.OSLoadString.restype =  c_short
        readBuffer = create_string_buffer(256)
        self.notesDll.OSLoadString(0,error,readBuffer,256)
        return readBuffer.value

    def PrintResult(self,result):
        if (result <>0 ):
            errString = self.NotesGetErrorString(result)
            print "[%s - %d] %s - Error %04X : %s" % \
                (sys._getframe(2).f_code.co_name,sys._getframe(2).f_lineno, sys._getframe(1).f_code.co_name,result, errString)
             
        
    def NSFDbOpen(self, dbFilename):
        self.notesDll.NSFDbOpen.argtypes = [c_char_p, POINTER(c_int)]
        self.notesDll.NSFDbOpen.restype =  c_short
        hDb = c_int()
        sDbName = c_char_p(dbFilename)
        
        result = self.notesDll.NSFDbOpen(sDbName, byref(hDb))
        self.PrintResult(result)
        return hDb
        
    def NSFDbClose(self, hDb):
        self.notesDll.NSFDbClose.argtypes = [c_int]
        self.notesDll.NSFDbClose.restype =  c_short
        result = self.notesDll.NSFDbClose(hDb)
        self.PrintResult(result)

    def NSFNoteOpen(self, hDb, notesId, open_flags):
        self.notesDll.NSFNoteOpen.argtypes = [c_int,c_int,c_int, POINTER(c_int)]
        self.notesDll.NSFNoteOpen.restype =  c_short
        hNote = c_int()

        result = self.notesDll.NSFNoteOpen(hDb,int(notesId,16), open_flags,byref(hNote))
        self.PrintResult(result)
        return hNote.value

    def NSFNoteClose(self, hNote):
        self.notesDll.NSFNoteClose.argtypes = [c_int]
        self.notesDll.NSFNoteClose.restype =  c_short
        result = self.notesDll.NSFNoteClose(hNote)
        self.PrintResult(result)
        return result
 
    def NSFNoteLSCompile(self, hDb, hNote,dwFlags):
        self.notesDll.NSFNoteLSCompile.argtypes = [c_int,c_int,c_int]
        self.notesDll.NSFNoteLSCompile.restype =  c_short
        result = self.notesDll.NSFNoteLSCompile(hDb, c_int(hNote),dwFlags)
        if (result <> 0x0222 ):
            self.PrintResult(result)
        return result

    def NSFNoteLSCompileExt(self, hDb, hNote,dwFlags,pfnErrProc,pCtr):
        self.notesDll.NSFNoteLSCompileExt.argtypes = [c_int,c_int,c_int]
        self.notesDll.NSFNoteLSCompileExt.restype =  c_short
        result = self.notesDll.NSFNoteLSCompileExt(hDb, hNote,dwFlags,pfnErrProc,pCtr)
        self.PrintResult(result)
        return result
    
    def NSFNoteSign(self, hNote):
        self.notesDll.NSFNoteSign.argtypes = [c_int]
        self.notesDll.NSFNoteSign.restype =  c_short

        result = self.notesDll.NSFNoteSign(hNote)
        self.PrintResult(result)
        return result

    def NSFNoteUpdate(self, hNote,flags):
        self.notesDll.NSFNoteUpdate.argtypes = [c_int,c_short]
        self.notesDll.NSFNoteUpdate.restype =  c_short

        result = self.notesDll.NSFNoteUpdate(hNote,flags)
        self.PrintResult(result)
        return result

    def NSFNoteGetClass(self, hNote):
        self.notesDll.NSFNoteGetInfo.argtypes = [c_int,c_short, POINTER(c_short)]
        noteclass = c_short()
        self.notesDll.NSFNoteGetInfo(hNote,3,byref(noteclass))
        return noteclass
