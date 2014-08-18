#!/usr/bin/python 
# encoding: utf-8
'''
Short : compileNotes compiles a notes database without having to run designer

Original stuff from Julian Robichaux http://www.nsftools.com

'''

import sys
import os
from os import listdir
import pythoncom
from notesapi import * 
from lotusscript import *
from toolbox import myLog, SetVerbosity

from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

__all__ = []
__version__ = 0.1

DEBUG = 0
TESTRUN = 0
PROFILE = 0

notesapi = NotesApiWrapper()

def RecompileLS(apiDb, comDb, noteID):
    global lastError

    try:
        if (apiDb is None):
            lastError = "Database is not open"
            return False
    
        comNotesDocument = comDb.GetDocumentByID(noteID)
    
        hNote = notesapi.NSFNoteOpen(apiDb,noteID,0)
        
    #     notesClass = notesapi.NSFNoteGetClass(hNote)
    #     print notesClass
        #** first, we compile the note
        result = notesapi.NSFNoteLSCompile(apiDb, hNote, 0)
        if (result <> 0):
            if (result <> 0x0222):
                lastError = "Cannot compile LotusScript"
                return False
            else:
                myLog("Compilation error  %04X" % result)
         
        #** then we sign it
        result = notesapi.NSFNoteSign(hNote)
        if (result <> 0):
            lastError="Cannot sign"
            return False
         
        #** then we save it
        result = notesapi.NSFNoteUpdate(hNote, 0)
        if (result <> 0):
            lastError="Cannot save"
            return False
        notesapi.NSFNoteClose(hNote)
        
        lastError = ""
        return True

    except Exception,err :
        print 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(err))
        return False


def CompileCollection(apiDB,comDB, theCollection,title):
    
    theCollection.BuildCollection()
    myLog("Compiling %d items (%s)" % (theCollection.Count, title))
    compileCount = 0

    noteID = theCollection.GetFirstNoteId()
    while (noteID <> ""):
        if RecompileLS(apiDB,comDB,noteID):
            myLog("  - " + GetTitle(comDB.GetDocumentByID(noteID)))
            compileCount = compileCount +1
        else:
            print "Error (%s) compiling %s " % (lastError,GetTitle(comDB.GetDocumentByID(noteID)))
            break
         
        noteID = theCollection.GetNextNoteId(noteID)

    myLog("Compiled %d items\n" % compileCount)
    return (theCollection.Count == compileCount)
    
def CompileDatabase(notesSession, sDatabaseName):
    global  notesapi
    result = False
    try:    
        
        # Get Database as a COM object
        comDB = notesSession.GetDatabase("", sDatabaseName)
        myLog("Compiling %s (%s)" % (sDatabaseName, comDB.Title.encode('ascii', 'ignore')))
        apiDB = notesapi.NSFDbOpen(sDatabaseName)
        #** compile the script libraries first (note that this will NOT build a
        #** dependency tree -- rather, we'll try to brute-force around the 
        #** dependencies by recompiling until either (A) there are no errors,
        #** or (B) the number of errors we get is the same as we got last time)

        nc = comDB.CreateNoteCollection(False)
        nc.SelectScriptLibraries = True
        if (CompileCollection(apiDB,comDB, nc,"Script libraries")):
            #** then compile everything else
            nc = comDB.CreateNoteCollection(True)
            nc.SelectScriptLibraries = False
            nc.SelectACL = False
            nc.SelectDataConnections= False
            nc.SelectHelpAbout= False
            nc.SelectHelpIndex= False
            nc.SelectHelpUsing= False
            nc.SelectIcon= False
            nc.SelectImageResources= False
            nc.SelectJavaResources= False
            nc.SelectProfiles= False
            nc.SelectScriptLibraries= False
            nc.SelectStyleSheetResources= False
            if CompileCollection(apiDB, comDB,nc,"The rest"):
                result = True
        
        notesapi.NSFDbClose(apiDB)
        return result
    except Exception,err :
        print 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(err))
        return False
        
    except pythoncom.com_error as error:
        print (error)
        print (vars(error))
        print (error.args)

def CompileDirectory(notesSession, sDirectory):
    myLog("Compiling %s directory" % sDirectory)
    
    # Get NTF/NSF list from directory
    included_ext = ['ntf'] ;
    fileLst = [ fn for fn in listdir(sDirectory) if any([fn.endswith(ext) for ext in included_ext])];
    
    for notesFile in fileLst:
        result = CompileDatabase(notesSession, sDirectory+"\\"+notesFile)
        if not result: 
			break;
			
    return result
        
def main(argv=None): # IGNORE:C0111
    global  notesapi

    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_name = os.path.basename(sys.argv[0])

    try:
        # Setup argument parser
        parser = ArgumentParser(description="", formatter_class=RawDescriptionHelpFormatter)
        parser.add_argument("-n", "--name",    dest="name", action="store", help="NSF/NTF to compile]")
        parser.add_argument("-f", "--fulldir",    dest="directory", action="store", help="path to compile]")
        parser.add_argument("-v", "--verbose", dest="verbose", action="count", help="set verbosity level [default: %(default)s]")

        # Process arguments
        args = parser.parse_args()

        SetVerbosity(args.verbose)

        notesSession = SetupNotesClient(args.verbose)
        notesapi.NotesInitExtended(sys.argv)

        if not (args.name is None):
            result = CompileDatabase(notesSession, args.name)
        elif (not args.directory is None):
            result = CompileDirectory(notesSession, args.directory.strip())
        else:
            print "No input specified"

        notesapi.NotesTerm
        return 0 if (result == True) else 1

    except KeyboardInterrupt:
        ### handle keyboard interrupt ###
        return 0
    except Exception, e:
		print 'Error on line {}\n'.format(sys.exc_info()[-1].tb_lineno)
		print ("Error: %s.\n" % str(e))
		if DEBUG or TESTRUN:
			raise(e)
		indent = len(program_name) * " "
		sys.stderr.write(program_name + ": " + repr(e) + "\n")
		sys.stderr.write(indent + "  for help use --help")
		notesapi.NotesTerm
		return 2

    
if __name__ == "__main__":
    if DEBUG:
        #sys.argv.append("-h")
        sys.argv.append("-v")
        #sys.argv.append("-r")
    if TESTRUN:
        import doctest
        doctest.testmod()
    if PROFILE:
        import cProfile
        import pstats
        profile_filename = 'upgradeCode_profile.txt'
        cProfile.run('main()', profile_filename)
        statsfile = open("profile_stats.txt", "wb")
        p = pstats.Stats(profile_filename, stream=statsfile)
        stats = p.strip_dirs().sort_stats('cumulative')
        stats.print_stats()
        statsfile.close()
        sys.exit(0)
    sys.exit(main())