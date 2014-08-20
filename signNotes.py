#!/usr/bin/python 
# encoding: utf-8
'''
signNotes sign templates/databases
'''

# import sys
import os
from os import listdir
from notesapi import * 
from lotusscript import *
import pythoncom
from toolbox import myLog, SetVerbosity



from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

__all__ = []
__version__ = 0.1

DEBUG = 0
TESTRUN = 0
PROFILE = 0

notesapi = NotesApiWrapper()


def SignCollection(apiDB, comDB, theCollection, title):
    theCollection.BuildCollection()
    myLog("Signing %d Docs (%s)" % (theCollection.Count, title))
    cleanCount = 0

    noteID = theCollection.GetFirstNoteId()
    comNotesDocument = comDB.GetDocumentByID(noteID)
    while not (comNotesDocument is None):
        #** a little trick to avoid this message on recompiled forms:
        #** This document has been altered since the last time it was signed! Intentional tampering may have occurred.
        comNotesDocument.Sign()
        comNotesDocument.Save(True, False)
        
        signCount = cleanCount +1
        
        noteID = theCollection.GetNextNoteId(noteID)
        if (noteID <> ""):
            comNotesDocument = comDB.GetDocumentByID(noteID)
        else:
            comNotesDocument = None

    return True   


def SignDatabase(notesSession, sDatabaseName):
    global  notesapi
    result = False
    
    # Get Database as a COM object
    try:
        comDB = notesSession.GetDatabase("", sDatabaseName)
        myLog("Signing %s (%s)" % (sDatabaseName, comDB.Title.encode('ascii', 'ignore')))
        apiDB = notesapi.NSFDbOpen(sDatabaseName)
        nc = comDB.CreateNoteCollection(True)
        result = SignCollection(apiDB, comDB, nc, "Signing database")
        notesapi.NSFDbClose(apiDB)
        return result
        
    except Exception,err :
        print 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(err))
        
    except pythoncom.com_error as error:
        print (error)
        print (vars(error))
        print (error.args)

def SignDirectory(notesSession, sDirectory):
    myLog("Signing %s directory" % sDirectory)
    
    # Get NTF/NSF list from directory
    included_ext = ['ntf'] ;
    fileLst = [ fn for fn in listdir(sDirectory) if any([fn.endswith(ext) for ext in included_ext])];
    
    for notesFile in fileLst:
        result = SignDatabase(notesSession, sDirectory+"\\"+notesFile)
        if not result:
            print result
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
        parser.add_argument("-u", "--user",    dest="user",     action="store", help="user id short filename (relative to Notes/Data)")
        parser.add_argument("-v", "--verbose", dest="verbose", action="count", help="set verbosity level [default: %(default)s]")

        # Process arguments
        args = parser.parse_args()
        SetVerbosity(args.verbose)

        if (args.user is not None):
            SwitchUser(args.user)

        notesSession = SetupNotesClient(args.verbose)
        notesapi.NotesInitExtended(sys.argv)

        if not (args.name is None):
            result = SignDatabase(notesSession, args.name)
        elif (not args.directory is None):
            result = SignDirectory(notesSession, args.directory.strip())
        else:
            print "Database not specified"

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