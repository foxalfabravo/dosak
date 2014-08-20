#!/usr/bin/python 
# encoding: utf-8
'''
Short : cleanNotes removes noise from databases

'''

import sys
import os
from os import listdir
import pythoncom

from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

from notesapi import NotesApiWrapper
from toolbox import SetVerbosity,myLog
from lotusscript import SetupNotesClient,SwitchUser


__all__ = []
__version__ = 0.1

DEBUG = 0
TESTRUN = 0
PROFILE = 0

notesapi = NotesApiWrapper()

def CleanCollection(apiDB, comDB, theCollection, title):
    theCollection.BuildCollection()
    myLog("Cleaning %d Docs (%s)" % (theCollection.Count, title))
    cleanCount = 0

    noteID = theCollection.GetFirstNoteId()
    comNotesDocument = comDB.GetDocumentByID(noteID)
    while not (comNotesDocument is None):
        #get it    
        comNotesDocument.RemoveItem("$Updatedby")
        
        comNotesDocument.Save(True,False,False)
        cleanCount = cleanCount +1
        
        noteID = theCollection.GetNextNoteId(noteID)
        if (noteID <> ""):
            comNotesDocument = comDB.GetDocumentByID(noteID)
        else:
            comNotesDocument = None
    
    return True   

def CleanDatabase(notesSession, sDatabaseName):
    global  notesapi
    result = False
    
    # Get Database as a COM object
    try:
        comDB = notesSession.GetDatabase("", sDatabaseName)
        myLog("Cleaning %s (%s)" % (sDatabaseName, comDB.Title.encode('ascii', 'ignore')))
        apiDB = notesapi.NSFDbOpen(sDatabaseName)
        nc = comDB.CreateNoteCollection(True)
        result = CleanCollection(apiDB, comDB, nc, "Cleaning design elements")
        notesapi.NSFDbClose(apiDB)
        return result
        
    except Exception,err :
        print 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(err))

    except pythoncom.com_error as error:
        print (error)
        print (vars(error))
        print (error.args)


def CleanDirectory(notesSession, sDirectory):
    myLog( "Cleaning %s directory" % sDirectory)
    
    # Get NTF/NSF list from directory
    included_ext = ['ntf'] ;
    fileLst = [ fn for fn in listdir(sDirectory) if any([fn.endswith(ext) for ext in included_ext])];
    
    for notesFile in fileLst:
        result = CleanDatabase(notesSession, sDirectory+"\\"+notesFile)
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
        parser.add_argument("-n", "--name",    dest="name", action="store", help="NSF/NTF to clean up]")
        parser.add_argument("-f", "--fulldir",    dest="directory", action="store", help="path to compile]")
        parser.add_argument("-u", "--user",    dest="user",     action="store", help="user id short filename (relative to Notes/Data)")
        parser.add_argument("-v", "--verbose", dest="verbose", action="count", help="set verbosity level [default: %(default)s]")

        # Process arguments
        args = parser.parse_args()

        SetVerbosity(args.verbose)
        myLog("Args are %s " % args)

        if (args.user is not None):
            SwitchUser(args.user)

        notesSession = SetupNotesClient(args.verbose)
        notesapi.NotesInitExtended(sys.argv)      

        if not (args.name is None):
            result = CleanDatabase(notesSession, args.name)
        elif (not args.directory is None):
            result = CleanDirectory(notesSession, args.directory.strip())
        else:
            print "No input specified"

        notesapi.NotesTerm
        myLog("Returning %d " % result)
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