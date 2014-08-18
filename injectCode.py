#!/usr/bin/python 
# encoding: utf-8
'''
injectCode inject code from central base code to templates
'''

import sys
import os

import pythoncom
from lotusscript import SetupNotesClient
from toolbox import myLog, SetVerbosity
from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

__all__ = []

DEBUG =0 
TESTRUN = 0
PROFILE = 0


def UpgradeDesign(notesSession, sTargetRootDir, sSourceCode):
    try:    
    
        myLog("Upgrade design %s <= %s " % (sTargetRootDir, sSourceCode))
        
        sourceDb = notesSession.GetDatabase("", sSourceCode)
        srcDesignView = sourceDb.GetView("$Design")
        # Get projets 
        projects = sourceDb.GetView("Projects")
        if not projects:
            raise Exception('No projects found')
    
        # build all projects
        #for myProject in iterateDocuments(projects):
        myProject  = projects.GetFirstDocument()
        projectCount = 1
        while myProject:
            myLog("Project " + str(projectCount))
            DesignList = myProject.GetItemValue("ProjectDesigns")
            DatabasesList = myProject.GetItemValue("ProjectDatabases")
            myProject = projects.GetNextDocument(myProject)
            projectCount = projectCount + 1
            # iterate through notes databases
            for DbFilename in DatabasesList:
                # first open the db
                targetDb = notesSession.GetDatabase("",sTargetRootDir+"\\"+DbFilename)

                if not(targetDb.IsOpen):
                    myLog("Cannot open local database: %s\\%s" % (sTargetRootDir, DbFilename))
                    return False
                else:
                    myLog("Opened: %s\\%s" % (sTargetRootDir,DbFilename))
                    targetView = targetDb.GetView("$Design")
                    if targetView is None:
                        myLog("Cannot find $Design in %s" % DbFilename)
                        ForceCopy = True
                    else:
                        targetView.Refresh
                        ForceCopy = False
                    
                        # iterate through design documents
                    for designElt in DesignList:
                        myLog("Applying : %s" % designElt)
                        curDesignDoc = srcDesignView.GetDocumentByKey(designElt,True)
                        if not (curDesignDoc is None):
                            if not ForceCopy:
                                # try to find targetDesignDoc in targetView
                                key = curDesignDoc.GetItemValue("$Title")
                                targetDesignDoc = targetView.GetDocumentByKey(key,True)
                                if not (targetDesignDoc is None):
                                    targetDesignDoc.Remove(True)
                                
                            myLog("Design copy of %s" % designElt)
                            curDesignDoc.CopyToDatabase(targetDb)
                            # We need to copy it to the destination database
                        else:
                            myLog("Source design element %s not found" % designElt)
                            return False

        return True
    
    except Exception,err :
        print 'Error on line {}'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(err))
        
    except pythoncom.com_error as error:
        print (error)
        print (vars(error))
        print (error.args)

def main(argv=None): # IGNORE:C0111
    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    
    program_name = os.path.basename(sys.argv[0])

    try:
        # Setup argument parser
        parser = ArgumentParser(description="", formatter_class=RawDescriptionHelpFormatter)
        parser.add_argument("-n", "--name",    dest="name", action="store", help="application name]")
        parser.add_argument("-v", "--verbose", dest="verbose", action="count", help="set verbosity level [default: %(default)s]")
        parser.add_argument("-b", "--build",   dest="build", action="store", help="build number")
        parser.add_argument("-r", "--release", dest="version", help="version string (1.2.1)")
        parser.add_argument("-s", "--source", dest="source",action="store",help="Source of code")
        parser.add_argument("-d", "--dir", dest="dir",action="store",help="directory where to apply modifications")

        # Process arguments
        args = parser.parse_args()

        SetVerbosity(args.verbose)
        notesSession = SetupNotesClient(args.verbose)

        result = UpgradeDesign(notesSession, args.dir,args.source)
        myLog("Done - result " +str( result))
        
        return 0 if (result == True) else 1

    except KeyboardInterrupt:
        ### handle keyboard interrupt ###
        return 0
    except Exception, e:
        if DEBUG or TESTRUN:
            raise(e)
        indent = len(program_name) * " "
        sys.stderr.write(program_name + ": " + repr(e) + "\n")
        sys.stderr.write(indent + "  for help use --help")
        return 2

    
if __name__ == "__main__":
    if DEBUG:
        sys.argv.append("-r 1.2.3")
        sys.argv.append("-b 666")
        sys.argv.append("-v")
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