#!/usr/bin/python 
# encoding: utf-8
'''
Short : add version information in $templatebuild
'''

import sys
import os
from os import listdir

import pythoncom
from datetime import datetime
from lotusscript import SetupNotesClient
from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter
from toolbox import myLog, SetVerbosity

__all__ = []

DEBUG = 0
TESTRUN = 0
PROFILE = 0


def SetReleaseInformation(notesSession, sTemplateCompletePath, sName, sVersion, sBuildNumber, sBuildDate):
    try:    
    
        myLog("Adding release information to template %s (%s.%s - %s)" % (sTemplateCompletePath, sVersion, sBuildNumber, sBuildDate))
        cnvDate = datetime.strptime(sBuildDate.strip(),"%Y%m%d")   
        TheDate = notesSession.CreateDateTime(cnvDate.strftime("%m-%d-%Y"))

        targetDb = notesSession.GetDatabase("", sTemplateCompletePath)
        targetView = targetDb.GetView("$Design")
        if targetView is None:
            print "$Design not found"
        else:
            targetDesignDoc = targetView.GetDocumentByKey("$TemplateBuild",True)
            targetDesignDoc.ReplaceItemValue("$TemplateBuild",(sVersion + "." + sBuildNumber).split())
            targetDesignDoc.ReplaceItemValue("$move4ideasSoftware",sName)
            targetDesignDoc.ReplaceItemValue("$move4ideasBuild",sBuildNumber)
            targetDesignDoc.ReplaceItemValue("$move4ideasVersion",sVersion)
            targetDesignDoc.ReplaceItemValue("$TemplateBuildDate",TheDate.DateOnly)
        
            targetDesignDoc.Save(0,0)

    except pythoncom.com_error as error:
        print (error)
        print (vars(error))
        print (error.args)
        
def TagDirectory(notesSession, sDirectory, sName,sVersion,sBuild,sDate):

    myLog("Tagging %s directory" % sDirectory)
    
    # Get NTF/NSF list from directory
    included_ext = ['ntf'] ;
    fileLst = [ fn for fn in listdir(sDirectory) if any([fn.endswith(ext) for ext in included_ext])];
    
    for notesFile in fileLst:
        SetReleaseInformation(notesSession, sDirectory+"\\"+notesFile, sName, sVersion, sBuild, sDate)

            
def main(argv=None): # IGNORE:C0111

    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_name = os.path.basename(sys.argv[0])

    try:
        # Setup argument parser
        parser = ArgumentParser(description="", formatter_class=RawDescriptionHelpFormatter)
        parser.add_argument("-n", "--name",     dest="name", action="store", help="application name]")
        parser.add_argument("-v", "--verbose",  dest="verbose", action="count", help="set verbosity level [default: %(default)s]")
        parser.add_argument("-b", "--build",    dest="build", action="store", help="build number")
        parser.add_argument("-r", "--release",  dest="version", help="version string (1.2.1)")
        parser.add_argument("-t", "--template", dest="template",action="store",help="lotus template name")
        parser.add_argument("-d", "--date",     dest="date",action="store",help="build date")
        parser.add_argument("-f", "--fulldir",  dest="directory", action="store", help="full directory tagging")

        # Process arguments
        args = parser.parse_args()
        SetVerbosity(args.verbose)

        notesSession = SetupNotesClient(args.verbose)

        if (args.template <> None):
            SetReleaseInformation(notesSession, args.template.strip(),args.name.strip(),args.version.strip(),args.build.strip(),args.date.strip())
        elif (args.directory <> None):
            TagDirectory(notesSession, args.directory.strip(), args.name.strip(),args.version.strip(),args.build.strip(),args.date.strip())
        else:
            print "No input specified"
        return 0
    
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
        sys.argv.append("-d 20140302")
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