#!/usr/bin/python 
# encoding: utf-8
'''
Short : rebuild NxF from On Disk structure project  
'''

import sys
import os
import tempfile
import subprocess
import time
from lotusscript import *
from toolbox import myLog, SetVerbosity
from argparse import ArgumentParser
from argparse import RawDescriptionHelpFormatter

__all__ = []
__version__ = 0.1

DEBUG = 1
TESTRUN = 0
PROFILE = 0

def ods2nxf(sDesigner, sFrom, sTo):
    # Make tmp command file
    bResult = True;
    prjFile = os.path.join(sFrom, ".project")
    path2Result, resultName= os.path.split(sTo)
    path2designer, desName = os.path.split(sDesigner)
    
    with tempfile.NamedTemporaryFile(delete = False) as commandFile:
#         commandFile.write('config,true,true\n')   # Stop on error / Exit on error
#         commandFile.write('importandbuild,%s,%s\n' % (prjFile,sTo) )
#         commandFile.write('clean\n')
#         commandFile.write('exit\n')
#         commandFile.flush()
# 
#         sInvoke = "%s -RPARAMS -vmargs  -Dcom.ibm.designer.cmd.file=%s" % (sDesigner,commandFile.name)
        sInvoke = "%s -RPARAMS -vmargs  -Dcom.ibm.designer.cmd=true,true,%s,importandbuild,%s,%s" % (sDesigner,resultName,prjFile,resultName)
        myLog("Invoking %s" % sInvoke)
        subprocess.call(sInvoke) 
        while isDesignerRunning():
            time.sleep(1)
        myLog("End of designer build process")
        # Check result
        for line in open("HEADLESS0.log"):
            if "job error" in line:
                print line
                bResult = False;
        
        if (bResult == True):
            sTempResult ="%s/Data/%s" % (path2designer,resultName) 
            os.rename(sTempResult, sTo)
        else:
            os.remove(sTempResult)
        return bResult 
    
def main(argv=None): # IGNORE:C0111
    if argv is None:
        argv = sys.argv
    else:
        sys.argv.extend(argv)

    program_name = os.path.basename(sys.argv[0])

    try:
        # Setup argument parser
        parser = ArgumentParser(description="", formatter_class=RawDescriptionHelpFormatter)
        parser.add_argument("-f", "--from",    dest="src",      action="store", help="From source On disk structure (full path)")
        parser.add_argument("-t", "--to",      dest="dst",      action="store", help="To NSF/NTF result (full path)")
        parser.add_argument("-u", "--user",    dest="user",     action="store", help="user id short filename (relative to Notes/Data)")
        parser.add_argument("-v", "--verbose", dest="verbose",  action="count", help="set verbosity level [default: %(default)s]")

        # Process arguments
        args = parser.parse_args()

        SetVerbosity(args.verbose)
        myLog("Args are %s " % args)

        if isDesignerRunning():
            print("Cannot run while notes is already running")
            return 2
        
        if (args.user is not None):
            SwitchUser(args.user)
        
        designer = FindDesigner()
        if (designer is not None):
            path2designer, desName = os.path.split(designer)
            EnableHeadlessDesigner(path2designer)
            result = ods2nxf(designer, args.src, args.dst)
        else:
            result = False;
                
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