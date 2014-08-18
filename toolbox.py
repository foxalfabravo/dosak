#!/usr/bin/python 
# encoding: utf-8
'''
toolbox.py -- Miscellaneous tools 
'''
import os

verbose = 0

def myLog(sMessage):
    global verbose
    
    if verbose > 0:
        print sMessage

def SetVerbosity(newVerbose):
    global verbose
    verbose = newVerbose
    if verbose > 0:
        print("Verbose mode on")

def GetVerbosity():
    global verbose
    return verbose

# http://stackoverflow.com/questions/377017/test-if-executable-exists-in-python
def which(program):
    myLog("Looking for %s" % program)

    def is_exe(fpath):
        return os.path.isfile(fpath) and os.access(fpath, os.X_OK)

    fpath, fname = os.path.split(program)
    if fpath:
        if is_exe(program):
            return program
    else:
        for path in os.environ["PATH"].split(os.pathsep):
            path = path.strip('"')
            exe_file = os.path.join(path, program)
            if is_exe(exe_file):
                myLog("%s found in %s" % (program,path))
                return exe_file

    return None

# INI kzey modification
def setIniKey(sIniFile, sIniSection, sIniKey, sIniValue):
    import ConfigParser
    import sys
    try:
        Config = ConfigParser.ConfigParser()
        Config.read(sIniFile)
        Config.set(sIniSection,sIniKey,sIniValue)
                    
        with open(sIniFile,'w') as configFile:
            Config.write(configFile)

    except Exception, e:
        print 'Error on line {}\n'.format(sys.exc_info()[-1].tb_lineno)
        print ("Error: %s.\n" % str(e))
    
def checkProcessRunning(sProcessName):
    r = os.popen('tasklist /FI "STATUS eq Running"').read().strip().split('\n')
    found = False
    for p in r:
        if p.startswith(sProcessName):
            found = True
    return found
