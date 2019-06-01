@setlocal enabledelayedexpansion && python -x "%~f0" %* & exit /b !ERRORLEVEL!
# NAME: WINDOWS-TRIPWIRE-Test_port_response-1380.bat
# DESCRIPTION: Test TRIPWIRE port 9898 response
# USAGE: < script > [-s "path"] [-p port] [-h]
# WHERE:
#   path = path to text file containing the server names to check
#        = default is current server
#   port = defaults to 9898
#
# EXAMPLE:
#   < script >
#   < script > -s "c:\temp\serverlist.txt"
#   < script > -p 9898
#
# RETURN:
#   RBA success  - Success.
#   RBA diagnose - Failed.
#
# AUTHOR(s): Co, Harris <harris.co@hpe.com>
#
# DATE WRITTEN: 03 Mar 2015
# MODIFICATION HISTORY: 03 Mar 2015 - Initial Release
#

import telnetlib
import time
import socket
import os
import sys
import getopt

servers = []
port = ""
serverpath = ""
diagnosed = False

def usage():
    print 'DESCRIPTION: Test TRIPWIRE port response'
    print 'USAGE: < script > [-s "path"] [-p port] [-h]'
    print 'WHERE:'
    print '  path = path to text file containing the server names to check'
    print '       = default is current server'
    print '  port = defaults to 9898'
    print
    print 'EXAMPLE:'
    print '  < script >'
    print '  < script > -s "c:\\temp\\serverlist.txt"'
    print '  < script > -p 9898'

    sys.exit(0)

try:
    opts, args = getopt.getopt(sys.argv[1:], "s:p:h")
    for opt, arg in opts:
        if opt == '-s':
            serverpath = arg
        elif opt == '-p':
            port = arg
        elif opt == '-h':
            usage()
except getopt.GetoptError:
    pass

try:
    if os.stat(serverpath).st_size > 0:
        f = open(serverpath, 'r')
        servers = f.readlines()
        f.close()
except:
    pass

if (port == ''):
    port = 9898

if (len(servers) == 0):
    servers.append(socket.gethostname())

print 'RBA script stdout'
print 'WFAN=\"'

for server in servers:
    server = server.strip()

    if (server == ""):
        continue

    output = ""

    try:
        session = telnetlib.Telnet(server, port)
    except:
        output = str(sys.exc_info()[1])
        diagnosed = True
    else:
        time.sleep(1)
        output = session.read_until(" ", 8)
        output = ''.join(ch for ch in output if ch.isalnum())
        session.close()

    status = "Testing"
    #if ("ICA" in output):
    #    status = "Success"
    #else:
    #    status = "Failed"
    #    diagnosed = True

    print "Server: " + server
    print "Port  : " + str(port)
    print "Status: " + status
    print "Output: [" + output + "]"
    print

print '"'

if diagnosed:
    print 'RBA diagnose'
else:
    print 'RBA success'

sys.exit(0)
