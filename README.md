owexec
======

Tool from officewarfare.net to execute files on a remote comptuer (with admin priviledges) in the context of the running user

USAGE
owexec v-1.1 USAGE
owexec -c computername -k command [ -p parameters ] [ -u domain\user ] [ -copy ]
 [ -nowait ]

        -c the computer host name or ip of the target computer

        -k the command to be run, relative to the destination
           computer. ex: c:\windows\system32\notepad.exe

        -p the parameters to pass to the program, optional

        -u the user whose context the program should be run in
           if ommitted the first user that is found will be used

        -copy finds the command referenced with -k on the local
           machine and copies it to the comptuer referenced in
           -c on the admin$ share then runs it from there

        -nowait does not ask to press a key when the program finishes

download the current version at officewarfare.net

Example
=======
————– create FindNetUse.bat file ————-
@echo off
REM get net use
net use > c:\%USERNAME%-%COMPUTERNAME%-NetUse.txt”
copy c:\%USERNAME%-%COMPUTERNAME%-NetUse.txt j:\share\
del c:\%USERNAME%-%COMPUTERNAME%-NetUse.txt
———————————————

run the owexec command with -copy parameter:
owexec -nowait -k “NetUseRemoteBatch.bat” -copy -c “PC123″
