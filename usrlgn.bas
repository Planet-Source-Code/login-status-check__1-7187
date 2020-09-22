Attribute VB_Name = "Module1"
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Type INFOS
    user As String
    comp As String
    logdat As String
    logtime As String
End Type
Global inf As INFOS
Global user As String
Function getusr() 'get the user name
Dim usr As String

usr = Space(256)
aa = GetUserName(usr, 256)
getusr = Left(RTrim(usr), Len(RTrim(usr)) - 1)

End Function
Function getcomp() 'get the computer name
Dim comp As String

comp = Space(256)
aa = GetComputerName(comp, 256)
getcomp = Left(RTrim(comp), Len(RTrim(comp)) - 1)

End Function
Function write2f() 'write login status to "users.lgs" file
numfis = App.Path + "\" + "users.lgs"
tempf = App.Path + "\" + "users.tmp"
Open numfis For Input As #1
Open tempf For Output As #2
While Not (EOF(1))
    Input #1, txt
    Print #2, txt
Wend
Print #2, inf.user + vbTab + vbTab + vbTab + inf.logdat + vbTab + vbTab + inf.logtime
Close #2
Close #1
Kill numfis
Name tempf As numfis
End Function

Function check4f() 'check if the file exists...
numfis = App.Path + "\users.lgs"

On Error GoTo err 'if the file doesn't exist, will go to err
Open numfis For Input As #1
Close #1
GoTo ende

err: 'create the file
Open numfis For Output As #1
    Print #1, "Login status, on computer " + inf.comp + ":"
    Print #1, "----------------------------------------------------------"
    Print #1, "User:" + vbTab + vbTab + vbTab + "Login date:" + vbTab + vbTab + vbTab + "Login time:"
Close #1

ende:
End Function

Function checkname() 'checks if the user is in the accepted-users list
numefis = App.Path + "\accusrs.lst"
Dim found As Boolean
found = False
On Error GoTo err 'if the file doesn't exist, will create it at "err:"
Open numefis For Input As #1
While Not (EOF(1))
    Input #1, user1
    If LCase(user1) = LCase(inf.user) Then found = True
Wend
Close #1
checkname = found
GoTo ende
err:
Open numefis For Output As #1
Close #1
MsgBox "Users list not created...", vbOKOnly, "Users list..."
ende:
End Function
