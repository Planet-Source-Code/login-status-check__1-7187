VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim bun As Boolean
inf.user = getusr()
inf.comp = getcomp
inf.logtime = Format(Time(), "hh:mm:ss")
inf.logdat = Format(Date, "yyyy.mm.dd")
aa = check4f()
aa = write2f()
bun = checkname()
If bun = True Then GoTo ende
MsgBox "Unaccepted user....", vbCritical, "Oooops!"
'you may insert a call function to restart, reboot or shutdown computer
ende:
Unload Me
End Sub
