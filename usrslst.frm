VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor for accepted users list..."
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "usrslst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Delete user"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Add user"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Accepted users:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Show vbModal
End Sub

Private Sub Command2_Click()
num = List1.SelCount
contor = 0
i = 0
While contor < num
    If List1.Selected(i) = True Then
        List1.RemoveItem (i)
        contor = contor + 1
        i = i - 1
    End If
    i = i + 1
    If i > List1.ListCount Then GoTo gata
Wend
gata:
List1.Refresh
End Sub

Private Sub Command3_Click()
usrfis = App.Path + "\accusrs.lst"
Open usrfis For Output As #1
    For i = 0 To List1.ListCount - 1
        Print #1, List1.List(i)
    Next i
Close #1
Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo err
usrfis = App.Path + "\accusrs.lst"
Open usrfis For Input As #1
While Not (EOF(1))
    Input #1, usrlst
    List1.AddItem usrlst
Wend
Close #1
GoTo ende
err:
Open usrfis For Output As #1
Close #1
ende:
End Sub
