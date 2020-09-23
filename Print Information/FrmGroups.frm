VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmGroups 
   Caption         =   "View & Print Groups"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   8415
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      Caption         =   "Save To File"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   18
      Top             =   5760
      Width           =   1215
   End
   Begin VB.OptionButton Option4 
      Caption         =   "Preview Only - All Groups and thier Members"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   17
      Top             =   2520
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   7560
      TabIndex        =   16
      Text            =   "10"
      Top             =   4680
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5400
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   14
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6360
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Print Selected Group with it's Members"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   1920
      Value           =   -1  'True
      Width           =   2775
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Print All Groups and their Members"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   11
      Top             =   1320
      Width           =   2775
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Print All Groups"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   10
      Top             =   720
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4320
      Width           =   6975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4800
      Top             =   120
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Groups"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3480
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2610
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Font Size"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Total Users:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Total Groups:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Enter or Select Domain,Computer Name or IP"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Select Group"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Users"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2610
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "FrmGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
On Error Resume Next
List1.Clear
List2.Clear
Label3.Caption = "All Users Of " & Combo1.Text
frmpleasewait.Show
DoEvents

Dim container As IADsContainer
Dim containername As String
containername = Combo1.Text
Set container = GetObject("WinNT://" & containername)

container.Filter = Array("User")
Dim user As IADsUser
For Each user In container
List1.AddItem user.Name
Next

container.Filter = Array("Group")
Dim group As IADsGroup
For Each group In container
List2.AddItem group.Name
Next

Err = 0
DoEvents
Unload frmpleasewait
End Sub

Private Sub Command2_Click()
On Error Resume Next
CD1.CancelError = False
If Option1.Value = True Then
Text1.Text = ""
Text1.Text = "(Domain, Computer Name or IP) - " & Combo1.Text & vbCrLf & vbCrLf
Text1.Text = Text1.Text & "(All Groups:)" & vbCrLf

Do Until List2.ListCount = 0
List2.ListIndex = 0
Text1.Text = Text1.Text & vbTab & List2.Text & vbCrLf
List2.RemoveItem List2.ListIndex
Loop
DoEvents
DoEvents
CD1.ShowPrinter
DoEvents
Printer.FontSize = Text2.Text
Printer.Print Text1.Text
DoEvents
Printer.EndDoc
Call Command1_Click
End If

If Option2.Value = True Then
Text1.Text = ""
Text1.Text = "(Domain, Computer Name or IP) - " & Combo1.Text & vbCrLf & vbCrLf

Do Until List2.ListCount = 0
List2.ListIndex = 0
Call List2_DblClick
DoEvents
DoEvents
Text1.Text = Text1.Text & "(Group) - " & List2.Text & vbCrLf
Text1.Text = Text1.Text & vbTab & "(Members:) - " & List1.ListCount & vbCrLf
DoEvents
Do Until List1.ListCount = 0
List1.ListIndex = 0
Text1.Text = Text1.Text & vbTab & vbTab & List1.Text & vbCrLf
List1.RemoveItem List1.ListIndex
Loop
Text1.Text = Text1.Text & vbCrLf
DoEvents
List2.RemoveItem List2.ListIndex
Loop

DoEvents
DoEvents
CD1.ShowPrinter
DoEvents
Printer.FontSize = Text2.Text
Printer.Print Text1.Text
DoEvents
Printer.EndDoc
Call Command1_Click
End If

If Option3.Value = True Then
Text1.Text = ""
Text1.Text = "(Domain, Computer Name or IP) - " & Combo1.Text & vbCrLf & vbCrLf
Text1.Text = Text1.Text & "(Group) - " & List2.Text & vbCrLf
Text1.Text = Text1.Text & vbTab & "(Members:)" & vbCrLf

Do Until List1.ListCount = 0
List1.ListIndex = 0
Text1.Text = Text1.Text & vbTab & vbTab & List1.Text & vbCrLf
List1.RemoveItem List1.ListIndex
Loop
DoEvents
DoEvents
CD1.ShowPrinter
DoEvents
Printer.FontSize = Text2.Text
Printer.Print Text1.Text
DoEvents
Printer.EndDoc
Call List2_DblClick
End If

If Option4.Value = True Then
Text1.Text = ""
Text1.Text = "(Domain, Computer Name or IP) - " & Combo1.Text & vbCrLf & vbCrLf

Do Until List2.ListCount = 0
List2.ListIndex = 0
Call List2_DblClick
DoEvents
DoEvents
Text1.Text = Text1.Text & "(Group) - " & List2.Text & vbCrLf
Text1.Text = Text1.Text & vbTab & "(Members:) - " & List1.ListCount & vbCrLf
DoEvents
Do Until List1.ListCount = 0
List1.ListIndex = 0
Text1.Text = Text1.Text & vbTab & vbTab & List1.Text & vbCrLf
List1.RemoveItem List1.ListIndex
Loop
Text1.Text = Text1.Text & vbCrLf
DoEvents
List2.RemoveItem List2.ListIndex
Loop
DoEvents
DoEvents
Call Command1_Click
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
On Error Resume Next
CD1.CancelError = False
CD1.Filter = "Text Document (*.txt)|*.txt"
CD1.ShowSave
  If CD1.FileName = "" Then Exit Sub
  Open CD1.FileName For Output As #1
  Print #1, Text1.Text
  Close #1

End Sub

Private Sub Form_Load()
Combo1.AddItem FrmMain.Winsock1.LocalHostName
Dim namespace As IADsContainer
Dim domain As IADs
 'Loads Combobox1 with all the current domains
Set namespace = GetObject("WinNT:")

For Each domain In namespace
Combo1.AddItem domain.Name
Next
End Sub



Private Sub List2_DblClick()
On Error Resume Next
List1.Clear
Label3.Caption = "Members of " & List2.Text
frmpleasewait.Show
DoEvents

Dim group As IADsGroup
Dim groupname As String
Dim groupdomain As String

groupname = List2.Text
groupdomain = Combo1.Text
Set group = GetObject("WinNT://" & groupdomain & "/" & groupname & ",group")

For Each member In group.Members
List1.AddItem member.Name
Next
Err = 0
DoEvents
Unload frmpleasewait
End Sub

Private Sub Timer1_Timer()
Label4.Caption = "Total Groups: " & List2.ListCount
Label5.Caption = "Total Users: " & List1.ListCount

If Combo1.Text = "" Then
Command1.Enabled = False
Else
Command1.Enabled = True
End If

If List2.ListCount = 0 Then
Command2.Enabled = False
Else
Command2.Enabled = True
End If
End Sub
