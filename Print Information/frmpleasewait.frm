VERSION 5.00
Begin VB.Form frmpleasewait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please wait....."
   ClientHeight    =   1290
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   2235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      Caption         =   "Please wait this will take a few moments....."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   150
      TabIndex        =   0
      Top             =   158
      Width           =   1935
   End
End
Attribute VB_Name = "frmpleasewait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
