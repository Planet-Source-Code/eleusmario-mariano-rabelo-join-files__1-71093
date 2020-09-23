VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Join Files"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   7890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtOutput 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   5415
   End
   Begin VB.CommandButton BtnJoin 
      Caption         =   "Join Files"
      Height          =   1095
      Left            =   2040
      TabIndex        =   5
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox TxtFile1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton BtnFile1 
      Caption         =   "Find"
      Height          =   375
      Left            =   5760
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.TextBox TxtFile2 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin VB.CommandButton BtnFile2 
      Caption         =   "Find"
      Height          =   375
      Left            =   5760
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   7200
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Output File"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "File 2"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "File 1"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub BtnFile1_Click()
CD.DialogTitle = "Find file"
CD.Filter = "All files *.*|*.*;"
CD.ShowOpen
TxtFile1.Text = CD.FileName
End Sub

Private Sub BtnFile2_Click()
CD.DialogTitle = "Find file"
CD.Filter = "All files *.*|*.*;"
CD.ShowOpen
TxtFile2.Text = CD.FileName
End Sub

Private Sub BtnJoin_Click()
Dim S As String, T As Long

BtnJoin.Enabled = False

Open TxtOutput.Text For Binary As #2

Open TxtFile1.Text For Binary As #1
T = LOF(1)
Do While T > 0
  If T > 6400000 Then
    S = Space(6400000)
    T = T - 6400000
  Else
    S = Space(T)
    T = 0
  End If
  Get #1, , S
  Put #2, , S
  DoEvents
Loop
Close #1

Open TxtFile2.Text For Binary As #1
T = LOF(1)
Do While T > 0
  If T > 6400000 Then
    S = Space(6400000)
    T = T - 6400000
  Else
    S = Space(T)
    T = 0
  End If
  Get #1, , S
  Put #2, , S
  DoEvents
Loop
Close #1

MsgBox "Files joined successfully!", vbInformation
BtnJoin.Enabled = True
End Sub

Private Sub TxtOutput_GotFocus()
TxtOutput.Text = Mid(TxtFile2.Text, 1, InStrRev(TxtFile2.Text, "\")) & "Output"
End Sub
