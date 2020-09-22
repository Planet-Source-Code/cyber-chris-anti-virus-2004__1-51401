VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivirus 2004"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSearch.frx":0000
   ScaleHeight     =   2850
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picHelp 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "frmSearch.frx":10B6
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
   End
   Begin VB.PictureBox picUpdate 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "frmSearch.frx":31E0
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.PictureBox picAbout 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "frmSearch.frx":537E
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   0
      Top             =   2040
      Width           =   2055
   End
   Begin VB.PictureBox picFileSearch 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      Picture         =   "frmSearch.frx":8A10
      ScaleHeight     =   615
      ScaleWidth      =   2055
      TabIndex        =   1
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Run on startup"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   2640
      TabIndex        =   13
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4440
      TabIndex        =   12
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "OFF"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   4440
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tray window:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2640
      TabIndex        =   10
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   4440
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Signatures:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   2640
      TabIndex        =   6
      Top             =   600
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   0
      X1              =   2160
      X2              =   2160
      Y1              =   2760
      Y2              =   0
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   2
      X1              =   7680
      X2              =   2160
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Anti Virus Definitions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   2640
      TabIndex        =   4
      Top             =   240
      Width           =   1695
   End
   Begin VB.Line lines 
      BorderColor     =   &H00FFC0C0&
      Index           =   1
      X1              =   7680
      X2              =   2160
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0.0.0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   4440
      TabIndex        =   5
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Number of Files checked:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub Form_Load()

    frmMain.Cls
    BuildUI

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub lblText_Click(Index As Integer)

    On Error Resume Next
    If Index = 6 Then
        If lblText(7).Caption = "OFF" Then
            frmTray.Show , Me
         Else 'NOT LBLTEXT(7).CAPTION...
            Unload frmTray
        End If
     ElseIf Index = 9 Then 'NOT INDEX...
        If lblText(8).Caption = "OFF" Then
            SetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname, App.Path & "\" & App.EXEName & ".exe /T", 1
            lblText(8).Caption = "ON"
            SaveSetting AV.AVname, "Settings", "Startup", "ON"
         Else 'NOT LBLTEXT(8).CAPTION...
            DeleteValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", AV.AVname
            lblText(8).Caption = "OFF"
            SaveSetting AV.AVname, "Settings", "Startup", "OFF"
        End If
    End If
    On Error GoTo 0

End Sub

Private Sub picAbout_Click()

    frmAbout.Show , Me

End Sub

Private Sub picFileSearch_Click()

    Call ShowFileSearch

End Sub

Private Sub picHelp_Click()

    frmHelp.Show , Me

End Sub

Private Sub picUpdate_Click()

    frmUpdate.Show , Me

End Sub

Public Sub ShowFileSearch()

  Dim strFilename As String

    On Error Resume Next
    strFilename = (ShowOpenDlg(Me, , "All Files|*.*", , "Scan File"))
    If FileLen(strFilename) <> 0 Then
        CheckFile (strFilename)
    End If
    On Error GoTo 0

End Sub

