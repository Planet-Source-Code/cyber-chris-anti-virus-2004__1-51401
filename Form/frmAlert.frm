VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "CCAntivir 2004"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2325
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CommandButton cmdIgnore 
      Caption         =   "&Ignore"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   1800
      Width           =   1335
   End
   Begin VB.PictureBox picIcon 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   585
      Left            =   240
      ScaleHeight     =   585
      ScaleWidth      =   645
      TabIndex        =   4
      Top             =   120
      Width           =   645
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label lblText 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Virus found!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   (c) Copyright by Cyber Chris
'       Email: cyber_chris235@gmx.net
'
'   Please mail me when you want to use my code!
Option Explicit

Private Sub BuildAlert()

    lblText(1).Caption = Virus.FileName & "  (" & FileLen(Virus.FileName) & " Bytes )"
    lblText(1).ToolTipText = Virus.FileName & "  (" & FileLen(Virus.FileName) & " Bytes )"
    lblText(2).Caption = Virus.Reason
    picIcon.Picture = LoadIcon(Large, Virus.FileName)

End Sub

Private Sub cmdIgnore_Click()

    Unload Me

End Sub

Private Sub cmdRemove_Click()

    RemoveFile (Virus.FileName)

End Sub

Private Sub cmdView_Click()

    If MsgBox("WARNING! This will execute the file with the associated program !WARNING" & vbCrLf & "Continue?", vbCritical + vbYesNo, AV.AVname) = vbYes Then
        Call ShellExecute(Me.hwnd, "Open", Virus.FileName, vbNullString, "c:\", 1)
    End If

End Sub

Private Sub Form_Load()

    BuildAlert

End Sub

