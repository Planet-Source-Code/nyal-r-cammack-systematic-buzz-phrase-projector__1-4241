VERSION 5.00
Begin VB.Form frmBuzzPhrase 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Buzz Phrase Projector"
   ClientHeight    =   960
   ClientLeft      =   1116
   ClientTop       =   2640
   ClientWidth     =   5892
   Icon            =   "BuzzPhrase.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   960
   ScaleWidth      =   5892
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   480
      Width           =   1104
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   360
      Left            =   2520
      TabIndex        =   4
      Top             =   480
      Width           =   1104
   End
   Begin VB.CommandButton cmdBuzz 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buzz Phrase"
      Default         =   -1  'True
      Height          =   360
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   1260
   End
   Begin VB.Label lblOne 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   324
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   1728
   End
   Begin VB.Label lblThree 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   324
      Left            =   4200
      TabIndex        =   2
      Top             =   0
      Width           =   1572
   End
   Begin VB.Label lblTwo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   324
      Left            =   1920
      TabIndex        =   1
      Top             =   0
      Width           =   2196
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuBuzz 
         Caption         =   "Buzz Phrase Projector"
      End
      Begin VB.Menu mnuPolar 
         Caption         =   "Polar Bear Software"
      End
   End
End
Attribute VB_Name = "frmBuzzPhrase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuzz_Click()
Dim inOne As Integer
Dim inTwo As Integer
Dim inThree As Integer
lblOne.Caption = ""
lblTwo.Caption = ""
lblThree.Caption = ""
Randomize

    inOne = ((10 - 1) * Rnd + 1)
    inTwo = ((10 - 1) * Rnd + 1)
    inThree = ((10 - 1) * Rnd + 1)
Select Case inOne
        Case 1
        lblOne.Caption = "Integrated"
        Case 2
        lblOne.Caption = "Total"
        Case 3
        lblOne.Caption = "Systemized"
        Case 4
        lblOne.Caption = "Parallel"
        Case 5
        lblOne.Caption = "Functional"
        Case 6
        lblOne.Caption = "Responsive"
        Case 7
        lblOne.Caption = "Optional"
        Case 8
        lblOne.Caption = "Synchronized"
        Case 9
        lblOne.Caption = "Compatible"
        Case 10
        lblOne.Caption = "Balanced"
    End Select
Select Case inTwo
        Case 1
        lblTwo.Caption = "Management"
        Case 2
        lblTwo.Caption = "Organizational"
        Case 3
        lblTwo.Caption = "Monitored"
        Case 4
        lblTwo.Caption = "Reciprocal"
        Case 5
        lblTwo.Caption = "Digital"
        Case 6
        lblTwo.Caption = "Logistical"
        Case 7
        lblTwo.Caption = "Transitional"
        Case 8
        lblTwo.Caption = "Incremental"
        Case 9
        lblTwo.Caption = "Third-Generation"
        Case 10
        lblTwo.Caption = "Policy"
    End Select
Select Case inThree
        Case 1
        lblThree.Caption = "Options"
        Case 2
        lblThree.Caption = "Flexibility"
        Case 3
        lblThree.Caption = "Capability"
        Case 4
        lblThree.Caption = "Mobility"
        Case 5
        lblThree.Caption = "Programming"
        Case 6
        lblThree.Caption = "Concept"
        Case 7
        lblThree.Caption = "Time-Phase"
        Case 8
        lblThree.Caption = "Projection"
        Case 9
        lblThree.Caption = "Hardware"
        Case 10
        lblThree.Caption = "Contingency"
    End Select
cmdOk.SetFocus

    
End Sub


Private Sub cmdExit_Click()
Unload Me

End Sub

Private Sub cmdOk_Click()
frmBuzzPhrase.WindowState = 1


End Sub

Private Sub Form_Load()
frmBuzzPhrase.Top = (Screen.Height - frmBuzzPhrase.Height) / 2
frmBuzzPhrase.Left = (Screen.Width - frmBuzzPhrase.Width) / 2

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Title As String
Dim Msg
Dim cr
cr = Chr(13) + Chr(10)
frmBuzzPhrase.Hide
Title = "Polar Bear Software"
Msg = "Thank you for using the Systematic Buzz Phrase Projector!"
Msg = Msg + cr + cr
Msg = Msg + "                             Polar Bear Software"
Msg = Msg + cr
Msg = Msg + "                                  Copyright 1999"
MsgBox Msg, vbExclamation, Title
End Sub

Private Sub Form_Resize()
lblOne.Caption = ""
lblTwo.Caption = ""
lblThree.Caption = ""
cmdBuzz.SetFocus
End Sub






Private Sub mnuBuzz_Click()
Dim Title As String
Dim Msg
Dim cr
cr = Chr(13) + Chr(10)
Title = "The Systematic Buzz Phrase Projector"
    Msg = "After many years of hacking through etymological thickets at the "
    Msg = Msg + "United States Public Health Service, an official named Phillip "
    Msg = Msg + "Broughton hit upon a sure fire method for converting frustration into fullfilment.  Termed the "
    Msg = Msg + "Systematic Buzz Phrase Projector, the system employs a lexicon "
    Msg = Msg + "of thirty selected Buzz Words. "
    Msg = Msg + cr + cr
    Msg = Msg + "The procedure is simple.  Just push the button and the "
    Msg = Msg + "Systematic Buzz Phrase Projector produces for you a phrase "
    Msg = Msg + "that can be dropped into virtually any report with a ring of decisive, knowledgeable, authority."
    Msg = Msg + cr + cr
    Msg = Msg + "No one will have the remotest idea of what you're talking about "
    Msg = Msg + "but the important thing is that they're not about to admit it!"
    MsgBox Msg, vbInformation, Title
    
    End Sub


Private Sub mnuPolar_Click()
Dim Title As String
Dim Msg
Dim cr
cr = Chr(13) + Chr(10)
Title = "About Polar Bear Software"
Msg = "All software published by Polar Bear Software is FreeShareWare.  "
Msg = Msg + "The idea behind FreeShareWare is that the software is free for "
Msg = Msg + "you to use as long as you share what you like/don't like about "
Msg = Msg + "the software with me."
Msg = Msg + cr + cr
Msg = Msg + "By sharing what you like/don't like I have the opportunity to "
Msg = Msg + "improve on the software, making the world a better place and your "
Msg = Msg + "life just a little bit easier.  You can EMail your suggestions to me at:"
Msg = Msg + cr + cr
Msg = Msg + "                   ncammack@efn.org   <Nyal R. Cammack>"
Msg = Msg + cr + cr
Msg = Msg + "                   Thank You for using Polar Bear Software!"
Msg = Msg + cr
Msg = Msg + "                                             Copyright 1999"
MsgBox Msg, vbInformation, Title

End Sub
