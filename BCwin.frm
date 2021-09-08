VERSION 5.00
Begin VB.Form BCwin 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Bulls And Cows"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Guess4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      ItemData        =   "BCwin.frx":0000
      Left            =   3480
      List            =   "BCwin.frx":0022
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   1320
      Width           =   540
   End
   Begin VB.ComboBox Guess3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      ItemData        =   "BCwin.frx":0044
      Left            =   2880
      List            =   "BCwin.frx":0066
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   1320
      Width           =   540
   End
   Begin VB.ComboBox Guess2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      ItemData        =   "BCwin.frx":0088
      Left            =   2280
      List            =   "BCwin.frx":00AA
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   1320
      Width           =   540
   End
   Begin VB.ComboBox Guess1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      ItemData        =   "BCwin.frx":00CC
      Left            =   1680
      List            =   "BCwin.frx":00EE
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1320
      Width           =   540
   End
   Begin VB.CommandButton cmdCheck 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "_"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5130
      TabIndex        =   20
      Top             =   30
      Width           =   150
   End
   Begin VB.Label cmdAnal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Analysis"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   4200
      TabIndex        =   18
      Top             =   120
      Width           =   750
   End
   Begin VB.Line lnJunk 
      X1              =   0
      X2              =   5790
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label cmdClose 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5370
      TabIndex        =   17
      Top             =   120
      Width           =   180
   End
   Begin VB.Label lblCows 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "•"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   5040
      TabIndex        =   10
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label lblBulls 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "•"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   4320
      TabIndex        =   9
      Top             =   1320
      Width           =   510
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bulls"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   4320
      TabIndex        =   8
      Top             =   795
      Width           =   510
   End
   Begin VB.Label lblJunk 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Cows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   5040
      TabIndex        =   7
      Top             =   795
      Width           =   510
   End
   Begin VB.Label lblAttempt 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attempt 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   1395
      Width           =   855
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Secret Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   795
      Width           =   1305
   End
   Begin VB.Label rndNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   510
   End
   Begin VB.Label rndNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   720
      Width           =   510
   End
   Begin VB.Label rndNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   1
      Left            =   2280
      TabIndex        =   2
      Top             =   720
      Width           =   510
   End
   Begin VB.Label rndNumber 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   390
      Index           =   0
      Left            =   1680
      TabIndex        =   1
      Top             =   720
      Width           =   510
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   5775
   End
   Begin VB.Label lblJunk 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1950
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5790
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bulls & Cows"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   1380
   End
End
Attribute VB_Name = "BCwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SecretNumber(3) As String
Dim MaxIndex As Integer
Dim Xinit As Integer
Dim Yinit As Integer

Private Sub cmdAnal_Click()
    If BCanal.Visible Then
        Move Left + (BCanal.Width / 2)
        BCanal.Hide
    Else
        Move Left - (BCanal.Width / 2)
        BCanal.Move Left + Width, Top
        BCanal.Show
    End If
End Sub

Private Sub cmdCheck_Click()

'Checking for Bulls and Cows
    
    Select Case Guess1(MaxIndex).Text
        Case SecretNumber(0)
            lblBulls(MaxIndex).Caption = Trim$(Val(lblBulls(MaxIndex).Caption) + 1)
        Case SecretNumber(1), SecretNumber(2), SecretNumber(3)
            lblCows(MaxIndex).Caption = Trim$(Val(lblCows(MaxIndex).Caption) + 1)
    End Select
    Select Case Guess2(MaxIndex).Text
        Case Guess1(MaxIndex).Text
            MsgBox "All digits must be unique."
            lblBulls(MaxIndex).Caption = "•"
            lblCows(MaxIndex).Caption = "•"
            Exit Sub
        Case SecretNumber(1)
            lblBulls(MaxIndex).Caption = Trim$(Val(lblBulls(MaxIndex).Caption) + 1)
        Case SecretNumber(0), SecretNumber(2), SecretNumber(3)
            lblCows(MaxIndex).Caption = Trim$(Val(lblCows(MaxIndex).Caption) + 1)
    End Select
    Select Case Guess3(MaxIndex).Text
        Case Guess1(MaxIndex).Text, Guess2(MaxIndex).Text
            MsgBox "All digits must be unique."
            lblBulls(MaxIndex).Caption = "•"
            lblCows(MaxIndex).Caption = "•"
            Exit Sub
        Case SecretNumber(2)
            lblBulls(MaxIndex).Caption = Trim$(Val(lblBulls(MaxIndex).Caption) + 1)
        Case SecretNumber(0), SecretNumber(1), SecretNumber(3)
            lblCows(MaxIndex).Caption = Trim$(Val(lblCows(MaxIndex).Caption) + 1)
    End Select
    Select Case Guess4(MaxIndex).Text
        Case Guess1(MaxIndex).Text, Guess2(MaxIndex).Text, Guess3(MaxIndex).Text
            MsgBox "All digits must be unique."
            lblBulls(MaxIndex).Caption = "•"
            lblCows(MaxIndex).Caption = "•"
            Exit Sub
        Case SecretNumber(3)
            lblBulls(MaxIndex).Caption = Trim$(Val(lblBulls(MaxIndex).Caption) + 1)
        Case SecretNumber(0), SecretNumber(1), SecretNumber(2)
            lblCows(MaxIndex).Caption = Trim$(Val(lblCows(MaxIndex).Caption) + 1)
    End Select
    CheckResult
End Sub

Private Sub CheckResult()
    If lblBulls(MaxIndex).Caption = "4" Then
        MsgBox "All bulls discovered in " & Trim$(MaxIndex + 1) & " moves"
        ResetAll
    ElseIf MaxIndex = 9 Then
        MsgBox "Failed to discover? The secret number was " & SecretNumber(0) & SecretNumber(1) & SecretNumber(2) & SecretNumber(3)
        ResetAll
    Else
    'Creating new controls, positioning them and locking previous controls
        If BCanal.chkAuto.Value = 1 Then
            Analyze
        End If
        MaxIndex = MaxIndex + 1
        Load Guess1(MaxIndex)
        Load Guess2(MaxIndex)
        Load Guess3(MaxIndex)
        Load Guess4(MaxIndex)
        Load lblAttempt(MaxIndex)
        Load lblBulls(MaxIndex)
        Load lblCows(MaxIndex)
        Guess1(MaxIndex).Move Guess1(MaxIndex - 1).Left, Guess1(MaxIndex - 1).Top + 510
        Guess2(MaxIndex).Move Guess2(MaxIndex - 1).Left, Guess2(MaxIndex - 1).Top + 510
        Guess3(MaxIndex).Move Guess3(MaxIndex - 1).Left, Guess3(MaxIndex - 1).Top + 510
        Guess4(MaxIndex).Move Guess4(MaxIndex - 1).Left, Guess4(MaxIndex - 1).Top + 510
        lblAttempt(MaxIndex).Move lblAttempt(MaxIndex - 1).Left, lblAttempt(MaxIndex - 1).Top + 510
        lblBulls(MaxIndex).Move lblBulls(MaxIndex - 1).Left, lblBulls(MaxIndex - 1).Top + 510
        lblCows(MaxIndex).Move lblCows(MaxIndex - 1).Left, lblCows(MaxIndex - 1).Top + 510
        cmdCheck.Top = cmdCheck.Top + 510
        Top = Top - 255
        BCanal.Top = Top
        Height = Height + 510
        lblJunk(4).Height = Height
        lblAttempt(MaxIndex).Caption = "Attempt " & Trim$(MaxIndex + 1)
        lblBulls(MaxIndex).Caption = "•"
        lblCows(MaxIndex).Caption = "•"
        Guess1(MaxIndex - 1).Enabled = False
        Guess2(MaxIndex - 1).Enabled = False
        Guess3(MaxIndex - 1).Enabled = False
        Guess4(MaxIndex - 1).Enabled = False
        Guess1(MaxIndex).AddItem "0"
        Guess1(MaxIndex).AddItem "1"
        Guess1(MaxIndex).AddItem "2"
        Guess1(MaxIndex).AddItem "3"
        Guess1(MaxIndex).AddItem "4"
        Guess1(MaxIndex).AddItem "5"
        Guess1(MaxIndex).AddItem "6"
        Guess1(MaxIndex).AddItem "7"
        Guess1(MaxIndex).AddItem "8"
        Guess1(MaxIndex).AddItem "9"
        
        Guess2(MaxIndex).AddItem "0"
        Guess2(MaxIndex).AddItem "1"
        Guess2(MaxIndex).AddItem "2"
        Guess2(MaxIndex).AddItem "3"
        Guess2(MaxIndex).AddItem "4"
        Guess2(MaxIndex).AddItem "5"
        Guess2(MaxIndex).AddItem "6"
        Guess2(MaxIndex).AddItem "7"
        Guess2(MaxIndex).AddItem "8"
        Guess2(MaxIndex).AddItem "9"
        
        Guess3(MaxIndex).AddItem "0"
        Guess3(MaxIndex).AddItem "1"
        Guess3(MaxIndex).AddItem "2"
        Guess3(MaxIndex).AddItem "3"
        Guess3(MaxIndex).AddItem "4"
        Guess3(MaxIndex).AddItem "5"
        Guess3(MaxIndex).AddItem "6"
        Guess3(MaxIndex).AddItem "7"
        Guess3(MaxIndex).AddItem "8"
        Guess3(MaxIndex).AddItem "9"
        
        Guess4(MaxIndex).AddItem "0"
        Guess4(MaxIndex).AddItem "1"
        Guess4(MaxIndex).AddItem "2"
        Guess4(MaxIndex).AddItem "3"
        Guess4(MaxIndex).AddItem "4"
        Guess4(MaxIndex).AddItem "5"
        Guess4(MaxIndex).AddItem "6"
        Guess4(MaxIndex).AddItem "7"
        Guess4(MaxIndex).AddItem "8"
        Guess4(MaxIndex).AddItem "9"
                
        Guess1(MaxIndex).Text = "1"
        Guess2(MaxIndex).Text = "2"
        Guess3(MaxIndex).Text = "3"
        Guess4(MaxIndex).Text = "4"
        
        lblAttempt(MaxIndex).Visible = True
        lblBulls(MaxIndex).Visible = True
        lblCows(MaxIndex).Visible = True
        Guess1(MaxIndex).Enabled = True
        Guess2(MaxIndex).Enabled = True
        Guess3(MaxIndex).Enabled = True
        Guess4(MaxIndex).Enabled = True
        Guess1(MaxIndex).Visible = True
        Guess2(MaxIndex).Visible = True
        Guess3(MaxIndex).Visible = True
        Guess4(MaxIndex).Visible = True
        Guess1(MaxIndex).SetFocus
    End If
End Sub
Private Sub cmdClose_Click()
    End
End Sub

Private Sub Form_Load()
    Guess1(0).Text = "1"
    Guess2(0).Text = "2"
    Guess3(0).Text = "3"
    Guess4(0).Text = "4"
    Load BCanal
    Show
    Visualize
    Guess1(0).Enabled = True
    Guess2(0).Enabled = True
    Guess3(0).Enabled = True
    Guess4(0).Enabled = True
    cmdCheck.Enabled = True
    Randomize
    SecretNumber(0) = Trim$(Int(9 * Rnd) + 1)
    Do
        SecretNumber(1) = Trim$(Int(10 * Rnd))
    Loop Until SecretNumber(1) <> SecretNumber(0)
    Do
        SecretNumber(2) = Trim$(Int(10 * Rnd))
    Loop Until SecretNumber(2) <> SecretNumber(1) And SecretNumber(2) <> SecretNumber(0)
    Do
        SecretNumber(3) = Trim$(Int(10 * Rnd))
    Loop Until SecretNumber(3) <> SecretNumber(2) And SecretNumber(3) <> SecretNumber(1) And SecretNumber(3) <> SecretNumber(0)
    Guess1(0).SetFocus
End Sub

Private Sub Guess1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If Guess1(Index).Text = "9" Then
                Guess1(Index).ListIndex = 0
                KeyCode = 0
            End If
        Case vbKeyUp
            If Guess1(Index).Text = "0" Then
                Guess1(Index).ListIndex = 9
                KeyCode = 0
            End If
        Case vbKeyRight, vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
            Guess2(Index).SetFocus
            KeyCode = 0
        Case vbKeyLeft, vbKeyClear, vbKeyBack
            Guess4(Index).SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub Guess2_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If Guess2(Index).Text = "9" Then
                Guess2(Index).ListIndex = 0
                KeyCode = 0
            End If
        Case vbKeyUp
            If Guess2(Index).Text = "0" Then
                Guess2(Index).ListIndex = 9
                KeyCode = 0
            End If
        Case vbKeyRight, vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
            Guess3(Index).SetFocus
            KeyCode = 0
        Case vbKeyLeft, vbKeyClear, vbKeyBack
            Guess1(Index).SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub Guess3_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If Guess3(Index).Text = "9" Then
                Guess3(Index).ListIndex = 0
                KeyCode = 0
            End If
        Case vbKeyUp
            If Guess3(Index).Text = "0" Then
                Guess3(Index).ListIndex = 9
                KeyCode = 0
            End If
        Case vbKeyRight, vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
            Guess4(Index).SetFocus
            KeyCode = 0
        Case vbKeyLeft, vbKeyClear, vbKeyBack
            Guess2(Index).SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub Guess4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            If Guess4(Index).Text = "9" Then
                Guess4(Index).ListIndex = 0
                KeyCode = 0
            End If
        Case vbKeyUp
            If Guess4(Index).Text = "0" Then
                Guess4(Index).ListIndex = 9
                KeyCode = 0
            End If
        Case vbKeyRight, vbKey0, vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKeyNumpad0, vbKeyNumpad1, vbKeyNumpad2, vbKeyNumpad3, vbKeyNumpad4, vbKeyNumpad5, vbKeyNumpad6, vbKeyNumpad7, vbKeyNumpad8, vbKeyNumpad9
            Guess1(Index).SetFocus
            KeyCode = 0
        Case vbKeyLeft, vbKeyClear, vbKeyBack
            Guess3(Index).SetFocus
            KeyCode = 0
    End Select
End Sub

Private Sub ResetAll()
    Dim Temp As Integer
    For Temp = MaxIndex To 1 Step -1
        Unload lblAttempt(Temp)
        Unload lblBulls(Temp)
        Unload lblCows(Temp)
        Unload Guess1(Temp)
        Unload Guess2(Temp)
        Unload Guess3(Temp)
        Unload Guess4(Temp)
    Next
    MaxIndex = 0
    Height = 1950
    lblJunk(4).Height = Height
    lblBulls(0).Caption = "•"
    lblCows(0).Caption = "•"
    cmdCheck.Top = Guess1(0).Top
    Show
    Unload BCanal
    cmdCheck.Enabled = False
    Form_Load
End Sub

Private Sub Visualize()
    Dim oTimer As Single
    Dim mTimer As Single
    oTimer = Timer
    Randomize
    rndNumber(0).Caption = "4"
    rndNumber(1).Caption = "2"
    rndNumber(2).Caption = "9"
    rndNumber(3).Caption = "6"
    Do
        rndNumber(0).Caption = Trim$(Int(9 * Rnd) + 1)
        rndNumber(1).Caption = Trim$(Int(10 * Rnd))
        rndNumber(2).Caption = Trim$(Int(10 * Rnd))
        rndNumber(3).Caption = Trim$(Int(10 * Rnd))
        DoEvents
        mTimer = Timer
        Do
        Loop Until Timer - mTimer > 1 / 100
    Loop Until Timer - oTimer > 1
    rndNumber(0).Caption = "?"
    rndNumber(1).Caption = "?"
    rndNumber(2).Caption = "?"
    rndNumber(3).Caption = "?"
End Sub

Private Sub Analyze()
    Dim Temp As Integer
    Dim Temp2 As Integer
    For Temp = 0 To MaxIndex
        If Val(lblCows(Temp).Caption) + Val(lblBulls(Temp).Caption) = 4 Then
            For Temp2 = 0 To 9
                If Temp2 <> Val(Guess1(Temp).Text) And Temp2 <> Val(Guess2(Temp).Text) And Temp2 <> Val(Guess3(Temp).Text) And Temp2 <> Val(Guess4(Temp).Text) Then
                    BCanal.UnchkRow Temp2
                End If
            Next
        End If
        If Val(lblCows(Temp).Caption) + Val(lblBulls(Temp).Caption) = 0 Then
            For Temp2 = 0 To 9
                If Temp2 = Val(Guess1(Temp).Text) Or Temp2 = Val(Guess2(Temp).Text) Or Temp2 = Val(Guess3(Temp).Text) Or Temp2 = Val(Guess4(Temp).Text) Then
                    BCanal.UnchkRow Temp2
                End If
            Next
        End If
        If Val(lblBulls(Temp).Caption) = 0 Then
            BCanal.Possible(4 * Val(Guess1(Temp).Text)).Value = 0
            BCanal.Possible((4 * Val(Guess2(Temp).Text)) + 1).Value = 0
            BCanal.Possible((4 * Val(Guess3(Temp).Text)) + 2).Value = 0
            BCanal.Possible((4 * Val(Guess4(Temp).Text)) + 3).Value = 0
        End If
    Next
End Sub

Private Sub Label1_Click()
    WindowState = 1
    If BCanal.Visible Then
        BCanal.Hide
    End If
End Sub

Private Sub lblTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Xinit = X
        Yinit = Y
    End If
End Sub

Private Sub lblTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Move Left + X - Xinit, Top + Y - Yinit
        If BCanal.Visible Then
            BCanal.Move BCanal.Left + X - Xinit, BCanal.Top + Y - Yinit
        End If
    End If
End Sub
