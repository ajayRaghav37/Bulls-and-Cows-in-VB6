VERSION 5.00
Begin VB.Form BCanal 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Bulls & Cows Analysis"
   ClientHeight    =   4590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4590
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAuto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Automatic Analysis"
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
      Height          =   255
      Left            =   600
      TabIndex        =   108
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   79
      Left            =   5160
      TabIndex        =   107
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   78
      Left            =   4680
      TabIndex        =   106
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   77
      Left            =   4200
      TabIndex        =   105
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   76
      Left            =   3720
      TabIndex        =   104
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   75
      Left            =   5160
      TabIndex        =   103
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   74
      Left            =   4680
      TabIndex        =   102
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   73
      Left            =   4200
      TabIndex        =   101
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   72
      Left            =   3720
      TabIndex        =   100
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   71
      Left            =   5160
      TabIndex        =   99
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   70
      Left            =   4680
      TabIndex        =   98
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   69
      Left            =   4200
      TabIndex        =   97
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   68
      Left            =   3720
      TabIndex        =   96
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   67
      Left            =   5160
      TabIndex        =   95
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   66
      Left            =   4680
      TabIndex        =   94
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   65
      Left            =   4200
      TabIndex        =   93
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   64
      Left            =   3720
      TabIndex        =   92
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   63
      Left            =   5160
      TabIndex        =   91
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   62
      Left            =   4680
      TabIndex        =   90
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   61
      Left            =   4200
      TabIndex        =   89
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   60
      Left            =   3720
      TabIndex        =   88
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   59
      Left            =   5160
      TabIndex        =   87
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   58
      Left            =   4680
      TabIndex        =   86
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   57
      Left            =   4200
      TabIndex        =   85
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   56
      Left            =   3720
      TabIndex        =   84
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   55
      Left            =   5160
      TabIndex        =   83
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   54
      Left            =   4680
      TabIndex        =   82
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   53
      Left            =   4200
      TabIndex        =   81
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   52
      Left            =   3720
      TabIndex        =   80
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   51
      Left            =   5160
      TabIndex        =   79
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   50
      Left            =   4680
      TabIndex        =   78
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   49
      Left            =   4200
      TabIndex        =   77
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   48
      Left            =   3720
      TabIndex        =   76
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   47
      Left            =   5160
      TabIndex        =   75
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   46
      Left            =   4680
      TabIndex        =   74
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   45
      Left            =   4200
      TabIndex        =   73
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   44
      Left            =   3720
      TabIndex        =   72
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   43
      Left            =   5160
      TabIndex        =   71
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   42
      Left            =   4680
      TabIndex        =   70
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   41
      Left            =   4200
      TabIndex        =   69
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   40
      Left            =   3720
      TabIndex        =   54
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   39
      Left            =   2400
      TabIndex        =   53
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   38
      Left            =   1920
      TabIndex        =   52
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   37
      Left            =   1440
      TabIndex        =   51
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   36
      Left            =   960
      TabIndex        =   50
      Top             =   3840
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   35
      Left            =   2400
      TabIndex        =   49
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   34
      Left            =   1920
      TabIndex        =   48
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   33
      Left            =   1440
      TabIndex        =   47
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   32
      Left            =   960
      TabIndex        =   46
      Top             =   3480
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   31
      Left            =   2400
      TabIndex        =   45
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   30
      Left            =   1920
      TabIndex        =   44
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   29
      Left            =   1440
      TabIndex        =   43
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   28
      Left            =   960
      TabIndex        =   42
      Top             =   3120
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   27
      Left            =   2400
      TabIndex        =   41
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   26
      Left            =   1920
      TabIndex        =   40
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   25
      Left            =   1440
      TabIndex        =   39
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   24
      Left            =   960
      TabIndex        =   38
      Top             =   2760
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   23
      Left            =   2400
      TabIndex        =   37
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   22
      Left            =   1920
      TabIndex        =   36
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   21
      Left            =   1440
      TabIndex        =   35
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   20
      Left            =   960
      TabIndex        =   34
      Top             =   2400
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   19
      Left            =   2400
      TabIndex        =   33
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   18
      Left            =   1920
      TabIndex        =   32
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   17
      Left            =   1440
      TabIndex        =   31
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   16
      Left            =   960
      TabIndex        =   30
      Top             =   2040
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   15
      Left            =   2400
      TabIndex        =   29
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   14
      Left            =   1920
      TabIndex        =   28
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   13
      Left            =   1440
      TabIndex        =   27
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   12
      Left            =   960
      TabIndex        =   26
      Top             =   1680
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   11
      Left            =   2400
      TabIndex        =   25
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   10
      Left            =   1920
      TabIndex        =   24
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   9
      Left            =   1440
      TabIndex        =   23
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   8
      Left            =   960
      TabIndex        =   22
      Top             =   1320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   2400
      TabIndex        =   21
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   1920
      TabIndex        =   20
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   1440
      TabIndex        =   19
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   4
      Left            =   960
      TabIndex        =   18
      Top             =   960
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   2400
      TabIndex        =   17
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   2
      Left            =   1920
      TabIndex        =   16
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   15
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox Possible 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.Label cmdCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Copy Unchecked"
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
      Height          =   255
      Left            =   3600
      TabIndex        =   110
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   4440
      Y2              =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   3240
      TabIndex        =   68
      Top             =   3840
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   3240
      TabIndex        =   67
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   3240
      TabIndex        =   66
      Top             =   3120
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   3240
      TabIndex        =   65
      Top             =   2760
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   3240
      TabIndex        =   64
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   3240
      TabIndex        =   63
      Top             =   2040
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   3240
      TabIndex        =   62
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   3240
      TabIndex        =   61
      Top             =   1320
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   3240
      TabIndex        =   60
      Top             =   960
      Width           =   120
   End
   Begin VB.Label Unchkr 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   3240
      TabIndex        =   59
      Top             =   600
      Width           =   120
   End
   Begin VB.Label Unchkc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4th"
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
      Index           =   3
      Left            =   5160
      TabIndex        =   58
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Unchkc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3rd"
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
      Index           =   2
      Left            =   4680
      TabIndex        =   57
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Unchkc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2nd"
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
      Index           =   1
      Left            =   4200
      TabIndex        =   56
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Unchkc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1st"
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
      Left            =   3720
      TabIndex        =   55
      Top             =   240
      Width           =   330
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   9
      Left            =   480
      TabIndex        =   14
      Top             =   3840
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   8
      Left            =   480
      TabIndex        =   13
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   7
      Left            =   480
      TabIndex        =   12
      Top             =   3120
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   6
      Left            =   480
      TabIndex        =   11
      Top             =   2760
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   2400
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   4
      Left            =   480
      TabIndex        =   9
      Top             =   2040
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   480
      TabIndex        =   7
      Top             =   1320
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      TabIndex        =   6
      Top             =   960
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   480
      TabIndex        =   5
      Top             =   600
      Width           =   120
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "4th"
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
      Index           =   13
      Left            =   2400
      TabIndex        =   4
      Top             =   240
      Width           =   330
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "3rd"
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
      Index           =   12
      Left            =   1920
      TabIndex        =   3
      Top             =   240
      Width           =   330
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "2nd"
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
      Index           =   11
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   330
   End
   Begin VB.Label lblJunk 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "1st"
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
      Index           =   10
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   330
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
      Height          =   4590
      Index           =   28
      Left            =   0
      TabIndex        =   109
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   5880
   End
End
Attribute VB_Name = "BCanal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
    Dim Temp As Integer
    For Temp = 0 To 39
        If Possible(Temp).Value = 0 Then
            Possible(Temp + 40).Value = 0
        End If
    Next
End Sub

Private Sub Unchkc_Click(Index As Integer)
    Dim Temp As Integer
    For Temp = (40 + Index) To (76 + Index) Step 4
        Possible(Temp).Value = IIf(Possible(Temp).Value = 0, 1, 0)
    Next
End Sub

Private Sub Unchkr_Click(Index As Integer)
    Dim Temp As Integer
    For Temp = (40 + 4 * Index) To (43 + 4 * Index)
        Possible(Temp).Value = IIf(Possible(Temp).Value = 0, 1, 0)
    Next
End Sub

Public Sub UnchkRow(Index As Integer)
    Dim Temp As Integer
    For Temp = 4 * Index To (3 + 4 * Index)
        Possible(Temp).Value = 0
    Next
End Sub

Public Sub UnchkCol(Index As Integer)
    Dim Temp As Integer
    For Temp = Index To (36 + Index) Step 4
        Possible(Temp).Value = 0
    Next
End Sub
