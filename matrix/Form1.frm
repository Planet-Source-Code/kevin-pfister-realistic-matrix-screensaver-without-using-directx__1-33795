VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmStartup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Matrix Settings"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Fading"
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   3000
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.CheckBox Check1 
      Caption         =   "From Top"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   2790
      Value           =   1  'Checked
      Width           =   1860
   End
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4605
      Begin VB.Timer TmrAutoStart 
         Interval        =   5000
         Left            =   4080
         Top             =   2400
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   285
         Left            =   90
         TabIndex        =   6
         Top             =   450
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   10
         Max             =   100
         SelStart        =   75
         TickStyle       =   3
         Value           =   100
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   285
         Left            =   90
         TabIndex        =   7
         Top             =   1035
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   150
         SelStart        =   30
         TickStyle       =   3
         Value           =   30
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   285
         Left            =   90
         TabIndex        =   8
         Top             =   1620
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   1000
         SelStart        =   20
         TickStyle       =   3
         Value           =   20
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   285
         Left            =   90
         TabIndex        =   9
         Top             =   2160
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         SelStart        =   4
         TickStyle       =   3
         Value           =   4
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fading Speed"
         Height          =   195
         Left            =   90
         TabIndex        =   10
         Top             =   1890
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Number of Dropping Columns"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   1350
         Width           =   2070
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Wait Before Clearing"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   765
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maximum Drop Length"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   180
         Width           =   1590
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   3375
      TabIndex        =   1
      Top             =   2835
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2070
      TabIndex        =   0
      Top             =   2835
      Width           =   1275
   End
End
Attribute VB_Name = "FrmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Slider4.Enabled = True
    Else
        Slider4.Enabled = False
    End If
End Sub

Private Sub Command1_Click()
    FrmMain.Show
    Call FrmMain.StartUp(Slider1.Value, Slider2.Value, Slider3.Value, Slider4.Value, Check1.Value, Check2.Value)
    Unload Me
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()
    If Trim(LCase(Command)) <> "/s" Then End    'Exit if not a screensaver
    If App.PrevInstance Then End    'Exit if there is a prev version running
End Sub

Private Sub TmrAutoStart_Timer()
    FrmMain.Show
    Call FrmMain.StartUp(Slider1.Value, Slider2.Value, Slider3.Value, Slider4.Value, Check1.Value, Check2.Value)
    Unload Me
End Sub
