VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "API_ProgressBar Demo"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   6525
   StartUpPosition =   2  'CenterScreen
   Begin API_ProgressBar.ctlProgressBar pb 
      Height          =   255
      Left            =   2640
      Top             =   720
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   450
      Appearance      =   1
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5640
      Top             =   1680
   End
   Begin VB.CommandButton btnGo 
      Caption         =   "Do stuff"
      Height          =   495
      Left            =   4440
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.ComboBox cboState 
      Height          =   315
      ItemData        =   "frmMain.frx":3452
      Left            =   480
      List            =   "frmMain.frx":345F
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3120
      Width           =   2055
   End
   Begin VB.ComboBox cboScrolling 
      Height          =   315
      ItemData        =   "frmMain.frx":349B
      Left            =   480
      List            =   "frmMain.frx":34A5
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   1920
      TabIndex        =   8
      Text            =   "0"
      Top             =   4440
      Width           =   615
   End
   Begin VB.TextBox txtMax 
      Height          =   285
      Left            =   1920
      TabIndex        =   7
      Text            =   "100"
      Top             =   4080
      Width           =   615
   End
   Begin VB.TextBox txtMin 
      Height          =   285
      Left            =   1920
      TabIndex        =   6
      Text            =   "0"
      Top             =   3720
      Width           =   615
   End
   Begin VB.CheckBox chkMarquee 
      Caption         =   "Marquee"
      Height          =   195
      Left            =   480
      TabIndex        =   5
      Top             =   3480
      Width           =   930
   End
   Begin VB.ComboBox cboBorderStyle 
      Height          =   315
      ItemData        =   "frmMain.frx":34D9
      Left            =   480
      List            =   "frmMain.frx":34E3
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1920
      Width           =   2055
   End
   Begin VB.ComboBox cboAppearance 
      Height          =   315
      ItemData        =   "frmMain.frx":3506
      Left            =   480
      List            =   "frmMain.frx":3510
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cboAlign 
      Height          =   315
      ItemData        =   "frmMain.frx":352A
      Left            =   480
      List            =   "frmMain.frx":353D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Note that to see XP or Vista visual styles, you will need to be using a style manifest.  See included readme for details."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Index           =   8
      Left            =   3720
      TabIndex        =   21
      Top             =   3720
      Width           =   2295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "XP and above only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   7
      Left            =   1440
      TabIndex        =   20
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Vista and above only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   19
      Top             =   2880
      Width           =   1755
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "State:"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   18
      Top             =   2880
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "No effect with XP and above styles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Index           =   4
      Left            =   1200
      TabIndex        =   17
      Top             =   2280
      Width           =   2895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Scrolling:"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   16
      Top             =   2280
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "BorderStyle:"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Appearance:"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Alignment:"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   13
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Value........................."
      Height          =   195
      Left            =   480
      TabIndex        =   12
      Top             =   4440
      Width           =   1890
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Max............................."
      Height          =   195
      Left            =   480
      TabIndex        =   11
      Top             =   4080
      Width           =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Min..........................."
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   3720
      Width           =   1860
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub btnGo_Click()
  Timer1.Enabled = Not (Timer1.Enabled)
  btnGo.Caption = IIf(Timer1.Enabled = True, "Stop doing stuff", "Do stuff")
End Sub

Private Sub cboAlign_Click()
  pb.Visible = False
  pb.Align = Right(cboAlign.Text, 1)
  If Right(cboAlign.Text, 1) = 3 Or Right(cboAlign.Text, 1) = 4 Then pb.Width = 300
  If Right(cboAlign.Text, 1) = 1 Or Right(cboAlign.Text, 1) = 2 Then pb.Height = 300
  If Right(cboAlign.Text, 1) = 0 Then
    pb.Move 2640, 720, 3375, 255
  End If
  pb.Visible = True
End Sub

Private Sub cboAppearance_Click()
  pb.Appearance = Right(cboAppearance.Text, 1)
End Sub

Private Sub cboBorderStyle_Click()
  pb.BorderStyle = Right(cboBorderStyle.Text, 1)
End Sub

Private Sub cboScrolling_Click()
  pb.Scrolling = Right(cboScrolling.Text, 1)
End Sub

Private Sub cboState_Click()
  pb.State = Right(cboState.Text, 1)
End Sub

Private Sub chkMarquee_Click()
  pb.Marquee = CBool(chkMarquee.Value)
End Sub

Private Sub Form_Load()
  cboAlign.ListIndex = 0
  cboAppearance.ListIndex = 1
  cboBorderStyle.ListIndex = 0
  cboScrolling.ListIndex = 0
  cboState.ListIndex = 0
End Sub

Private Sub Label4_Click(Index As Integer)
  chkMarquee.Value = Abs(chkMarquee.Value - 1)
End Sub

Private Sub Timer1_Timer()
  On Error Resume Next
  pb.Value = pb.Value + 5
  txtValue.Text = CStr(pb.Value)
End Sub

Private Sub txtValue_LostFocus()
  txtValue.Text = Val(txtValue.Text)
  pb.Value = Val(txtValue.Text)
End Sub

Private Sub txtMax_LostFocus()
  txtMax.Text = Val(txtMax.Text)
  pb.Max = Val(txtMax.Text)
End Sub

Private Sub txtMin_LostFocus()
  txtMin.Text = Val(txtMin.Text)
  pb.Min = Val(txtMin.Text)
End Sub
