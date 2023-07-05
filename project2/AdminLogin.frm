VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Login 
   BorderStyle     =   0  'None
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   ScaleHeight     =   4845
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc ADODB 
      Height          =   375
      Left            =   2760
      Top             =   4200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"AdminLogin.frx":0000
      OLEDBString     =   $"AdminLogin.frx":008B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "OfficeLog"
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   3615
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   6735
      Begin VB.TextBox txtUsername1 
         Height          =   495
         Left            =   2520
         TabIndex        =   5
         Top             =   480
         Width           =   3735
      End
      Begin VB.TextBox txtPassword1 
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H0000C000&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   3
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H0000C000&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3360
         TabIndex        =   2
         Top             =   2280
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080FF80&
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FF80&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1440
         Width           =   1215
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3975
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   7011
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Login"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdLogin_Click()
On Error Resume Next
rs.MoveFirst
rs.Find ("Username='" & txtUsername1 & "'")
If Not rs.EOF Then
If rs("Password='" & txtPassword1 & "'") Then
MsgBox "Well Done..Login Successful"
Registration.Show
Else
MsgBox "Login Failed...Please login with correct password"
Frame1.Visible = False
End If
Else
MsgBox "Login Failed...Please login with correct username"
Frame1.Visible = False
End If
End Sub


Private Sub Form_Load()
Frame1.Visible = True
Me.Left = (Screen.Width - Me.Width) / 5
Me.Top = (Screen.Height - Me.Height) / 5
cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=StudentsReg;Data Source=CLASSPC\SQLEXPRESS2019"
rs.Open "OfficeLog", cn, adOpenDynamic, adLockOptimistic
End Sub


Private Sub TabStrip1_Click()

If TabStrip1.Tabs(1).Selected = True Then
Frame1.Visible = True
End If
End Sub

