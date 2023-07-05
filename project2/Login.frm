VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H008080FF&
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   6615
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H00000080&
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
         Height          =   495
         Left            =   3240
         TabIndex        =   13
         Top             =   2400
         Width           =   1935
      End
      Begin VB.CommandButton cmdSign 
         BackColor       =   &H00000080&
         Caption         =   "Sign Up"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtPassword3 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtPassword2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   2520
         TabIndex        =   10
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox txtUser 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label5 
         BackColor       =   &H008080FF&
         Caption         =   "Re-type Password"
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
         Left            =   240
         TabIndex        =   16
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackColor       =   &H008080FF&
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
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
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
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.PictureBox ADODB 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   3240
      ScaleHeight     =   435
      ScaleWidth      =   1995
      TabIndex        =   17
      Top             =   4560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      Begin VB.CommandButton cmdSignUp 
         BackColor       =   &H0000C000&
         Caption         =   "Sign Up"
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
         TabIndex        =   7
         Top             =   2280
         Width           =   2175
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
         TabIndex        =   6
         Top             =   2280
         Width           =   1935
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
      Begin VB.TextBox txtUsername1 
         Height          =   495
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   3735
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
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
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6800
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Login"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sign Up"
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
Splash.Show
Else
MsgBox "Login Failed...Please login with correct password"
End If
Else
MsgBox "Login Failed...Please login with correct username"
Frame2.Visible = True
Frame1.Visible = False
End If
End Sub

Private Sub cmdSign_Click()
'Register a new user
    rs.AddNew
'Copy data from form to table

rs("Username") = txtUser
rs("Password") = txtPassword3
If txtPassword2.Text = txtPassword3.Text Then
'save data in the table
rs.Update
MsgBox "Sign Up Successful, Please logon with Username and Password"
Frame1.Visible = True
Frame2.Visible = False
Else
MsgBox "please confirm your password"
End If
End Sub

Private Sub cmdSignUp_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub Form_Load()
Frame1.Visible = True
Frame2.Visible = False
cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=StudentsReg;Data Source=CLASSPC\SQLEXPRESS2019"
rs.Open "StudentLog", cn, adOpenDynamic, adLockOptimistic

End Sub


Private Sub TabStrip1_Click()

If TabStrip1.Tabs(1).Selected = True Then
Frame1.Visible = True
Frame2.Visible = False
Else
Frame1.Visible = False
Frame2.Visible = True
End If
End Sub



