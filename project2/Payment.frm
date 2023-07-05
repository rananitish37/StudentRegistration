VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Payment 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   ScaleHeight     =   10395
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc ADODB 
      Height          =   375
      Left            =   2280
      Top             =   9600
      Width           =   1455
      _ExtentX        =   2566
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
      Connect         =   $"Payment.frx":0000
      OLEDBString     =   $"Payment.frx":008B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Account"
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
   Begin VB.Frame payment 
      Height          =   9135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CommandButton cmdCancel 
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
         Height          =   375
         Left            =   5640
         TabIndex        =   28
         Top             =   8640
         Width           =   1455
      End
      Begin VB.CommandButton cmdDone 
         Caption         =   "Done"
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
         Left            =   3840
         TabIndex        =   27
         Top             =   8640
         Width           =   1575
      End
      Begin VB.CommandButton cndTotal 
         Caption         =   "Total"
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
         Left            =   120
         TabIndex        =   26
         Top             =   7560
         Width           =   1695
      End
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   2280
         TabIndex        =   25
         Top             =   7560
         Width           =   3375
      End
      Begin VB.TextBox txtLate 
         Height          =   285
         Left            =   2280
         TabIndex        =   24
         Top             =   7080
         Width           =   3375
      End
      Begin VB.TextBox txtAdmission 
         Height          =   285
         Left            =   2280
         TabIndex        =   23
         Top             =   6600
         Width           =   3375
      End
      Begin VB.TextBox txtHostel 
         Height          =   285
         Left            =   2280
         TabIndex        =   22
         Top             =   6120
         Width           =   3375
      End
      Begin VB.TextBox txtLibrary 
         Height          =   285
         Left            =   2280
         TabIndex        =   21
         Top             =   5760
         Width           =   3375
      End
      Begin VB.TextBox txtIntitute 
         Height          =   285
         Left            =   2280
         TabIndex        =   20
         Top             =   5280
         Width           =   3375
      End
      Begin VB.TextBox txtDevelop 
         Height          =   285
         Left            =   2280
         TabIndex        =   19
         Top             =   4680
         Width           =   3375
      End
      Begin VB.TextBox txtTuition 
         Height          =   285
         Left            =   2280
         TabIndex        =   18
         Top             =   4080
         Width           =   3375
      End
      Begin VB.TextBox txtRoll 
         Height          =   285
         Left            =   2280
         TabIndex        =   10
         Top             =   3480
         Width           =   3375
      End
      Begin VB.TextBox txtDDNo 
         Height          =   285
         Left            =   2280
         TabIndex        =   8
         Top             =   8040
         Width           =   3375
      End
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FF8080&
         Caption         =   "OK"
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
         Left            =   2280
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin VB.ComboBox cboSemester 
         Height          =   315
         ItemData        =   "Payment.frx":0116
         Left            =   2280
         List            =   "Payment.frx":012C
         TabIndex        =   5
         Text            =   "Select Semester"
         Top             =   1560
         Width           =   3375
      End
      Begin VB.ComboBox cboCourse 
         Height          =   315
         ItemData        =   "Payment.frx":0142
         Left            =   2280
         List            =   "Payment.frx":014C
         TabIndex        =   4
         Text            =   "Select Course"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label12 
         Caption         =   "Late Fine:-"
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
         Left            =   120
         TabIndex        =   17
         Top             =   7080
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Addmission Fee:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   6600
         Width           =   2055
      End
      Begin VB.Label Label10 
         Caption         =   "Hostel Fee:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   6240
         Width           =   1455
      End
      Begin VB.Label Label9 
         Caption         =   "Library Fee:-"
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
         Left            =   120
         TabIndex        =   14
         Top             =   5760
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Intitute Exam. Fee:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   5280
         Width           =   2175
      End
      Begin VB.Label Label7 
         Caption         =   "Development Fee:-"
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
         Left            =   120
         TabIndex        =   12
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label6 
         Caption         =   "Tuition Fee:-"
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
         Left            =   120
         TabIndex        =   11
         Top             =   4080
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "RollNo:-"
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
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label btnTotal 
         Caption         =   "DD No:-"
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
         Left            =   120
         TabIndex        =   7
         Top             =   8160
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Semester:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Course:-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C000C0&
         Caption         =   "Student Fee Paymenet Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7215
      End
   End
End
Attribute VB_Name = "Payment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDone_Click()
rs.AddNew
rs("RollNo") = txtRoll
rs("DDNo") = txtDDNo
rs.Update
Registration.Show
End Sub

Private Sub cmdOK_Click()
If cboSemester.Text = 1 Then
txtTuition.Text = "25300"
txtDevelop.Text = "1700"
txtIntitute.Text = "1500"
txtHostel.Text = "3500"
txtAdmission.Text = "3000"
txtLate.Text = "0"
txtLibrary.Text = "0"
Else
If cboSemester.Text = 2 Then
txtTuition.Text = "25300"
txtDevelop.Text = "1700"
txtIntitute.Text = "1500"
txtHostel.Text = "3500"
txtAdmission.Text = "0"
txtLate.Text = "0"
txtLibrary.Text = "0"
Else
If cboSemester.Text = 3 Then
txtTuition.Text = "26500"
txtDevelop.Text = "1800"
txtIntitute.Text = "1600"
txtHostel.Text = "3700"
txtAdmission.Text = "0"
txtLate.Text = "0"
txtLibrary.Text = "0"
Else
If cboSemester.Text = 4 Then
txtTuition.Text = "26500"
txtDevelop.Text = "1800"
txtIntitute.Text = "1600"
txtHostel.Text = "3700"
txtAdmission.Text = "0"
txtLate.Text = "0"
txtLibrary.Text = "0"
Else
If cboSemester.Text = 5 Then
txtTuition.Text = "25300"
txtDevelop.Text = "1700"
txtIntitute.Text = "1500"
txtHostel.Text = "3500"
txtAdmission.Text = "0"
txtLate.Text = "0"
txtLibrary.Text = "0"
Else
txtTuition.Text = "27800"
txtDevelop.Text = "1900"
txtIntitute.Text = "1700"
txtHostel.Text = "0"
txtAdmission.Text = "0"
txtLate.Text = "0"
txtLibrary.Text = "0"
End If
End If
End If
End If
End If
End Sub

Private Sub cndTotal_Click()
Dim t As Double
 t = Val(txtTuition.Text + txtDevelop.Text + txtIntitute.Text + txtLibrary.Text + txtHostel.Text + txtAdmission.Text + txtLate.Text)
 txtTotal.Text = Val(t)
End Sub

Private Sub Form_Load()
Me.Left = (Screen.Width - Me.Width) / 4
Me.Top = (Screen.Height - Me.Height) / 4
cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=StudentsReg;Data Source=CLASSPC\SQLEXPRESS2019"
rs.Open "Account", cn, adOpenDynamic, adLockOptimistic
End Sub
