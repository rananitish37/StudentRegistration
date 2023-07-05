VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Registration 
   BackColor       =   &H8000000A&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14985
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   14985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAmount 
      Height          =   375
      Left            =   11520
      TabIndex        =   44
      Top             =   8280
      Width           =   1935
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Permanent"
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
      Left            =   12720
      TabIndex        =   43
      Top             =   9480
      Width           =   1695
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
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
      Index           =   0
      Left            =   11520
      TabIndex        =   42
      Top             =   9480
      Width           =   1095
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
      Height          =   495
      Left            =   9960
      TabIndex        =   41
      Top             =   9480
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc ADODB 
      Height          =   375
      Left            =   4800
      Top             =   10920
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Connect         =   $"Registration.frx":0000
      OLEDBString     =   $"Registration.frx":008B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "StudentReg"
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
   Begin VB.ComboBox cboBranch 
      Height          =   315
      ItemData        =   "Registration.frx":0116
      Left            =   8040
      List            =   "Registration.frx":012F
      TabIndex        =   40
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox cboGender 
      Height          =   315
      ItemData        =   "Registration.frx":0174
      Left            =   2400
      List            =   "Registration.frx":017E
      TabIndex        =   38
      Text            =   "Select Gender"
      Top             =   3960
      Width           =   1695
   End
   Begin VB.CommandButton cmdPay 
      Caption         =   "Pay"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   37
      Top             =   9480
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000B&
      Caption         =   "Same as present"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11880
      TabIndex        =   36
      Top             =   5880
      Width           =   2295
   End
   Begin VB.TextBox txtAge 
      Height          =   405
      Left            =   8040
      TabIndex        =   35
      Top             =   1440
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker DPDate 
      Height          =   375
      Left            =   8040
      TabIndex        =   34
      Top             =   720
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120586241
      CurrentDate     =   44286
   End
   Begin VB.TextBox txtReligion 
      Height          =   375
      Left            =   9360
      TabIndex        =   33
      Top             =   7440
      Width           =   2655
   End
   Begin VB.TextBox txtOccup 
      Height          =   375
      Left            =   9360
      TabIndex        =   32
      Top             =   6720
      Width           =   2655
   End
   Begin VB.TextBox txtBPlace 
      Height          =   285
      Left            =   2400
      TabIndex        =   31
      Top             =   9480
      Width           =   4215
   End
   Begin VB.ComboBox cboNationality 
      Height          =   315
      ItemData        =   "Registration.frx":0190
      Left            =   2400
      List            =   "Registration.frx":0197
      TabIndex        =   30
      Top             =   9000
      Width           =   4215
   End
   Begin VB.TextBox txtMob2 
      Height          =   405
      Left            =   2400
      TabIndex        =   29
      Top             =   8400
      Width           =   4215
   End
   Begin VB.TextBox txtMother 
      Height          =   495
      Left            =   2400
      TabIndex        =   28
      Top             =   7800
      Width           =   4215
   End
   Begin VB.TextBox txtMob1 
      Height          =   375
      Left            =   2400
      TabIndex        =   27
      Top             =   7200
      Width           =   4215
   End
   Begin VB.TextBox txtFather 
      Height          =   405
      Left            =   2400
      TabIndex        =   26
      Top             =   6720
      Width           =   4215
   End
   Begin VB.TextBox txtPerAdd 
      Height          =   855
      Left            =   2400
      TabIndex        =   25
      Top             =   5760
      Width           =   9255
   End
   Begin VB.TextBox hi 
      Height          =   975
      Left            =   2400
      TabIndex        =   24
      Top             =   4680
      Width           =   9255
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2400
      TabIndex        =   23
      Top             =   2640
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker DPDOB 
      Height          =   375
      Left            =   2400
      TabIndex        =   22
      Top             =   3240
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120586241
      CurrentDate     =   44285
   End
   Begin VB.TextBox txtRoll 
      Height          =   375
      Left            =   2400
      TabIndex        =   21
      Top             =   2040
      Width           =   3375
   End
   Begin VB.ComboBox cboAdmitedTo 
      Height          =   315
      ItemData        =   "Registration.frx":01A3
      Left            =   2400
      List            =   "Registration.frx":01B9
      TabIndex        =   20
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox txtRegNo 
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000A&
      Caption         =   "PaindAmount:-"
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
      Left            =   9600
      TabIndex        =   45
      Top             =   8280
      Width           =   1695
   End
   Begin VB.Label Label21 
      BackColor       =   &H8000000A&
      Caption         =   "Branch"
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
      Left            =   6240
      TabIndex        =   39
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label20 
      BackColor       =   &H8000000A&
      Caption         =   "REGISTRATION FORM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   6240
      TabIndex        =   19
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label19 
      BackColor       =   &H8000000A&
      Caption         =   "Religion"
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
      Left            =   7320
      TabIndex        =   18
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label Label18 
      BackColor       =   &H8000000A&
      Caption         =   "Occupation"
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
      Left            =   7200
      TabIndex        =   17
      Top             =   6720
      Width           =   1815
   End
   Begin VB.Label Label16 
      BackColor       =   &H8000000A&
      Caption         =   "Age"
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
      Left            =   6240
      TabIndex        =   16
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label15 
      BackColor       =   &H8000000A&
      Caption         =   "Roll No"
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
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label14 
      BackColor       =   &H8000000A&
      Caption         =   "Date "
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
      Left            =   6240
      TabIndex        =   14
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackColor       =   &H8000000A&
      Caption         =   "Birth Place"
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
      Left            =   360
      TabIndex        =   13
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Label Label12 
      BackColor       =   &H8000000A&
      Caption         =   "Nationality"
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
      Left            =   360
      TabIndex        =   12
      Top             =   9000
      Width           =   1935
   End
   Begin VB.Label Label11 
      BackColor       =   &H8000000A&
      Caption         =   "Mobile"
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
      Left            =   360
      TabIndex        =   11
      Top             =   8520
      Width           =   1935
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000A&
      Caption         =   "Mother's Name "
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
      TabIndex        =   10
      Top             =   7920
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000A&
      Caption         =   "Mobile"
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
      Left            =   240
      TabIndex        =   9
      Top             =   7320
      Width           =   1935
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000A&
      Caption         =   "Father's Name"
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
      Left            =   240
      TabIndex        =   8
      Top             =   6840
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000A&
      Caption         =   "Permanent Add."
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
      Left            =   240
      TabIndex        =   7
      Top             =   5760
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000A&
      Caption         =   "Present Add."
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
      TabIndex        =   6
      Top             =   4800
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H8000000A&
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000A&
      Caption         =   "Date Of Birth"
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
      Left            =   240
      TabIndex        =   4
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000A&
      Caption         =   "Name"
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
      TabIndex        =   2
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "Admited To"
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
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000A&
      Caption         =   "Registration No-"
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
      TabIndex        =   0
      Top             =   840
      Width           =   1935
   End
End
Attribute VB_Name = "Registration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Clear_Click(Index As Integer)
cboAdmitedTo.ListIndex = -1
txtRoll = ""
txtName = ""
cboBranch = -1
cboGender.ListIndex = -1
txtPreAdd = ""
txtPerAdd = ""
txtFather = ""
txtMob1 = ""
txtMother = ""
txtMob2 = ""
cboNationality = -1
txtBPlace = ""
txtAge = ""
txtOccup = ""
txtReligion = ""
End Sub

Private Sub cmdDone_Click()
rs.AddNew

rs("AdmitedTo") = cboAdmitedTo
rs("RollNo") = txtRoll
rs("Name") = txtName
rs("Branch") = cboBranch
rs("DOB") = DPDOB
rs("Gender") = cboGender
rs("preAddress") = txtPreAdd
rs("perAddress") = txtPerAdd
rs("FatherName") = txtFather
rs("MobileNo1") = txtMob1
rs("MotherName") = txtMother
rs("MobileNo2") = txtMob2
rs("Nationality") = cboNationality
rs("BirthPlace") = txtBPlace
rs("Date") = DPDate
rs("Age") = txtAge
rs("Occupation") = txtOccup
rs("Religion") = txtReligion
rs("PaidAmount") = txtAmount
rs.Update
MsgBox "registered"
End Sub

Private Sub cmdPay_Click()
payment.Show
Registration.Hide
End Sub


Private Sub Command2_Click()
txtPerAdd.Text = txtPreAdd
End Sub

Private Sub Form_Load()
cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=StudentsReg;Data Source=CLASSPC\SQLEXPRESS2019"
rs.Open "StudentReg", cn, adOpenDynamic, adLockOptimistic
End Sub
