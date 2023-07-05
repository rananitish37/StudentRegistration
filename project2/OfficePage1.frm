VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form OfficePage1 
   BackColor       =   &H8000000A&
   Caption         =   "Form1"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   12600
      TabIndex        =   46
      Top             =   9360
      Width           =   975
   End
   Begin VB.TextBox txtRegNo 
      Height          =   405
      Left            =   2160
      TabIndex        =   24
      Top             =   720
      Width           =   3375
   End
   Begin VB.ComboBox cboAdmitedTo 
      Height          =   315
      ItemData        =   "OfficePage1.frx":0000
      Left            =   2160
      List            =   "OfficePage1.frx":0016
      TabIndex        =   23
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox txtRoll 
      Height          =   375
      Left            =   2160
      TabIndex        =   22
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   2520
      Width           =   3375
   End
   Begin VB.TextBox txtPreAdd 
      Height          =   975
      Left            =   2160
      TabIndex        =   19
      Top             =   4560
      Width           =   9255
   End
   Begin VB.TextBox txtPerAdd 
      Height          =   855
      Left            =   2160
      TabIndex        =   18
      Top             =   5640
      Width           =   9255
   End
   Begin VB.TextBox txtFather 
      Height          =   405
      Left            =   2160
      TabIndex        =   17
      Top             =   6600
      Width           =   4215
   End
   Begin VB.TextBox txtMob1 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   7080
      Width           =   4215
   End
   Begin VB.TextBox txtMother 
      Height          =   495
      Left            =   2160
      TabIndex        =   15
      Top             =   7680
      Width           =   4215
   End
   Begin VB.TextBox txtMob2 
      Height          =   405
      Left            =   2160
      TabIndex        =   14
      Top             =   8280
      Width           =   4215
   End
   Begin VB.ComboBox cboNationality 
      Height          =   315
      ItemData        =   "OfficePage1.frx":007C
      Left            =   2160
      List            =   "OfficePage1.frx":0083
      TabIndex        =   13
      Top             =   8880
      Width           =   4215
   End
   Begin VB.TextBox txtBPlace 
      Height          =   285
      Left            =   2160
      TabIndex        =   12
      Top             =   9360
      Width           =   4215
   End
   Begin VB.TextBox txtOccup 
      Height          =   375
      Left            =   9120
      TabIndex        =   11
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox txtReligion 
      Height          =   375
      Left            =   9120
      TabIndex        =   10
      Top             =   7320
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   2175
      Left            =   11880
      ScaleHeight     =   2115
      ScaleWidth      =   2115
      TabIndex        =   9
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Image"
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
      Left            =   11880
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.TextBox txtAge 
      Height          =   405
      Left            =   7800
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
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
      Left            =   11640
      TabIndex        =   5
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Search"
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
      Left            =   8160
      TabIndex        =   4
      Top             =   9360
      Width           =   1335
   End
   Begin VB.ComboBox cboGender 
      Height          =   315
      ItemData        =   "OfficePage1.frx":008F
      Left            =   2160
      List            =   "OfficePage1.frx":0099
      TabIndex        =   3
      Text            =   "Select Gender"
      Top             =   3840
      Width           =   1695
   End
   Begin VB.ComboBox cboBranch 
      Height          =   315
      ItemData        =   "OfficePage1.frx":00AB
      Left            =   7800
      List            =   "OfficePage1.frx":00C4
      TabIndex        =   2
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
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
      Left            =   9720
      TabIndex        =   1
      Top             =   9360
      Width           =   1335
   End
   Begin VB.CommandButton cmdClear 
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
      Left            =   11280
      TabIndex        =   0
      Top             =   9360
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc ADODB 
      Height          =   375
      Left            =   4440
      Top             =   9960
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
      Connect         =   $"OfficePage1.frx":0109
      OLEDBString     =   $"OfficePage1.frx":0194
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
   Begin MSComDlg.CommonDialog CDL 
      Left            =   14160
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DPDate 
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120913921
      CurrentDate     =   44286
   End
   Begin MSComCtl2.DTPicker DPDOB 
      Height          =   375
      Left            =   2160
      TabIndex        =   21
      Top             =   3120
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   661
      _Version        =   393216
      Format          =   120913921
      CurrentDate     =   44285
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
      Left            =   0
      TabIndex        =   45
      Top             =   720
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
      Left            =   0
      TabIndex        =   44
      Top             =   1320
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
      Left            =   0
      TabIndex        =   43
      Top             =   2520
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
      Left            =   0
      TabIndex        =   42
      Top             =   3120
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
      Left            =   0
      TabIndex        =   41
      Top             =   3960
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
      Left            =   0
      TabIndex        =   40
      Top             =   4680
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
      Left            =   0
      TabIndex        =   39
      Top             =   5640
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
      Left            =   0
      TabIndex        =   38
      Top             =   6720
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
      Left            =   0
      TabIndex        =   37
      Top             =   7200
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
      Left            =   0
      TabIndex        =   36
      Top             =   7800
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
      Left            =   120
      TabIndex        =   35
      Top             =   8400
      Width           =   1935
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
      Left            =   120
      TabIndex        =   34
      Top             =   8880
      Width           =   1935
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
      Left            =   120
      TabIndex        =   33
      Top             =   9360
      Width           =   1695
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
      Left            =   6000
      TabIndex        =   32
      Top             =   720
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
      Left            =   0
      TabIndex        =   31
      Top             =   1920
      Width           =   1935
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
      Left            =   6000
      TabIndex        =   30
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label17 
      BackColor       =   &H8000000A&
      Caption         =   "Student Image"
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
      Left            =   12120
      TabIndex        =   29
      Top             =   1080
      Width           =   1695
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
      Left            =   6960
      TabIndex        =   28
      Top             =   6600
      Width           =   1815
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
      Left            =   7080
      TabIndex        =   27
      Top             =   7320
      Width           =   1815
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
      Left            =   6000
      TabIndex        =   26
      Top             =   0
      Width           =   3375
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
      Left            =   6000
      TabIndex        =   25
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "OfficePage1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub cmdCheck_Click()
rs.MoveFirst

'find record in the table
rs.Find ("RollNo='" & txtRoll & "'")

'if available,display details on form
If Not rs.EOF Then
cboAdmitedTo = rs("AdmitedTo")
txtRoll = rs("RollNo")
txtName = rs("Name")
cboBranch = rs("Branch")
DPDOB = rs("DOB")
cboGender = rs("Gender")
txtPreAdd = rs("PreAddress")
txtPerAdd = rs("PerAddress")
txtFather = rs("FatherName")
txtMob1 = rs("MobileNo1")
txtMother = rs("MotherName")
txtMob2 = rs("MobileNo2")
cboNationality = rs("Nationality")
txtBPlace = rs("BirthPlace")
DPDate = rs("Date")
txtAge = rs("Age")
txtOccup = rs("Occupation")
txtReligion = rs("Religion")
Picture1 = rs("Photo")
End Sub

Private Sub cmdClear_Click()
cboAdmitedTo.ListIndex = -1
txtRoll = ""
txtName = ""
cboBranch.ListIndex = -1
DPDOB = ""
cboGender = -1
txtPreAdd = ""
txtPerAdd = ""
txtFather = ""
txtMob1 = ""
txtMother = ""
txtMob2 = ""
cboNationality.ListIndex = -1
txtBPlace = ""
DPDate = ""
txtAge = ""
txtOccup = ""
txtReligion = ""
Picture1 = ""
End Sub

Private Sub cmdDelete_Click()
rs.Delete
MsgBox "Record deleted"
End Sub

Private Sub Form_Load()
cn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=StudentsReg;Data Source=CLASSPC\SQLEXPRESS2019"
rs.Open "StudentReg", cn, adOpenDynamic, adLockOptimistic
End Sub
