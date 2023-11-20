VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "STUDENT INFORMATION SYSTEM "
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12990
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12990
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "display record"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   32
      Top             =   6360
      Width           =   2655
   End
   Begin VB.TextBox result 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   28
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox totalmark 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   26
      Top             =   1920
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   10320
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1296
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=studentinfo;Data Source=LAPTOP-BT8CH199"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=studentinfo;Data Source=LAPTOP-BT8CH199"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "st_table"
      Caption         =   "Adodc1"
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6600
      TabIndex        =   25
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton FIND 
      Caption         =   "FIND"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6720
      TabIndex        =   24
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton EXIT 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   23
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton DELETE 
      Caption         =   "DETELE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   22
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton CLEAR 
      Caption         =   "CLEAR"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7800
      TabIndex        =   21
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton ADD 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   20
      Top             =   5400
      Width           =   1215
   End
   Begin VB.TextBox grade 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   19
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox percentage 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox webmark 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox rdbmsmark 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox javamark 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   15
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox st_class 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   14
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox st_dob 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox st_age 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox st_name 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   11
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox st_no 
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9
         Charset         =   0
         Weight          =   300
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "Subject - Mark"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   31
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label Record 
      Caption         =   "Record"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   375
      Left            =   5040
      TabIndex        =   30
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "result"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   29
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label total 
      Alignment       =   2  'Center
      Caption         =   "total"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   27
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "GRADE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      Caption         =   "PERCENTAGE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   8
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "WEB"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "RDBMS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "JAVA"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "STUDENT CLASS"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   4
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "STUDENT NUMBER"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "STUDENT DOB"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   2
      Top             =   3600
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "STUDENT AGE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "STUDENT NAME"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub ADD_Click()
Dim sno As Integer
Dim sname As String
Dim age As String
Dim dob As String
Dim class As String
Dim sub_java As Integer
Dim sub_rdbms As Integer
Dim sub_web As Integer
Dim avg As Integer
Dim gr_total As Integer
Dim grade1 As String
Dim result1 As String


rs.AddNew

sno = st_no.Text
rs.Fields(0).Value = sno

sname = st_name.Text
rs.Fields(1).Value = sname

age = st_age.Text
rs.Fields(2).Value = age

dob = st_dob.Text
rs.Fields(3).Value = dob

class = st_class.Text
rs.Fields(4).Value = class

sub_java = Val(javamark.Text)
rs.Fields(5).Value = sub_java

sub_rdbms = Val(rdbmsmark.Text)
rs.Fields(6).Value = sub_rdbms

sub_web = Val(webmark.Text)
rs.Fields(7).Value = sub_web

gr_total = sub_java + sub_rdbms + sub_web
totalmark.Text = gr_total
rs.Fields(10).Value = gr_total

avg = gr_total / 3
rs.Fields(8).Value = avg
percentage.Text = avg

If (avg > 90) Then
    grade1 = "A+"

ElseIf (avg <= 90 And avg > 80) Then
    grade1 = "A"

ElseIf (avg <= 80 And avg > 70) Then
    grade1 = "B"
    
ElseIf (avg <= 70 And avg > 60) Then
    grade1 = "c"

ElseIf (avg <= 60 And avg > 50) Then
    grade1 = "D"

ElseIf (avg <= 50) Then
    grade1 = "F"

End If

rs.Fields(9).Value = grade1
 grade.Text = grade1
 
 If Val(javamark.Text) >= 50 And Val(rdbmsmark.Text) >= 50 And Val(webmark.Text) >= 50 Then
 result1 = "PASS"

Else
  result1 = " fail"
  
End If

rs.Fields(11).Value = result1
result.Text = result1

rs.Update
MsgBox ("RECORD INSERTED SUCCESSFULLY")

End Sub

Private Sub CLEAR_Click()
 st_no.Text = " "
 st_name.Text = " "
 st_age.Text = " "
 st_dob.Text = " "
 st_class.Text = " "
 javamark.Text = " "
 rdbmsmark.Text = " "
 webmark.Text = " "
 totalmark.Text = " "
 percentage.Text = " "
 result.Text = " "
 grade.Text = " "
End Sub

Private Sub Combo1_Click()
 Dim rd As Integer
rs.MoveFirst
rd = Val(Combo1.Text)
Do While Not rs.EOF
If Val(rs(0)) = rd Then
st_no.Text = rs.Fields(0).Value
st_name.Text = rs.Fields(1).Value
st_age.Text = rs.Fields(2).Value
st_dob.Text = rs.Fields(3).Value
st_class = rs.Fields(4).Value
javamark.Text = rs.Fields(5).Value
rdbmsmark.Text = rs.Fields(6).Value
webmark.Text = rs.Fields(7).Value

percentage.Text = rs.Fields(8).Value

grade.Text = rs.Fields(9).Value
totalmark.Text = rs.Fields(10).Value
result.Text = rs.Fields(11).Value


MsgBox "record already exist"
End If

rs.MoveNext
Loop
Combo1.Text = " "
End Sub

Private Sub Command1_Click()
DataReport1.Show

End Sub

Private Sub DELETE_Click()
 Dim rd2 As Integer
rd2 = Val(st_no.Text)
rs.MoveFirst

Do While Not (rs.EOF)
If rs.Fields(0).Value = rd2 Then
rs.DELETE
MsgBox ("record deleted from the database")
Exit Do

Else
rs.MoveNext
End If
Loop

Call CLEAR_Click

End Sub

Private Sub EXIT_Click()

End

End Sub

Private Sub FIND_Click()
rs.MoveFirst
Do While (Not rs.EOF)
Combo1.AddItem (rs(0))
rs.MoveNext
Loop
MsgBox ("click the combobox to see the records")
End Sub

Private Sub Form_Load()
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.Open ("Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=studentinfo;Data Source=LAPTOP-BT8CH199")
rs.Open "select * from student_table", db, adOpenDynamic, adLockOptimistic

End Sub


