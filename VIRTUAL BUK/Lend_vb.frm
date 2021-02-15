VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0FFC0&
   Caption         =   "LEND BOOK FROM LIBRARY"
   ClientHeight    =   7680
   ClientLeft      =   2925
   ClientTop       =   810
   ClientWidth     =   15300
   LinkTopic       =   "Form3"
   ScaleHeight     =   7680
   ScaleWidth      =   15300
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "Lend_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2280
      TabIndex        =   20
      Top             =   600
      Width           =   2895
   End
   Begin VB.CommandButton Command8 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      TabIndex        =   19
      Top             =   3480
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "NEXT"
      DownPicture     =   "Lend_vb.frx":0000
      BeginProperty Font 
         Name            =   "Broadway"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   13680
      TabIndex        =   18
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "add new student"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   17
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CommandButton Command6 
      Caption         =   "update"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   7440
      TabIndex        =   16
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "Return_date"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9000
      TabIndex        =   14
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "Book_name"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9000
      TabIndex        =   12
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "Book_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   9000
      TabIndex        =   10
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Lent_date"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "Student_name"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2280
      TabIndex        =   6
      Top             =   2040
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "Student_ID"
      DataSource      =   "Adodc1"
      Height          =   735
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NUMBER OF STUDENTS BORROWED     BOOKS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   12240
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Lend_vb.frx":6649D
      Height          =   2415
      Left            =   1800
      TabIndex        =   1
      Top             =   4440
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   4260
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ALL LENT BOOKS DATABASE"
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   12240
      Top             =   1920
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1508
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
      Connect         =   $"Lend_vb.frx":664B2
      OLEDBString     =   $"Lend_vb.frx":6654C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Lent_books"
      Caption         =   "Lent Database Controller"
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
   Begin VB.CommandButton Command1 
      Caption         =   "BACK TO MAIN MENU"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   0
      Top             =   6840
      Width           =   4575
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "LENT ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   21
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "ADD STUDENT TO LENT LIST"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "RETURN DATE:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6720
      TabIndex        =   13
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "BOOK NAME:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6720
      TabIndex        =   11
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "BOOK  ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6720
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "LENT DATE :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   7
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "STUDENT NAMES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   -120
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "STUDENT ID :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form3.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
If Text1.Text <> "" Then
Dim given_id As Double
Dim a As Double
a = MsgBox("We_Are_checking_the_given_ID", "64", "CHECKING..ID")
Adodc1.Recordset.Open
Else
a = MsgBox("First_enter_ID_above", "16", "No_Given_ID")
End If
End Sub

Private Sub Command3_Click()
Dim countt As Integer
countt = Adodc1.Recordset.RecordCount
Dim a As Integer
a = MsgBox("Dou you want to see the number of students who are in lent table", 4 + 32, "CONFIRM TO SEE RECORD COUNT FOR LENT BOOKS")
If a = 6 Then
b = MsgBox(countt, 64, "ALLstudents lent books")
Else
End If
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveNext
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.AddNew
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.Save
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MovePrevious
End Sub
