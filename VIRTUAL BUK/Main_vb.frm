VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "          VIRTUAL  BUK  (Library Management Software)"
   ClientHeight    =   7830
   ClientLeft      =   2745
   ClientTop       =   465
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   15240
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4200
      TabIndex        =   7
      Top             =   7200
      Width           =   6615
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SETTING"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   5
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "BORROW"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12000
      TabIndex        =   4
      Top             =   2880
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LEND"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12000
      TabIndex        =   3
      Top             =   4320
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "REGIST"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   12000
      TabIndex        =   2
      Top             =   5760
      Width           =   3255
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HOME"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      TabIndex        =   1
      Top             =   2880
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackColor       =   &H0000C000&
      Caption         =   "VIRTUAL BUK (Library_Management_System)"
      BeginProperty Font 
         Name            =   "Modern No. 20"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click()
Form1.Hide
Form1.Show
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Form1.Hide
Form2.Show
End Sub

Private Sub Command3_Click()
Form1.Hide
Form3.Show
End Sub

Private Sub Command4_Click()
Form1.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
Form1.Hide
Form6.Show
End Sub

Private Sub Command6_Click()
Form1.Hide
Form5.Show
End Sub

