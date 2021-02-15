VERSION 5.00
Begin VB.Form Form5 
   BackColor       =   &H00C0FFC0&
   Caption         =   "CHANGE SETTINGS OF LIBRARY"
   ClientHeight    =   8010
   ClientLeft      =   3090
   ClientTop       =   300
   ClientWidth     =   15690
   LinkTopic       =   "Form5"
   ScaleHeight     =   8010
   ScaleWidth      =   15690
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
      Left            =   5520
      TabIndex        =   0
      Top             =   7080
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "THE SETTINGS OF THIS SOFTWARE IS NOT AVAILABLE IN THIS VERSION OF YOUR COMPUTER.    TRY TO VISIT Datchtheprogrammer.blogspot.com"
      BeginProperty Font 
         Name            =   "20th Century Font"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   4200
      TabIndex        =   2
      Top             =   2520
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "SETTINGS"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15615
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form5.Hide
Form1.Show
End Sub

