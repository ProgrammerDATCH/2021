VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0FFC0&
   Caption         =   "MORE HELP ON SOFTWARE"
   ClientHeight    =   7830
   ClientLeft      =   3090
   ClientTop       =   465
   ClientWidth     =   15480
   LinkTopic       =   "Form6"
   ScaleHeight     =   7830
   ScaleWidth      =   15480
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
      Left            =   4200
      TabIndex        =   0
      Top             =   6960
      Width           =   4575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "THIS SOFTWARE WAS DEVELOPED BY TUYISHIME DAVID +250781733332 DATCHDATCH2001@GMAIL.COM        IN GS RANGO"
      BeginProperty Font 
         Name            =   "20th Century Font"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   3600
      TabIndex        =   1
      Top             =   2280
      Width           =   6615
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form6.Hide
Form1.Show
End Sub
