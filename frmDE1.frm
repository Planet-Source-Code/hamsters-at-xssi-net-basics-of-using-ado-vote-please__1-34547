VERSION 5.00
Begin VB.Form frmDE1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB 6.0 working with ADO"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   2160
      TabIndex        =   7
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdedit 
      Caption         =   "Edit"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   2535
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2640
      Width           =   735
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">>>"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   495
   End
   Begin VB.CommandButton cmdBefore 
      Caption         =   "<<<"
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Author: J. Brandon George // email: josephbg@aol.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "info:"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   1320
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   2520
      Y2              =   2520
   End
End
Attribute VB_Name = "frmDE1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
cmdUpdate.Visible = True

 Text1.Text = ""
 ModADO.AddDB
 Text1.SetFocus

cmdDel.Visible = False
cmdedit.Visible = False
cmdAdd.Visible = False

End Sub

Private Sub cmdBefore_Click()
On Error Resume Next
ModADO.Rs1.MovePrevious
Text1.Text = ModADO.Rs1!info

End Sub

Private Sub cmdDel_Click()

ModADO.DeleteDB "info", Text1.Text
MsgBox "" & Text1.Text & " has been deleted from the database", vbCritical, "Deleted"
ModADO.Rs1.MoveFirst
Text1.Text = ModADO.Rs1!info

End Sub

Private Sub cmdedit_Click()
 cmdAdd.Visible = False
 cmdedit.Visible = False
 cmdDel.Visible = False
   
   cmdUpdate.Visible = True
   
 
End Sub

Private Sub cmdNext_Click()
 On Error Resume Next
 ModADO.Rs1.MoveNext
 Text1.Text = ModADO.Rs1!info
 
End Sub

Private Sub cmdSearch_Click()
frmSearch.Show

End Sub

Private Sub cmdUpdate_Click()
On Error Resume Next
cmdUpdate.Visible = False

ModADO.Rs1!info = Text1.Text
ModADO.UpdateDB

cmdAdd.Visible = True
cmdedit.Visible = True
cmdDel.Visible = True

End Sub



Private Sub Form_Load()
cmdUpdate.Visible = False
ModADO.OpenDB
Text1.Text = ModADO.Rs1!info


End Sub

Private Sub Form_Terminate()
ModADO.CloseDB
End

End Sub
