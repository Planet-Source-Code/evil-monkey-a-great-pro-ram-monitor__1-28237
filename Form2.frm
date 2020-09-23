VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAM Log"
   ClientHeight    =   3570
   ClientLeft      =   6735
   ClientTop       =   4470
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4560
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "60000"
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Disable Logging"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   3240
      Top             =   1560
   End
   Begin VB.ListBox List1 
      Height          =   2985
      ItemData        =   "Form2.frx":0000
      Left            =   0
      List            =   "Form2.frx":0002
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin VB.Label Label1 
      Caption         =   "How often RAM gets logged (In Milliseconds)"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
'You can figure it out
If Check1.Value = Unchecked Then
Text1.Enabled = False
Timer1.Interval = Text1.Text
Timer1.Enabled = True
Else
Timer1.Enabled = False
Text1.Enabled = True
End If
End Sub

Private Sub Form_Load()
'Hmmm...  I don't know. Do you?

End Sub

Private Sub Timer1_Timer()
'This adds a log item and adjusts the HBar to fit your size of your
'biggest listbox entry
List1.AddItem Date & " - " & Time & "   -   " & Form1.Label6.Caption & " / " & Form1.Label2.Caption
Call AddHScroll(List1)
End Sub
