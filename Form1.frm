VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RAM Monitor"
   ClientHeight    =   2520
   ClientLeft      =   4920
   ClientTop       =   3945
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4470
   Begin VB.Frame Frame2 
      Caption         =   "Amount Of RAM In Use"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   4215
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   135
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   500
      Left            =   3960
      Top             =   1560
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3600
      Top             =   1920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3600
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3960
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      Caption         =   "RAM Information"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   4215
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   3120
         Top             =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label6 
         Caption         =   "Label6"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   3975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1320
         Width           =   3975
      End
   End
   Begin VB.Menu mnuTrayStuff 
      Caption         =   "ass"
      Visible         =   0   'False
      Begin VB.Menu mnuTotal 
         Caption         =   "assd"
      End
      Begin VB.Menu mnuUsed 
         Caption         =   "fucka!"
      End
      Begin VB.Menu mnuAvail 
         Caption         =   "shitty"
      End
      Begin VB.Menu hyphonThing 
         Caption         =   "-"
      End
      Begin VB.Menu mnuShow 
         Caption         =   "&Show RAM Monitor"
      End
      Begin VB.Menu mnuloggerSYS 
         Caption         =   "Show &RAM Log"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuRAMlog 
         Caption         =   "Show &RAM Log"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuTray 
         Caption         =   "&Add to system tray"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuMBorKB 
         Caption         =   "&Show RAM in Mega Bytes"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuWarn 
         Caption         =   "&Warn me if my RAM gets to a dangerous level"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^E
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type MEMORYSTATUS
    dwLength        As Long
    dwMemoryLoad    As Long
    dwTotalPhys     As Long
    dwAvailPhys     As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual  As Long
    dwAvailVirtual  As Long
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)

Private Sub Form_Load()
'This is for setting labels 1,2, and 6 to what they
'should be so we can log our first RAM log
Dim MS As MEMORYSTATUS

On Local Error Resume Next

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
    Label1.Caption = "Total RAM: " & Format$(.dwTotalPhys / 1024, "#,###") & " KB"
    Label6.Caption = "Used RAM: " & Format$(Format$(.dwTotalPhys / 1024, "#,###") - Format$(.dwAvailPhys / 1024, "#,###"), "#,###") & " KB"
    If Format$(.dwAvailPhys / 1024, "#,###") = 0 Then
    Label2.Caption = "None"
    Exit Sub
    End If
    Label2.Caption = "Avaliable RAM: " & Format$(.dwAvailPhys / 1024, "#,###") & " KB"

End With
'This is just common stuff like checking the option to warn you
'and add it to the tray(I like it in the tray because I put it in
'my startup folder)
mnuWarn_Click
AddToTray Me.Icon, Me.Caption, Me
Form2.List1.AddItem "Program Loaded   -   " & Date & " - " & Time & "   -   " & Form1.Label6.Caption & " / " & Form1.Label2.Caption
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'This is for the systemtray menu
If RespondToTray(X) = 1 Then
ShowFormAgain Me
End If
If RespondToTray(X) = 2 Then
PopupMenu mnuTrayStuff
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Hmmm...  What could this do?
End
End Sub

Private Sub mnuExit_Click()
'Hmm...
End
End Sub

Private Sub mnuloggerSYS_Click()
'This is to check the appropreate things and to show the logger form
If mnuRAMlog.Checked = True Then
mnuRAMlog.Checked = False
mnuloggerSYS.Checked = False
Form2.Hide
Else
mnuRAMlog.Checked = True
mnuloggerSYS.Checked = True
Form2.Show
End If
Call AddHScroll(Form2.List1)
End Sub

Private Sub mnuMBorKB_Click()
'This just unchecks and checks the menus and changes
'the RAM from KB view to MB view
If mnuMBorKB.Checked = True Then
mnuMBorKB.Checked = False
Timer2.Enabled = False
Timer1.Enabled = True
Else
mnuMBorKB.Checked = True
Timer2.Enabled = True
Timer1.Enabled = False

End If
End Sub

Private Sub mnuRAMlog_Click()
'This does the same as mnuloggerSYS eccept its for the form
'and not the systray
If mnuRAMlog.Checked = True Then
mnuRAMlog.Checked = False
mnuloggerSYS.Checked = False
Form2.Hide
Else
mnuRAMlog.Checked = True
mnuloggerSYS.Checked = True
Form2.Show
End If
Call AddHScroll(Form2.List1)
End Sub

Private Sub mnuShow_Click()
'No comment
ShowFormAgain Form1
End Sub

Private Sub mnuTray_Click()
'This adds the form to the tray
AddToTray Me.Icon, Me.Caption, Me
End Sub

Private Sub mnuTrayExit_Click()
'This gets rid of the tray icon and exits the program
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub mnuWarn_Click()
'This enables a timer that warns you if your RAM is low
If mnuWarn.Checked = True Then
mnuWarn.Checked = False
Timer3.Enabled = False
Timer5.Enabled = False
Else
mnuWarn.Checked = True
Timer3.Enabled = True
End If
End Sub

Private Sub Timer1_Timer()
'This updates all the RAM labels
Dim MS As MEMORYSTATUS

On Local Error Resume Next

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
    Label1.Caption = "Total RAM: " & Format$(.dwTotalPhys / 1024, "#,###") & " KB"
    Label6.Caption = "Used RAM: " & Format$(Format$(.dwTotalPhys / 1024, "#,###") - Format$(.dwAvailPhys / 1024, "#,###"), "#,###") & " KB"
    If Format$(.dwAvailPhys / 1024, "#,###") = 0 Then
    Label2.Caption = "None"
    Exit Sub
    End If
    Label2.Caption = "Avaliable RAM: " & Format$(.dwAvailPhys / 1024, "#,###") & " KB"

End With
End Sub

Private Sub Timer2_Timer()
'This does the same as above eccept it shows it in MB
Dim MS As MEMORYSTATUS

On Local Error Resume Next

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
    Label1.Caption = "Total RAM: " & Format$(.dwTotalPhys / 1024, "###") / 1000 & " MB"
    Label6.Caption = "Used RAM: " & Format$(Format$(.dwTotalPhys / 1024, "###") - Format$(.dwAvailPhys / 1024, "###"), "###") / 1000 & " MB"
    If Format$(.dwAvailPhys / 1024, "###") = 0 Then
    Label2.Caption = "None"
    Exit Sub
    End If
    Label2.Caption = "Avaliable RAM: " & Format$(.dwAvailPhys / 1024, "###") / 1000 & " MB"

End With
End Sub

Private Sub Timer3_Timer()
'This is the timer that warns you if you are low on RAM
Dim MS As MEMORYSTATUS

On Local Error Resume Next

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
If Format$(.dwAvailPhys / 1024, "#,###") <= 1000 Then
MsgBox "Your RAM is down to 1 MB or less.  You should restart or exit some programs.  If your RAM is still less than 1 MB in 1 minute, this will come up again.  To get rid of this message uncheck the option to warn you when RAM reaches a dangerous level.", vbOKOnly Or vbExclamation, "RAM Low!"
Timer3.Enabled = False
Timer5.Enabled = True
End If
End With
End Sub

Private Sub Timer4_Timer()
'This controls the proggress bar
Dim MS As MEMORYSTATUS

On Local Error Resume Next

MS.dwLength = Len(MS)
Call GlobalMemoryStatus(MS)
With MS
ProgressBar1.Max = Format$(.dwTotalPhys / 1024, "#,###")
ProgressBar1.Value = Format$(Format$(.dwTotalPhys / 1024, "#,###") - Format$(.dwAvailPhys / 1024, "#,###"), "#,###")
End With
mnuTotal.Caption = Label1.Caption
mnuAvail.Caption = Label2.Caption
mnuUsed.Caption = Label6.Caption
End Sub

Private Sub Timer5_Timer()
'After you click OK on the RAM warning thing this is enabled and
'counts to 60, then it disables itself and enables the warning timer
Label3.Caption = Label3.Caption + 1
If Label3.Caption = 60 Then
Label3.Caption = 0
Timer3.Enabled = True
Timer5.Enabled = False
End If
End Sub
