VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Auto ShutDown"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   3795
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox txtSS 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "frmMain.frx":0000
      Left            =   1560
      List            =   "frmMain.frx":0002
      TabIndex        =   20
      Top             =   720
      Width           =   615
   End
   Begin VB.ListBox txtMM 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "frmMain.frx":0004
      Left            =   840
      List            =   "frmMain.frx":0006
      TabIndex        =   19
      Top             =   720
      Width           =   615
   End
   Begin VB.ListBox txtHH 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      ItemData        =   "frmMain.frx":0008
      Left            =   120
      List            =   "frmMain.frx":000A
      TabIndex        =   18
      Top             =   720
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3360
      Top             =   0
   End
   Begin VB.CommandButton cmdUnSet 
      Caption         =   "Un Set"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdSET 
      Caption         =   "Set"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   16
      Top             =   960
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "¿What would you like to do?"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   3495
      Begin VB.OptionButton optATX 
         Caption         =   "Power Off (ATX)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   1935
      End
      Begin VB.OptionButton optSD 
         Caption         =   "Shut Down"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   14
         Top             =   240
         Width           =   1335
      End
      Begin VB.OptionButton optRestart 
         Caption         =   "Restart"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optLogOff 
         Caption         =   "Logoff"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.ListBox lstAMPM 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2280
      TabIndex        =   10
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "&Hide"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   2
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ss"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1605
      TabIndex        =   9
      Top             =   1080
      Width           =   210
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1480
      TabIndex        =   8
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "mm"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "hh"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   315
      TabIndex        =   6
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ":"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   760
      TabIndex        =   5
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Set the Alarm Time:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1845
   End
   Begin VB.Label lblCurTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Time:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   210
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   1365
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public AMPMSel As Boolean 'Variable to store True/False for Selection of AM/PM.
Public optSel As Boolean  'Variable to store wether Option is selected ot not.

Private Sub cmdExit_Click()
'Remove System Tray Icon
Shell_NotifyIcon NIM_DELETE, nid
'Exit the Program
End
End Sub

Private Sub cmdHide_Click()
'Add to System Tray
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Me.Hide
End Sub

Private Sub cmdSET_Click()
'Check AMPM is slected or not?
If AMPMSel = False Then
    MsgBox "Please select AM/PM from the list."
    Exit Sub
End If
'Check Option is selected or not?
If optSel = False Then
    MsgBox "Please select Option."
    Exit Sub
End If
'Show Message
MsgBox "Alarm is set.. click ok to minimize it.."
'Disable Set Button
cmdSET.Enabled = False
'Enable UnSet Button
cmdUnSet.Enabled = True
'Enable the Checking of Time
Timer2.Enabled = True
'Disable all text boxes
txtHH.Enabled = False
txtMM.Enabled = False
txtSS.Enabled = False
lstAMPM.Enabled = False
'Disable all Options
optATX.Enabled = False
optSD.Enabled = False
optRestart.Enabled = False
optLogOff.Enabled = False
'Add to system tray
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Me.Hide
End Sub

Private Sub cmdUnSet_Click()
'Disable UnSet Button
cmdUnSet.Enabled = False
'Disable Checking of Time
Timer2.Enabled = True
'Enable Set Button
cmdSET.Enabled = True
'Enable All Text Fields
txtHH.Enabled = True
txtMM.Enabled = True
txtSS.Enabled = True
lstAMPM.Enabled = True
'Enable All Options
optATX.Enabled = True
optSD.Enabled = True
optLogOff.Enabled = True
optRestart.Enabled = True
End Sub

Private Sub Form_Load()
'Add AM and PM to List Box
lstAMPM.AddItem "AM"
lstAMPM.AddItem "PM"
'Enable UnSet Button
cmdUnSet.Enabled = False
'By Default set Options and Time Seted as False so,
'it doesnt create any probs.
AMPMSel = False
optSel = False

'Add hours in Hours Box
For i = 1 To 12
    txtHH.AddItem i
Next i
'Select first value from the Hour Box
txtHH.ListIndex = 0
'Add minutes in Minutes Box
For i = 0 To 59
    txtMM.AddItem i
Next i
'Select first value from the Minute Box
txtMM.ListIndex = 0
'Add Seconds in a Second Box
For i = 0 To 59
    txtSS.AddItem i
Next i
'Select first value from the Second Box
txtSS.ListIndex = 0

'Select first value from the AMPM box
lstAMPM.ListIndex = 0

'Select ShutDown Option by default
optSD.Value = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
    Sys = X / Screen.TwipsPerPixelX
Select Case Sys
Case WM_RBUTTONDOWN:
    Me.Show
    'Remove System Tray Icon
    Shell_NotifyIcon NIM_DELETE, nid
End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Remove System Tray Icon
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lstAMPM_Click()
'Store AM/PM Selected of not
If lstAMPM.Text = "AM" Then
    AMPMSel = True
    Exit Sub
ElseIf lstAMPM.Text = "PM" Then
    AMPMSel = True
    Exit Sub
Else
    AMPMSel = False
    Exit Sub
End If
End Sub

Private Sub optATX_Click()
If optATX.Value = True Then
    optSel = True
Else
    optSel = False
End If
End Sub

Private Sub optLogOff_Click()
If optLogOff.Value = True Then
    optSel = True
Else
    optSel = False
End If
End Sub

Private Sub optRestart_Click()
If optRestart.Value = True Then
    optSel = True
Else
    optSel = False
End If
End Sub

Private Sub optSD_Click()
If optSD.Value = True Then
    optSel = True
Else
    optSel = False
End If
End Sub

Private Sub Timer1_Timer()
'Displays Current Time
lblCurTime.Caption = Format(Time, "h:m:s AMPM")
End Sub

Private Sub Timer2_Timer()
'Call Class
Dim l_ExitWindows As New cls_ExitWindows
'Continously check the time and option
If Format(Time, "h:m:s AMPM") = txtHH.Text + ":" + txtMM.Text + ":" + txtSS.Text + " " + lstAMPM.Text Then
    'Check which option it is and do like that
    If optATX.Value = True Then
        l_ExitWindows.ExitWindows WE_POWEROFF
    ElseIf optSD.Value = True Then
        l_ExitWindows.ExitWindows WE_SHUTDOWN
    ElseIf optRestart.Value = True Then
        l_ExitWindows.ExitWindows WE_REBOOT
    ElseIf optLogOff.Value = True Then
        l_ExitWindows.ExitWindows WE_LOGOFF
    End If
End If
End Sub
