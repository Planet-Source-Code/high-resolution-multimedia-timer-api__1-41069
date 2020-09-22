VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "High Resolution Timer"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   142
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command6 
      Caption         =   "Start"
      Height          =   300
      Left            =   1950
      TabIndex        =   15
      Top             =   1275
      Width           =   540
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Start"
      Height          =   300
      Left            =   1950
      TabIndex        =   14
      Top             =   975
      Width           =   540
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Start"
      Height          =   300
      Left            =   1950
      TabIndex        =   13
      Top             =   675
      Width           =   540
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Start"
      Height          =   300
      Left            =   1950
      TabIndex        =   12
      Top             =   375
      Width           =   540
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      Height          =   300
      Left            =   1950
      TabIndex        =   11
      Top             =   75
      Width           =   540
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1350
      TabIndex        =   9
      Text            =   "25"
      Top             =   1275
      Width           =   540
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1350
      TabIndex        =   7
      Text            =   "10"
      Top             =   975
      Width           =   540
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1350
      TabIndex        =   5
      Text            =   "5"
      Top             =   675
      Width           =   540
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1350
      TabIndex        =   3
      Text            =   "2"
      Top             =   375
      Width           =   540
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1350
      TabIndex        =   1
      Text            =   "1"
      Top             =   75
      Width           =   540
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test All Timers for 5 Seconds"
      Height          =   450
      Left            =   75
      TabIndex        =   0
      Top             =   1650
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   75
      Top             =   2850
   End
   Begin VB.Label Label5 
      Caption         =   "Timer 5 Interval"
      Height          =   300
      Left            =   75
      TabIndex        =   10
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Timer 4 Interval"
      Height          =   300
      Left            =   75
      TabIndex        =   8
      Top             =   975
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Timer 3 Interval"
      Height          =   300
      Left            =   75
      TabIndex        =   6
      Top             =   675
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Timer 2 Interval"
      Height          =   300
      Left            =   75
      TabIndex        =   4
      Top             =   375
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Timer 1 Interval"
      Height          =   300
      Left            =   75
      TabIndex        =   2
      Top             =   75
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Very simple example of using Multimedia Timers in your application.
'Multimedia Timers can have an interval as low as 1 ms and in most
'cases can fire the whole 1000 times in 1 second. VB's timer control
'is only capable of firing every 10 ms on most Operating Systems and
'on some older Operating Systems can only fire every 50 ms. To start
'timer just call the timeSetEvent API with the interval and options
'coded below. The code used below will start a recurring timer event
'that will only end when the timeKillEvent API is called using the
'returned Timer ID from calling the timeSetEvent API. Be sure you kill
'all timers after your application is closed or the timers will not stop
'until you restart your machine. This can eat up your free system
'resources, possibly causing your other applications to run slower.
'If you have any questions about this code or any of my other code
'submissions at PSC, feel free to email me at battlestorm@cox.net
'I learned how to use these API calls from www.allapi.net
'I suggest you go there and download their API Guide if you want to
'learn more about Windows API that can be used in Visual Basic.

Option Explicit

Private Sub Command1_Click()
    'Reset counters
    MMTimer1 = 0
    MMTimer2 = 0
    MMTimer3 = 0
    MMTimer4 = 0
    MMTimer5 = 0
    
    'Disable buttons for duration of test
    Command1.Enabled = False
    Command2.Enabled = False
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    Command6.Enabled = False
    
    'Start test timer
    Timer1.Enabled = True
    
    'Start all Multimedia Timers
    MMTimerID1 = timeSetEvent(Text1.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    MMTimerID2 = timeSetEvent(Text2.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    MMTimerID3 = timeSetEvent(Text3.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    MMTimerID4 = timeSetEvent(Text4.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    MMTimerID5 = timeSetEvent(Text5.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
End Sub

Private Sub Command2_Click()
    'Start or Stop Timer 1 and display how many times it fired
    If Command2.Caption = "Start" Then
        MMTimer1 = 0
        Command2.Caption = "Stop"
        MMTimerID1 = timeSetEvent(Text1.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        Command2.Caption = "Start"
        timeKillEvent MMTimerID1
        MsgBox "HR Timer 1 was called " & MMTimer1 & " times."
    End If
End Sub

Private Sub Command3_Click()
    'Start or Stop Timer 2 and display how many times it fired
    If Command3.Caption = "Start" Then
        MMTimer2 = 0
        Command3.Caption = "Stop"
        MMTimerID2 = timeSetEvent(Text2.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        Command3.Caption = "Start"
        timeKillEvent MMTimerID2
        MsgBox "HR Timer 2 was called " & MMTimer2 & " times."
    End If
End Sub

Private Sub Command4_Click()
    'Start or Stop Timer 3 and display how many times it fired
    If Command4.Caption = "Start" Then
        MMTimer3 = 0
        Command4.Caption = "Stop"
        MMTimerID3 = timeSetEvent(Text3.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        Command4.Caption = "Start"
        timeKillEvent MMTimerID3
        MsgBox "HR Timer 3 was called " & MMTimer3 & " times."
    End If
End Sub

Private Sub Command5_Click()
    'Start or Stop Timer 4 and display how many times it fired
    If Command5.Caption = "Start" Then
        MMTimer4 = 0
        Command5.Caption = "Stop"
        MMTimerID4 = timeSetEvent(Text4.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        Command5.Caption = "Start"
        timeKillEvent MMTimerID4
        MsgBox "HR Timer 4 was called " & MMTimer4 & " times."
    End If
End Sub

Private Sub Command6_Click()
    'Start or Stop Timer 5 and display how many times it fired
    If Command6.Caption = "Start" Then
        MMTimer5 = 0
        Command6.Caption = "Stop"
        MMTimerID5 = timeSetEvent(Text5.Text, 0, AddressOf TimerProc, 0, TIME_PERIODIC Or TIME_CALLBACK_FUNCTION)
    Else
        Command6.Caption = "Start"
        timeKillEvent MMTimerID5
        MsgBox "HR Timer 5 was called " & MMTimer5 & " times."
    End If
End Sub

Private Sub Timer1_Timer()
    'Enable buttons
    Command1.Enabled = True
    Command2.Enabled = True
    Command3.Enabled = True
    Command4.Enabled = True
    Command5.Enabled = True
    Command6.Enabled = True

    'Stop test timer
    Timer1.Enabled = False
    
    'Stop all Multimedia Timers
    timeKillEvent MMTimerID1
    timeKillEvent MMTimerID2
    timeKillEvent MMTimerID3
    timeKillEvent MMTimerID4
    timeKillEvent MMTimerID5
    
    'Display how many times each timer was fired
    MsgBox "HR Timer 1 was called " & MMTimer1 & " times." & vbCrLf & _
           "HR Timer 2 was called " & MMTimer2 & " times." & vbCrLf & _
           "HR Timer 3 was called " & MMTimer3 & " times." & vbCrLf & _
           "HR Timer 4 was called " & MMTimer4 & " times." & vbCrLf & _
           "HR Timer 5 was called " & MMTimer5 & " times."
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Ensure that all timers have been stopped to preserve system resources
    Timer1.Enabled = False
    timeKillEvent MMTimerID1
    timeKillEvent MMTimerID2
    timeKillEvent MMTimerID3
    timeKillEvent MMTimerID4
    timeKillEvent MMTimerID5
End Sub
