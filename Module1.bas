Attribute VB_Name = "Module1"
'I have coded for 5 timers. Most systems can handle 16 Multimedia Timers.
'To add another timer, just duplicate one of the Timer Subs and add another
'Case to the Select Case structure.

Option Explicit

'Multimedia Timer APIs
Public Declare Function timeKillEvent Lib "winmm.dll" (ByVal uID As Long) As Long
Public Declare Function timeSetEvent Lib "winmm.dll" (ByVal uDelay As Long, ByVal uResolution As Long, ByVal lpFunction As Long, ByVal dwUser As Long, ByVal uFlags As Long) As Long

'Constants used by the Multimedia Timer
Public Const TIME_ONESHOT = 0  'Event occurs once, after uDelay milliseconds.
Public Const TIME_PERIODIC = 1  'Event occurs every uDelay milliseconds.
Public Const TIME_CALLBACK_EVENT_PULSE = &H20  'When the timer expires, Windows calls thePulseEvent function to pulse the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.
Public Const TIME_CALLBACK_EVENT_SET = &H10  'When the timer expires, Windows calls theSetEvent function to set the event pointed to by the lpTimeProc parameter. The dwUser parameter is ignored.
Public Const TIME_CALLBACK_FUNCTION = &H0   'When the timer expires, Windows calls the function pointed to by the lpTimeProc parameter. This is the default.

'Public variables used to test timers
Public MMTimer1 As Long
Public MMTimer2 As Long
Public MMTimer3 As Long
Public MMTimer4 As Long
Public MMTimer5 As Long

'Public variables used to store timer IDs
Public MMTimerID1 As Long
Public MMTimerID2 As Long
Public MMTimerID3 As Long
Public MMTimerID4 As Long
Public MMTimerID5 As Long

'Code for Timer 1 goes here
Private Sub MMTimer1_Timer()
    MMTimer1 = MMTimer1 + 1
End Sub

'Code for Timer 2 goes here
Private Sub MMTimer2_Timer()
    MMTimer2 = MMTimer2 + 1
End Sub

'Code for Timer 3 goes here
Private Sub MMTimer3_Timer()
    MMTimer3 = MMTimer3 + 1
End Sub

'Code for Timer 4 goes here
Private Sub MMTimer4_Timer()
    MMTimer4 = MMTimer4 + 1
End Sub

'Code for Timer 5 goes here
Private Sub MMTimer5_Timer()
    MMTimer5 = MMTimer5 + 1
End Sub

'Public sub that is called by system each time a timer event fires
Public Sub TimerProc(ByVal uID As Long, ByVal uMsg As Long, ByVal dwUser As Long, ByVal dw1 As Long, ByVal dw2 As Long)
    'Check Timer ID passed by system callback event
    'and call sub containing code for that timer
    Select Case uID
        Case Is = MMTimerID1
            'Call sub containing code for Timer 1
            Call MMTimer1_Timer
        Case Is = MMTimerID2
            'Call sub containing code for Timer 2
            Call MMTimer2_Timer
        Case Is = MMTimerID3
            'Call sub containing code for Timer 3
            Call MMTimer3_Timer
        Case Is = MMTimerID4
            'Call sub containing code for Timer 4
            Call MMTimer4_Timer
        Case Is = MMTimerID5
            'Call sub containing code for Timer 5
            Call MMTimer5_Timer
    End Select
End Sub
