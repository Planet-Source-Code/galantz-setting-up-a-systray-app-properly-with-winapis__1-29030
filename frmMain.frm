VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPopup 
      Caption         =   "myPopup"
      Begin VB.Menu mnuChange 
         Caption         =   "Change Tool Tip Text"
      End
      Begin VB.Menu mnuPopupSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'***********************************************************************
'Comments by the Author: Gary Lantz
'***********************************************************************

'I've picked up code from other people at PSC
'(sorry don't know who all), from textbooks, online
'tutorials, and just general asking questions from
'other coders and even occasionally from msdn.microsoft.com
'SOOOO if you feel so compelled to believe that you
'are the original author and NO ONE, even the microsoft
'programmers who developed the codes in the firstplace
'and wrapped them up into what we call Visual Basic, then
'by all means, contact me and I will be more than happy to
'add your name to the who's who in Gods of programming.

'Take care,

'Gary Lantz
'galantz@netzero.net

'***********************************************************************
'BEFORE YOU BEGIN!!!!
'***********************************************************************
'1. The icon that the form has will be the icon that is in the systray!
'2. If you do NOT want to see the form and ONLY work from the systray, _
    then set the form.visible = false
'3. If you want to use the form, but ALSO want a systray, what I do is _
    create a systray menu, call it mnuPopUp, then set its visible property _
    to false...then use it as a popupmenu mnuPopUP from the systray.  This _
    looks the cleanest.
'4. please READ EVERYTHING before spouting off ignorant comments to other _
    coders.  We are all here to help, not bash each other.
'5. Have fun while you program!
'***********************************************************************

Private strMyToolTip As String

Private Sub Form_Load()

  'load the systray icon
  'call AddSystray(Me,"Tool Tip Text")

    strMyToolTip = "Right Click for Server Options"
    Call AddSystray(Me, strMyToolTip)

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errorH
  Dim rtn As Long

  'this procedure receives the callbacks from the
  'System Tray icon and pops up the menu if the right
  'button is clicked.  I left the other button options
  'there, incase you want to have other options...

    'just incase we are trying to process something, let's block this out...
    'the value of X will vary depending on the scalemode setting
    If Me.ScaleMode = vbPixels Then
        rtn = X
      Else
        rtn = X / Screen.TwipsPerPixelX
    End If

    Select Case rtn
      Case WM_LBUTTONDOWN         '= &H201 - Left Button down
        'nothing happens, yet
      Case WM_LBUTTONUP           '= &H202 - Left Button up
        'nothing happens, yet
      Case WM_LBUTTONDBLCLK       '= &H203 - Left Double-click
        'nothing happens, yet
      Case WM_RBUTTONDOWN         '= &H204 - Right Button down
        'nothing happens, yet
      Case WM_RBUTTONUP           '= &H205 - Right Button up
        SetForegroundWindow Me.hWnd
        Me.PopupMenu Me.mnuPopup
      Case WM_RBUTTONDBLCLK       '= &H206 - Right Double-click
        'nothing happens, yet
    End Select

Exit Sub

errorH:
    MsgBox "Sub Form_MouseMove, frmMain.frm", Err.Number, Err.Description
End Sub

Private Sub mnuChange_Click()

  'here is a quick routine to cycle a tooltip
  'message for the systray application
  'you could use this to cycle the form
  'icon and therefore change the systray.

  'for quick and dirty animated systray apps.
  'I usually create an image list of each frame
  'cycle through them with a timer...it's very effective
  'and easy to do

    If (strMyToolTip = "Right Click for Menu") Then
        strMyToolTip = "This is the new tooltip"
      Else
        strMyToolTip = "Right Click for Menu"
    End If

    'the actual call to modify the systray icon
    Call ModifySystray(Me, strMyToolTip)

End Sub

Private Sub mnuExit_Click()

    Call RemoveSystray
    Unload Me
    End

End Sub

' Copyright 2001 Gary Lantz
