VERSION 5.00
Begin VB.Form frmSysTray 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   $"frmSystray.frx":0000
      Height          =   1065
      Left            =   450
      TabIndex        =   0
      Top             =   75
      Width           =   3840
   End
   Begin VB.Menu Menu 
      Caption         =   "File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As.."
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print"
      End
      Begin VB.Menu mnuPrintSetup 
         Caption         =   "Print Setup"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created By Michael Cowell
'VorTech Software Web: www.vortech.freeservers.com
'Copyrighted 2000-2001

'Please Vote For This Code On www.planetsourcecode.com



Private Sub Form_Initialize()
'This gets Loaded when your form starts
try.cbSize = Len(try)
try.hwnd = Me.hwnd
try.uId = vbNull
try.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
try.uCallBackMessage = WM_MOUSEMOVE

'To Change the Icon Displayed in the systray
'Change the Forms Icon
'This uses whatever Icon the Form Displays
try.hIcon = Me.Icon

'Tool Tip
try.szTip = "This Deserves A 5" & vbNullChar

Call Shell_NotifyIcon(NIM_ADD, try)
Call Shell_NotifyIcon(NIM_MODIFY, try)

'If u just want the systay icon to appear at start Hide the Form
'Me.Hide
End Sub

'Right Click and Dbl Click to launch an event

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case X
        Case 7755:   'Right Click
            PopupMenu Menu  'The systray menu works the same as
                            'clicking file on the form. Anything
                            'you can do with a menu on the form
                            'you can do in the systray.
            
        
        Case 7725:    'Dbl Left Click
            MsgBox "Dbl Click in the systray needs event"
    End Select
End Sub

Private Sub mnuQuit_Click()
End
End Sub
