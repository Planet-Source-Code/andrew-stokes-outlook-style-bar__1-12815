VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outlook style bar"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   1110
      TabIndex        =   13
      Top             =   60
      Width           =   3405
   End
   Begin VB.PictureBox Picture4 
      Height          =   3525
      Left            =   90
      ScaleHeight     =   3465
      ScaleWidth      =   885
      TabIndex        =   0
      Top             =   60
      Width           =   945
      Begin VB.CommandButton Command4 
         Height          =   165
         Left            =   120
         Picture         =   "Form1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3270
         Width           =   645
      End
      Begin VB.Frame Frame4 
         BorderStyle     =   0  'None
         Height          =   3015
         Left            =   30
         TabIndex        =   2
         Top             =   210
         Width           =   825
         Begin VB.PictureBox Picture5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   8205
            Left            =   60
            ScaleHeight     =   8205
            ScaleWidth      =   735
            TabIndex        =   3
            Top             =   0
            Width           =   735
            Begin VB.Image Image8 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":00A2
               Top             =   7410
               Width           =   630
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Caption         =   "Help"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               TabIndex        =   11
               Top             =   7950
               Width           =   735
            End
            Begin VB.Image Image7 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":1264
               Top             =   6450
               Width           =   630
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   "Settings"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               TabIndex        =   10
               Top             =   6990
               Width           =   735
            End
            Begin VB.Image Image3 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":2426
               Top             =   1950
               Width           =   630
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "Pause"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               TabIndex        =   9
               Top             =   2490
               Width           =   735
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   "Mailing List"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   8
               Top             =   5790
               Width           =   735
            End
            Begin VB.Image Image2 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":35E8
               Top             =   5250
               Width           =   630
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "View Log"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   0
               TabIndex        =   7
               Top             =   4590
               Width           =   735
            End
            Begin VB.Image Image1 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":47AA
               Top             =   4050
               Width           =   630
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               Caption         =   "Save Emails"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   6
               Top             =   3450
               Width           =   735
            End
            Begin VB.Image Image6 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":596C
               Top             =   2910
               Width           =   630
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               Caption         =   "Stop"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               TabIndex        =   5
               Top             =   1530
               Width           =   735
            End
            Begin VB.Image Image5 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":6B2E
               Top             =   990
               Width           =   630
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   "Start"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   0
               TabIndex        =   4
               Top             =   570
               Width           =   735
            End
            Begin VB.Image Image4 
               Height          =   525
               Left            =   30
               Picture         =   "Form1.frx":7CF0
               Top             =   0
               Width           =   630
            End
         End
      End
      Begin VB.CommandButton Command3 
         Enabled         =   0   'False
         Height          =   165
         Left            =   120
         Picture         =   "Form1.frx":8EB2
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   645
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Outlook Style Bar by Andrew Stokes (grimbyte1@yahoo.com)
'
'This Idea came to me out of the blue, I was just thinking
'of ways to make my apps appearance better so I decided
'to try make an outlook type bar
'
'I've seen alot of other methods to do a similar thing and
'they are all very complicated, but this is very easy. Hope
'You enjoy and find this code usful
'
'You may use this code freely but please tell me
'if you use it, I like to know where my code is going
'
'All the buttons are found in Picture5
'Picture5 is inside Frame4
'Frame4 is inside Picture4
'
'You dont need to change much to use it, just add buttons
'and then set what they must do
'
'If you like this code, please vote for me at Planet Source
'Code. Thanks
'
'For freeware programs made by me, goto
'http://trafficattractor2k.hypermart.net

Private Sub Command3_Click()
If Picture5.Top >= 0 Then Picture5.Top = 0 Else Picture5.Top = Picture5.Top + 450: Command4.Enabled = True: If Picture5.Top >= 0 Then Picture5.Top = 0: Command3.Enabled = False 'Check to see if it can go up or down and disable the buttons if it cant and move up or down if it can
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Command4_Click()
If Picture5.Top <= Frame4.Height - Picture5.Height Then Picture5.Top = Frame4.Height - Picture5.Height Else Picture5.Top = Picture5.Top - 450: Command3.Enabled = True: If Picture5.Top <= Frame4.Height - Picture5.Height Then Picture5.Top = Frame4.Height - Picture5.Height: Command4.Enabled = False 'Check to see if it can go up or down and disable the buttons if it cant and move up or down if it can
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed 'Make sure the buttons dont stay pressed
End Sub
Private Sub Form_Load()
ClearSelections 'Make sure the buttons dont stay pressed 'Make sure the buttons dont stay pressed
MsgBox "Have fun ;>, I hope you find this code useful and dont forget to vote for me at Planet Source Code", , "Wazup"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed 'Make sure the buttons dont stay pressed
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.BorderStyle = 1 'Set this button to be pressed in
Image5.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image5.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub

Private Sub Image3_Click()
List1.AddItem "Item3 Clicked"

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image5.BorderStyle = 0
Image4.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub

Private Sub Image4_Click()
List1.AddItem "Item1 Clicked"
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image5.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub

Private Sub Image5_Click()
List1.AddItem "Item2 Clicked"

End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub

Private Sub Image6_Click()
List1.AddItem "Item4 Clicked"

End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image5.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image5.BorderStyle = 0
Image6.BorderStyle = 0
Image8.BorderStyle = 0
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.BorderStyle = 1 'Set this button to be pressed in
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image5.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
End Sub


Private Sub Label16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label19_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label20_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Label22_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub List1_Click()
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Text1_Change()
ClearSelections 'Make sure the buttons dont stay pressed
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
ClearSelections 'Make sure the buttons dont stay pressed
End Sub
Sub ClearSelections() 'Make sure the buttons dont stay pressed()
Image1.BorderStyle = 0
Image2.BorderStyle = 0
Image3.BorderStyle = 0
Image4.BorderStyle = 0
Image5.BorderStyle = 0
Image6.BorderStyle = 0
Image7.BorderStyle = 0
Image8.BorderStyle = 0

End Sub
