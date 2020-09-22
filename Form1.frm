VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "VB 6.0 List Box Resize Demo"
   ClientHeight    =   6630
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   8835
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   11
      Left            =   495
      TabIndex        =   11
      Text            =   "txt(11)"
      Top             =   6045
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   10
      Left            =   495
      TabIndex        =   10
      Text            =   "txt(10)"
      Top             =   5535
      Width           =   3075
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5550
      TabIndex        =   14
      Top             =   2190
      Width           =   1260
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Fill List Box"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4020
      TabIndex        =   13
      Top             =   2190
      Width           =   1260
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   9
      Left            =   495
      TabIndex        =   9
      Text            =   "txt(9)"
      Top             =   5040
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   8
      Left            =   495
      TabIndex        =   8
      Text            =   "txt(8)"
      Top             =   4500
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   7
      Left            =   495
      TabIndex        =   7
      Text            =   "txt(7)"
      Top             =   3945
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   6
      Left            =   495
      TabIndex        =   6
      Text            =   "txt(6)"
      Top             =   3405
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   495
      TabIndex        =   5
      Text            =   "txt(5)"
      Top             =   2865
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   4
      Left            =   495
      TabIndex        =   4
      Text            =   "txt(4)"
      Top             =   2325
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   495
      TabIndex        =   3
      Text            =   "txt(3)"
      Top             =   1785
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   495
      TabIndex        =   2
      Text            =   "txt(2)"
      Top             =   1230
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   495
      TabIndex        =   1
      Text            =   "txt(1)"
      Top             =   768
      Width           =   3075
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   495
      TabIndex        =   0
      Text            =   "txt(0)"
      Top             =   225
      Width           =   3075
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1710
      Left            =   3870
      TabIndex        =   12
      Top             =   240
      Width           =   3075
   End
   Begin VB.Image imgCheck 
      Appearance      =   0  'Flat
      Height          =   300
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label lblInstructions 
      BackColor       =   &H00C0FFC0&
      Caption         =   "lblInstructions"
      Height          =   2985
      Left            =   3990
      TabIndex        =   16
      Top             =   3195
      Width           =   3405
   End
   Begin VB.Label lblTextSizer 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblTextSizer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7140
      TabIndex        =   17
      Top             =   285
      Visible         =   0   'False
      Width           =   1560
   End
   Begin VB.Label lblBorder 
      BackColor       =   &H0080FF80&
      Height          =   3255
      Left            =   3870
      TabIndex        =   15
      Top             =   3060
      Width           =   3645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' VB list box resizing demo
' by Andy Fielding (ander5151@yahoo.com)
'
' This demo shows how to control the width of a standard list box's text area
' so its widest string is always visible.
'
' When VB 's list box control contains too many items to show vertically,
' a vertical scrollbar automatically appears. This demo shows how to detect
' the scrollbar, find its width, and include it when setting the list box's width.

Option Explicit
 
 Private Sub Form_Load()
 
   Dim X As Integer
  
   ' Seed the random-number generator (just for this demo)
   Randomize
  
   lblInstructions.Caption = "1. If you wish, change any of the strings in the " & _
     vbCrLf & "     text boxes." & vbCrLf & vbCrLf & "2. Click the Fill " & _
     "List Box button." & vbCrLf & vbCrLf & "3. A random number of strings are copied to " & _
     vbCrLf & "     the list box. If more than 7 strings are copied, " & vbCrLf & _
     "     the list box's vertical scrollbar automatically " & vbCrLf & "     appears." & _
     vbCrLf & vbCrLf & "4. The list box expands to the width of the " & vbCrLf & "     longest string, " & _
     "compensating for the" & vbCrLf & "     width of the scrollbar (if present). " & _
     "If the " & vbCrLf & "     longest string isn't visible, you can scroll " & _
     vbCrLf & "     down to see it."
  
    ' Fill the text boxes with sample strings
    txt(0) = "Sample text begins here."
    txt(1) = "This is a string."
    txt(2) = "Here is a longer string."
    txt(3) = "And, amazingly, yet a longer string."
    txt(4) = "Finally, here's a very, very long string, dudes!!"
    txt(5) = "Some short text..."
    txt(6) = "Some longer text..."
    txt(7) = "Some slightly longer text..."
    txt(8) = "And, amazingly, some yet longer text."
    txt(9) = "I say, have you seen the cat?"
    txt(10) = "Yes, he's moved to Florida."
    txt(11) = "But he didn't finish his Spam!"
    
    ' load & position the hidden checkmark images
    For X = 1 To 11
      Load imgCheck(X)
      imgCheck(X).Top = txt(X).Top + 15
      imgCheck(X).Left = imgCheck(0).Left
    Next X
    
End Sub
  
Private Sub Command1_Click()

  Dim X As Integer, Y As Integer
  Dim ScrollbarWidth As Long
  Dim WidestString As Integer
  
  ' For this demo, using a static variable lets the sub "remember" the
  ' last random number so we can use a different one each time
  Static RandInt As Integer
  
  List1.Clear  ' Clear the list box
 
  ' Add text to the list box
  '
  ' We want to handle the list box with a scrollbar and without one---
  ' so for this demo, we'll add a random number of strings on each click.
  ' When more than seven are added, the scrollbar automatically appears
  '
  Do
    X = Rnd * 11
  Loop While X = RandInt  ' Find a random number we didn't use last time

  For Y = 0 To X
    List1.AddItem txt(Y).Text   ' Add the strings to the list box
  Next Y
  
  For Y = 0 To 11
    imgCheck(Y).Visible = (Y <= X)  ' Make checkmarks visible next to the
  Next Y                            ' selected strings, and hide the rest
  
  RandInt = X  ' Remember this number
  
  ' Find the width of the list box's longest string by making each string the
  ' caption of an auto-sizing label, then using the label's widest width
  ' (Note: This is a lot easier than the ton of API calls Microsoft suggests. ;?)
  For X = 0 To 11
    lblTextSizer.Caption = List1.List(X)
    If lblTextSizer.Width > WidestString Then WidestString = lblTextSizer.Width
  Next X

  ' Add a bit extra so the end of the widest string will be visible
  WidestString = WidestString + 100
  
  ' Find out if the list box has a scrollbar by sending its name to
  ' the scrollbar-check procedure
  If HasVerticalScrollbar(List1) Then
    '
    ' It has a scrollbar, so:
    '
    ' (1) Use the GetSystemMetrics API to get the user's vertical scrollbar
    ' width setting. (We could probably just do this once at load. However,
    ' technically, it's possible for the user to change the system's scrollbar
    ' setting while the app is running---so you be the judge)
    ScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL)
    '
    ' (2) Make the list box as wide as:
    '   the widest string
    '   + the scrollbar width (converted from pixels to twips)
    List1.Width = WidestString + Me.ScaleX(ScrollbarWidth, vbPixels, vbTwips)
    '
  Else
    '
    ' It doesn't have a scrollbar, so just make it wide enough
    ' for the widest string
    List1.Width = WidestString
    '
  End If

End Sub

Private Sub cmdExit_Click()

  Unload Me

End Sub
