VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKeyGen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Key Master by Chazter"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3495
   Icon            =   "KeyGen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5400
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   4560
      Width           =   3495
      Begin VB.CommandButton Command3 
         Caption         =   "&Register"
         Height          =   375
         Left            =   1200
         Picture         =   "KeyGen.frx":0442
         TabIndex        =   13
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2760
         Picture         =   "KeyGen.frx":0884
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   480
         Picture         =   "KeyGen.frx":0CC6
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Test"
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   3360
      Width           =   3495
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label4 
         Caption         =   "Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Name:"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Key Generator"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
      Begin VB.TextBox txtkey 
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtname 
         Height          =   285
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
      Begin VB.Label Label2 
         Caption         =   "Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   5
      Top             =   2520
      Width           =   3495
      Begin VB.CommandButton Command1 
         Caption         =   "&Generate"
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2760
         Picture         =   "KeyGen.frx":1108
         Top             =   240
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   360
         Picture         =   "KeyGen.frx":154A
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   $"KeyGen.frx":198C
      Height          =   855
      Left            =   0
      TabIndex        =   14
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmKeyGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------
'Hi thanks for testing out this little example of how to generate a unique key based on a username
'The code has been fully commented, for easier comprehension
'If you have any problems then email me at chazter_uk@yahoo.com
'
'Your are free to edit or distribute this source code, but i would like to be informed of where
'the code is!
'
'
'==============================================
'           Useful Websites (No Order)
'==============================================
'             1) www.vbweb.co.uk              '
'             2) www.vbsquare.co.uk           '
'             3) www.vbworld.co.uk            '
'             4) www.VBIP.com                 '
'             5) www.allapi.com               '
'             6) www.planet-source-code.com   '
'             7) www.programmingtutorials.com '
'==============================================
            
'I am also available on DalNet in channels: #VBWorld, #VisualBasic, #Programmers
'-------------------------------------------------------------------------------------------------
Option Explicit
Private Declare Function HideCaret Lib "user32" (ByVal Hwnd As Long) As Long 'Declares API function for hiding the caret
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long 'This declares the API function that we will use to automatically copy the text of txtkey
Const WM_COPY = &H301 'This is used to copy the text of txtkey in conjuntion with SendMessage()
Const WM_PASTE = &H302 'This is used to paste the text of txtkey in conjuntion with SendMessage()






Private Sub Command1_Click()

If Len(Trim(txtname.Text)) = 0 Then 'This checks to see if the username is all spaces
    MsgBox "Uhhh...try entering some characters aswell!", vbInformation, "" 'If so.. tell 'em!
    txtname.SetFocus 'Sets the focus on txtname
    Exit Sub 'Exits this sub
End If

If Len(Trim(txtname.Text)) < 6 Then 'This checks the length of the text and makes sure it's not less than 6
    MsgBox "Your user name but be at least 6 characters long!", vbInformation, "" 'Display warning if it is
    txtname.SetFocus 'Return to the form and set the focus on txtname
    Exit Sub 'Exits this sub
Else
    txtkey.Text = GenKey(Trim(txtname.Text)) 'If all is well, then call our function and place the result in txtkey
End If


End Sub

Public Function GenKey(Username As String) As String
Dim TVal As Long
Dim i As Integer            'Variable Declaration
Dim TText As String
Dim TString As String


    TString = "" 'Reset the variable
    pb.Max = Len(Username) 'Set the scroll bars MAX property to the length of the Username
    Me.Caption = Me.Caption & " [Building Key...]" 'Change the caption of the window while generating
    For i = 1 To Len(Username) ' Start the loop using the length of the username
        pb.Value = i 'Show the status of the generation
        TVal = Asc(Mid(Username, i, 1)) + 2 'Converts the next letter of username to it's ASCII value, then add's 2
        TVal = TVal + Fix((TVal * (16 + Len(Username)))) 'This adds the last result with 16 * the length of the username
        TVal = TVal + Len(Username) 'It adds to the last result the length of the Username
        TString = TString & Trim(StrReverse(Str(TVal))) 'This reverses the last result and appends it to the last result in TString
    Next i 'Continue getting the next letter in Username
    
    TText = TString 'This puts the generated key into TText
    
   If Len(TText) >= 8 Then 'This tests to see if the length of the key is 8 or greater
        Mid(TText, 4, 1) = "-" 'If so then place a hypen in the key
        Mid(TText, 12, 1) = "-" 'Place another hypen in the key
   End If
   
    TText = Left(TText, 16) 'This trims the key down making it look nice :o)
    Me.Caption = "Key Master by Chazter" 'This returns the caption to it's former state
    GenKey = TText 'This makes the function equal to the generated key
    
End Function

Private Sub Command3_Click()

If Len(Text1.Text) < 6 Then 'This checks to see if the length of text in Text1 is less than 6
    MsgBox "Your user name must be at least 6 characters long!", vbInformation, "" 'If so then show the warning
    Text1.SetFocus 'Return to the form and set the focus on text1
    Exit Sub 'Exit the sub
End If

If (Text1.Text = "") Or (Text2.Text = "") Then 'This checks to see if any of the two boxes are empty
    MsgBox "Please complete all the boxes!", vbInformation, "Error" 'If so then show the warning
    Exit Sub 'Exit the sub
End If

If Text2.Text <> GenKey(Text1.Text) Then 'This checks to see wether the text in text2 matches the generated key
    MsgBox "Incorrect Key for username '" & Text1.Text & "'", vbCritical, "" 'If not then show the bad Key message box
Else
    MsgBox "This Key is valid for " & Text1.Text & "!", vbInformation, "" 'If the key is good then show the Thankyou screen :o)
End If
End Sub

Private Sub txtkey_Change()
Dim Res As Long 'Declare variables
Dim Hwnd As Long

    Text2.Text = "" 'Deletes any text that might be in Text2
    Hwnd = txtkey.Hwnd 'This gets the handle of txtkey
    txtkey.SelLength = Len(txtkey.Text) 'In order to use the message WM_COPY, the text needs to be selected. This is whats happening here :)
    Call SendMessage(Hwnd, WM_COPY, 0&, 0&) 'This calls our SendMessage() function to automaically copy the text of txtkey
    Hwnd = Text2.Hwnd 'This gets the handle of Text2
    Call SendMessage(Hwnd, WM_PASTE, 0&, 0&) 'This pastes into text2 the key we copied ealier
    Text1.SetFocus 'Sets the focus on Text1
    
    'Now i bet your thinking why the hell didn't he just use text2.text=txtkey.text ???????
    'Well...quite simply...i couldn't be ARSED!
    'I fancied a change :o)
    
    'I thought to keep you motivated you could type the username yourself, let you do some of the work :þ
End Sub

Private Sub txtkey_GotFocus()
Dim Hwnd As Long 'Declare Variable

    Hwnd = txtkey.Hwnd 'This gets the handle of txtkey
    Call HideCaret(Hwnd) 'For a bit of fun we'll hide the caret in txtkey
End Sub
