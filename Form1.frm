VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   5040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2520
      TabIndex        =   10
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   8
      Top             =   2640
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "..."
      Height          =   255
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1080
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      Caption         =   "..."
      Height          =   255
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "XML Upload Folder"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Display XML file:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "Select Table:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Convert into XML"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Select Access Database:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Declare the Class
Private eTextBox As clsETextBox
Private Con As ADODB.Connection
Private RS As ADODB.Recordset
Private Type OPENFILENAME
lStructSize As Long
hWndOwner As Long
hInstance As Long
lpstrFilter As String
lpstrCustomFilter As String
nMaxCustFilter As Long
nFilterIndex As Long
lpstrFile As String
nMaxFile As Long
lpstrFileTitle As String
nMaxFileTitle As Long
lpstrInitialDir As String
lpstrTitle As String
flags As Long
nFileOffset As Integer
nFileExtension As Integer
lpstrDefExt As String
lCustData As Long
lpfnHook As Long
lpTemplateName As String
End Type
Private Type BrowseInfo
   hWndOwner As Long
   pIDLRoot As Long
   pszDisplayName As Long
   lpszTitle As Long
   ulFlags As Long
   lpfnCallback As Long
   lParam As Long
   iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2
Private Const MAX_PATH = 260

Private Declare Function SHBrowseForFolder Lib _
"shell32" (lpbi As BrowseInfo) As Long

Private Declare Function SHGetPathFromIDList Lib _
"shell32" (ByVal pidList As Long, ByVal lpBuffer _
As String) As Long

Private Declare Function lstrcat Lib "kernel32" _
Alias "lstrcatA" (ByVal lpString1 As String, ByVal _
lpString2 As String) As Long

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias _
"GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Combo1_GotFocus()
    If Combo1.ListIndex = -1 Then
        Combo1.AddItem Combo1.Text
        SendKeys (Combo1.Text)
        'SendKeys "%{Down}"
    End If
End Sub

Private Sub Command1_Click()
    Combo1.Text = ""
    Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hWndOwner = hwnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "Access Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) _
+ "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0)


ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = CurDir
ofn.lpstrTitle = "Our File Open Title"
ofn.flags = 0
Dim lngRetVal As Long
lngRetVal = GetOpenFileName(ofn)

If (lngRetVal) Then
'MsgBox "File to Open: " + Trim$(ofn.lpstrFile)
Text1.Text = Trim$(ofn.lpstrFile)
Else
Text1.Text = ""
MsgBox "Cancel was pressed"
End If


End Sub

Private Sub Command2_Click()
Dim j As Integer
    For j = 0 To Combo1.ListCount - 1
Combo1.List(j) = ""
    Next
    If Combo1.ListIndex = -1 Then
        Combo1.AddItem Combo1.Text
            SendKeys (Combo1.Text)
            SendKeys "%{Down}"
    End If
    If Combo1.Text = "" Then
        Combo1.Clear
    End If
    
    
        Dim sTableName As String
Dim i As Integer
Dim Con As New ADODB.Connection
Con.Open "Provider=Microsoft.Jet.OLEDB.3.51;DATA SOURCE=" & Trim(Text1.Text) & ";"
'Set RS = Con.Execute("Select * from providers")
    
Set RS = Con.OpenSchema(adSchemaTables)
Do While Not RS.EOF
    If (left(RS!TABLE_NAME, 4) <> "MSys") And (left(RS!TABLE_NAME, 1) <> "~") Then
    Combo1.AddItem RS!TABLE_NAME
    End If
    RS.MoveNext
Loop
End Sub


Private Sub Form_Load()
       'Size form to fill most of the screen
    'Width = Screen.Width * 0.8
    'Height = Screen.Height * 0.8
    'Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
   ' Capture the Title of the program
    Caption = App.Title
    
    ' Assign the textbox
    Set eTextBox = New clsETextBox
    
    With eTextBox
        Set .eTextBox = Text2
        'Set .eTextBox = Form2.txtText
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    ' Free the class
    Set eTextBox = Nothing
End Sub


Private Sub Label1_Click()
Combo1.Text = ""
    Dim ofn As OPENFILENAME
ofn.lStructSize = Len(ofn)
ofn.hWndOwner = hwnd
ofn.hInstance = App.hInstance
ofn.lpstrFilter = "Access Database (*.mdb)" + Chr$(0) + "*.mdb" + Chr$(0) _
+ "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0)


ofn.lpstrFile = Space$(254)
ofn.nMaxFile = 255
ofn.lpstrFileTitle = Space$(254)
ofn.nMaxFileTitle = 255
ofn.lpstrInitialDir = CurDir
ofn.lpstrTitle = "Our File Open Title"
ofn.flags = 0
Dim lngRetVal As Long
lngRetVal = GetOpenFileName(ofn)

If (lngRetVal) Then
'MsgBox "File to Open: " + Trim$(ofn.lpstrFile)
Text1.Text = Trim$(ofn.lpstrFile)
Else
Text1.Text = ""
MsgBox "Cancel was pressed"
End If

End Sub

Private Sub Label2_Click()
Dim sTableName As String
Dim i As Integer
Dim Con As New ADODB.Connection
Con.Open "Provider=Microsoft.Jet.OLEDB.3.51;DATA SOURCE=" & Trim(Text1.Text) & ";"
'Set RS = Con.Execute("Select * from providers")
Set RS = Con.Execute("Select * from " & Combo1.Text & ";")
     Open Text3.Text & "\" & Combo1.Text & ".xml" _
        For Output Access Write As #1
        'For Append As #1
        If EOF(1) Then
        Print #1, "<?xml version='1.0' ?>"
        Print #1, "<!-- " & Combo1.Text & ".xml -->"
        Print #1, "<" & Combo1.Text & ">"
    Do While Not RS.EOF
        For i = 0 To RS.Fields.Count - 1
            Print #1, "<" & RS.Fields(i).Name & ">" & RS.Fields(i) & "</" & RS.Fields(i).Name & ">"
      
        Next
    RS.MoveNext
    Loop
        Print #1, "</" & Combo1.Text & ">"
        Close #1
          End If
End Sub

Private Sub Label3_Click()
    Dim j As Integer
    For j = 0 To Combo1.ListCount - 1
Combo1.List(j) = ""
    Next
    If Combo1.ListIndex = -1 Then
        Combo1.AddItem Combo1.Text
            SendKeys (Combo1.Text)
            SendKeys "%{Down}"
    End If
    If Combo1.Text = "" Then
        Combo1.Clear
    End If
    
    
        Dim sTableName As String
Dim i As Integer
Dim Con As New ADODB.Connection
Con.Open "Provider=Microsoft.Jet.OLEDB.3.51;DATA SOURCE=" & Trim(Text1.Text) & ";"
'Set RS = Con.Execute("Select * from providers")
    
Set RS = Con.OpenSchema(adSchemaTables)
Do While Not RS.EOF
    If (left(RS!TABLE_NAME, 4) <> "MSys") And (left(RS!TABLE_NAME, 1) <> "~") Then
    Combo1.AddItem RS!TABLE_NAME
    End If
    RS.MoveNext
Loop
End Sub

Private Sub Label4_Click()
 Dim strFileName As String
  With eTextBox
        .Clear
    End With
 strFileName = Text3.Text & "\" & Combo1.Text & ".xml"
  With eTextBox
        .LoadTEXT strFileName
   End With
End Sub
    


Private Sub Label5_Click()
    'Opens a Browse Folders Dialog Box that displays the
'directories in your computer
Dim lpIDList As Long 'Declare Varibles
Dim sBuffer As String
Dim szTitle As String
Dim tBrowseInfo As BrowseInfo

szTitle = "Hello World. Click on a directory and " & _
"it's path will be displayed in a message box"
'Text to appear in the the gray area under the title bar
'telling you what to do
With tBrowseInfo
   .hWndOwner = Me.hwnd 'Owner Form
   .lpszTitle = lstrcat(szTitle, "")
   .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
End With
lpIDList = SHBrowseForFolder(tBrowseInfo)
If (lpIDList) Then
   sBuffer = Space(MAX_PATH)
   SHGetPathFromIDList lpIDList, sBuffer
   sBuffer = left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
   'MsgBox sBuffer
   Text3.Text = sBuffer
End If
End Sub
