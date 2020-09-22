VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{972DE6B5-8B09-11D2-B652-A1FD6CC34260}#1.0#0"; "ACTIVESKIN.OCX"
Begin VB.Form frmMainForm 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Skin Pad Plus"
   ClientHeight    =   3495
   ClientLeft      =   -150
   ClientTop       =   5475
   ClientWidth     =   9435
   Icon            =   "frmMainForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5495.282
   ScaleLeft       =   10000
   ScaleMode       =   0  'User
   ScaleWidth      =   11374.32
   WindowState     =   2  'Maximized
   Begin ACTIVESKINLibCtl.SkinButton cmdFindNext 
      Height          =   312
      Left            =   5760
      OleObjectBlob   =   "frmMainForm.frx":08CA
      TabIndex        =   16
      Top             =   45
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinButton cmdFind 
      Height          =   312
      Left            =   4440
      OleObjectBlob   =   "frmMainForm.frx":095C
      TabIndex        =   15
      Top             =   45
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage10 
      Height          =   480
      Left            =   8880
      OleObjectBlob   =   "frmMainForm.frx":09E4
      TabIndex        =   14
      Top             =   2880
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage9 
      Height          =   480
      Left            =   8280
      OleObjectBlob   =   "frmMainForm.frx":44FD
      TabIndex        =   13
      Top             =   2880
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage8 
      Height          =   480
      Left            =   8880
      OleObjectBlob   =   "frmMainForm.frx":107AB
      TabIndex        =   12
      Top             =   2280
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage7 
      Height          =   480
      Left            =   8280
      OleObjectBlob   =   "frmMainForm.frx":12168
      TabIndex        =   11
      Top             =   2280
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage5 
      Height          =   480
      Left            =   8280
      OleObjectBlob   =   "frmMainForm.frx":19D31
      TabIndex        =   10
      Top             =   1680
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage6 
      Height          =   480
      Left            =   8880
      OleObjectBlob   =   "frmMainForm.frx":1DD34
      TabIndex        =   9
      Top             =   1680
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorageOriginal 
      Height          =   480
      Left            =   8520
      OleObjectBlob   =   "frmMainForm.frx":210E7
      TabIndex        =   8
      Top             =   3840
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage4 
      Height          =   480
      Left            =   8880
      OleObjectBlob   =   "frmMainForm.frx":24162
      TabIndex        =   7
      Top             =   1080
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage3 
      Height          =   480
      Left            =   8280
      OleObjectBlob   =   "frmMainForm.frx":31D60
      TabIndex        =   6
      Top             =   1080
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage2 
      Height          =   480
      Left            =   8880
      OleObjectBlob   =   "frmMainForm.frx":3503D
      TabIndex        =   5
      Top             =   480
      Width           =   480
   End
   Begin ACTIVESKINLibCtl.SkinStorage SkinStorage1 
      Height          =   480
      Left            =   8280
      OleObjectBlob   =   "frmMainForm.frx":39C64
      TabIndex        =   4
      Top             =   480
      Width           =   480
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   38
      Width           =   4095
   End
   Begin MSComctlLib.StatusBar stasb1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtMain 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9975
      _Version        =   393217
      BackColor       =   12640511
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   1e13
      OLEDragMode     =   0
      OLEDropMode     =   1
      TextRTF         =   $"frmMainForm.frx":421E4
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVESKINLibCtl.SkinForm SkinForm1 
      Height          =   480
      Left            =   5160
      OleObjectBlob   =   "frmMainForm.frx":4229E
      TabIndex        =   1
      Top             =   3000
      Width           =   480
   End
   Begin VB.Menu mmnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   " &New Document"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   " &Open Document"
      End
      Begin VB.Menu mnuSave 
         Caption         =   " &Save Document"
      End
      Begin VB.Menu mnuDash7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   " Save Document &As"
      End
      Begin VB.Menu mnuPrinterSetup 
         Caption         =   " Pr&inter Setup"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   " &Print"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   " E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   " &Undo"
      End
      Begin VB.Menu mnuDash3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   " &Cut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   " &Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   " &Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   " &Clear All Text"
      End
      Begin VB.Menu mnuDash4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   " Select &All"
      End
      Begin VB.Menu mnuTimeDate 
         Caption         =   " Time/&Date"
      End
      Begin VB.Menu mnuDash5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWordCount 
         Caption         =   " &Word Count"
      End
      Begin VB.Menu mnuEncrypt 
         Caption         =   " Encr&ypt Message"
      End
      Begin VB.Menu mnuDecrypt 
         Caption         =   " Decr&ypt Message"
      End
      Begin VB.Menu mnuDash6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSetFont 
         Caption         =   " Set &Font"
      End
      Begin VB.Menu mnuColor 
         Caption         =   " Set Font Co&lor"
      End
      Begin VB.Menu mnuBGColor 
         Caption         =   " Set &Background Color"
      End
   End
   Begin VB.Menu mnuSkins 
      Caption         =   "&Skin Options"
      Begin VB.Menu mnuSkin1 
         Caption         =   " Skin &1"
      End
      Begin VB.Menu mnuSkin2 
         Caption         =   " Skin &2"
      End
      Begin VB.Menu mnuSkin3 
         Caption         =   " Skin &3"
      End
      Begin VB.Menu mnuSkin4 
         Caption         =   " Skin &4"
      End
      Begin VB.Menu mnuskin5 
         Caption         =   " Skin &5"
      End
      Begin VB.Menu mnuSkin6 
         Caption         =   " Skin &6"
      End
      Begin VB.Menu mnuSkin7 
         Caption         =   " Skin &7"
      End
      Begin VB.Menu mnuSkin8 
         Caption         =   " Skin &8"
      End
      Begin VB.Menu mnuSkin9 
         Caption         =   " Skin &9"
      End
      Begin VB.Menu mnuSkin10 
         Caption         =   " Skin &10"
      End
      Begin VB.Menu mnuDashSkins 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRestore 
         Caption         =   " &Restore Default Skin"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   " &About"
      End
   End
   Begin VB.Menu Popup 
      Caption         =   " Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuUndo1 
         Caption         =   " &Undo"
      End
      Begin VB.Menu mnuPopDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut1 
         Caption         =   " &Cut"
      End
      Begin VB.Menu mnuCopy1 
         Caption         =   " &Copy"
      End
      Begin VB.Menu mnuPaste1 
         Caption         =   " &Paste"
      End
      Begin VB.Menu mnuDelete1 
         Caption         =   " &Clear All Text"
      End
      Begin VB.Menu mnuDash2Pop 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll1 
         Caption         =   " Select &All"
      End
      Begin VB.Menu mnuPrint1 
         Caption         =   " &Print"
      End
   End
End
Attribute VB_Name = "frmMainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************
'******************************************
'******************************************
'****** Deceloped By Chris Palladino ******
'**** E-mail: midknightlover@home.com *****
'********** ICQ Number: 61141177 **********
'********** If You Use Anything ***********
'********* Please Give Me Credit **********
'******** If You Change Anything **********
'******* Please E-mail Me The New *********
'********* Version. Thank You!!! **********
'******************************************
'******************************************
'******************************************

Dim CharCount As Boolean
Option Explicit

' Color the tags in the RichTextBox's text.
' This version is a little simple and does not
' ignores comment properly. It cannot handle nested
' brackets as in:
'
' <A HREF= <!-- here's a comment -->
'    http://www.planetsourcecode.com>
'
Private Sub ColorTags(rch As RichTextBox)
Dim txt As String
Dim tag_open As Integer
Dim tag_close As Integer

    txt = rch.Text
    tag_close = 1
    Do
        
        tag_open = InStr(tag_close, txt, "<") ' See where the next tag starts.
        If tag_open = 0 Then Exit Do
        
        
        tag_close = InStr(tag_open, txt, ">") ' See where the tag ends.
        If tag_open = 0 Then tag_close = Len(txt)
        
        
        rch.SelStart = tag_open - 1 ' Color the tag.
        rch.SelLength = tag_close - tag_open + 1
        rch.SelColor = vbRed
        rch.SelUnderline = True
        rch.SelBold = True
        
    Loop
End Sub

Private Sub cmdFind_Click()
Dim textfound As Integer

    
cmdFindNext.Enabled = True ' Enables the cmdFindNext

                          ' Finds the text in the search box and highlights it,
                          ' then sets the focus on the richtextbox so the selected
                          ' text is editable.
txtMain.Find (Text1.Text)
txtMain.SetFocus

    ' The richtextbox1.find method returns an integer
    ' value of -1 if the searched for text is not found.
    ' If this is true then it displays a message box.
textfound = txtMain.Find(Text1.Text)
If textfound = -1 Then
MsgBox "End of Document" & vbCr & "Text Not Found", vbInformation, _
    App.Title
End If
  
  
End Sub


Private Sub cmdFindNext_Click()

    ' Set the focus so the selected text can be directly edited.
txtMain.SetFocus

    ' Finds the next instance of the word, starting from
    ' the selected text.
txtMain.Find (Text1.Text), txtMain.SelStart + 1
End Sub

Private Sub Form_Load()
cmdFind.ApplySkin SkinForm1
cmdFindNext.ApplySkin SkinForm1
End Sub

Private Sub Form_Resize()
  txtMain.Width = frmMainForm.ScaleWidth
  txtMain.Height = frmMainForm.ScaleHeight
End Sub

Private Sub mnuAbout_Click()
   frmAbout.Show
End Sub

Private Sub mnuBGColor_Click()
    On Error GoTo ErrorHandler 'Set cancel to true
    CommonDialog1.ShowColor
    txtMain.BackColor = CommonDialog1.Color
ErrorHandler:     ' User Pressed The Cancel Button
 Exit Sub
End Sub

Private Sub mnuColor_Click()
    On Error GoTo ErrorHandler 'Set cancel to true
    CommonDialog1.ShowColor
    txtMain.SelColor = CommonDialog1.Color
ErrorHandler: ' User Pressed The Cancel Button
 Exit Sub
End Sub

Private Sub mnuCopy_Click()
    Clipboard.Clear 'Clear the clipboard
    Clipboard.SetText Screen.ActiveControl.SelText 'Get and store text from the active control on our form
End Sub

Private Sub mnuCopy1_Click()
    Clipboard.Clear 'Clear the clipboard
    Clipboard.SetText Screen.ActiveControl.SelText 'Get and store text from the active control on our form
End Sub

Private Sub mnuCut_Click()
   Clipboard.Clear 'Clear The Clipboard
   Clipboard.SetText Screen.ActiveControl.SelText 'Get and store text from the active control on our form
   Screen.ActiveControl.SelText = ""
End Sub

Private Sub mnuCut1_Click()
    Clipboard.Clear 'Clear The Clipboard
    Clipboard.SetText Screen.ActiveControl.SelText 'Get and store text from the active control on our form
    Screen.ActiveControl.SelText = ""
End Sub

Private Sub mnuDecrypt_Click()
    Dim AsciiOf As Integer
    Dim NewText As String
    Dim OldText As String
    Dim x As Long
        OldText = txtMain.Text
        frmMainForm.Caption = "Skin Pad Plus - DeEncrypting..."
        txtMain.Text = "DeEncrypting..."

    For x = 1 To Len(OldText)
        DoEvents
        AsciiOf = Asc(Mid(OldText, x, 1))
        If AsciiOf <= 25 Then AsciiOf = AsciiOf + 255
        NewText = NewText & Chr(AsciiOf - 25)
    Next

        txtMain.Text = NewText
        frmMainForm.Caption = "Skin Pad Plus"
    Call Chars_Lines
End Sub

Private Sub mnuDelete_Click()
    txtMain.Text = ""
End Sub

Private Sub mnuDelete1_Click()
    txtMain.Text = ""
End Sub

Private Sub mnuEncrypt_Click()
    Dim Letter1 As String
    Dim AsciiOf As Integer
    Dim NewText As String
    Dim MemText As String
    Dim x As Long
        MemText = txtMain.Text
        frmMainForm.Caption = "Skin Pad Plus - Encrypting...."
        txtMain.Text = "Encrypting..."
        For x = 1 To Len(MemText)
    DoEvents
    Letter1 = Mid(MemText, x, 1)
    AsciiOf = Asc(Letter1)
    AsciiOf = AsciiOf + 25
        If AsciiOf > 255 Then AsciiOf = AsciiOf - 255
        NewText = NewText & Chr(AsciiOf)
    Next
        txtMain.Text = NewText
        frmMainForm.Caption = "Skin Pad Plus"
    Call Chars_Lines
End Sub

Private Sub mnuExit_Click()
    Dim a As Integer
        If txtMain.Text <> "" Then
        a = MsgBox("Would you like to save before exiting?", vbYesNoCancel, "Save?")
        If a = vbYes Then
    Call mnuSaveAs_Click
  End If
        If a = vbCancel Then
Exit Sub
  End If
  End If
  Unload Me
  End
End Sub

Private Sub mnuFindNext_Click()
txtMain.SetFocus ' Set the focus so the selected text can be directly edited.

    
txtMain.Find (txtMain.Text), txtMain.SelStart + 1 ' Finds the next instance of the word, starting from
                                                  ' the selected text.
End Sub

Private Sub mnuNew_Click()
   txtMain.Text = ""
   
End Sub

Private Sub mnuOpen_Click()
Dim strOpenFile As String
On Error GoTo ErrorHandler 'Set cancel to true
   CommonDialog1.CancelError = True 'Set cancel to true
   CommonDialog1.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly ' Set flags
   CommonDialog1.Filter = "Text Documents (*.txt)|*.txt|Rich Text Files (*.rtf)|*.rtf|HTML Files (*.htm;*.html)|*.htm;*.html|NFO Files (*.nfo)|*.nfo|DIZ Files (*.diz)|*.diz|All Files (*.*)|*.*" 'Set filters
   CommonDialog1.FilterIndex = 6 'Specify default filter
   CommonDialog1.ShowOpen
   strOpenFile = CommonDialog1.FileName
   txtMain.LoadFile strOpenFile
   
   ColorTags txtMain
   
ErrorHandler: ' User Pressed The Cancel Button
 Exit Sub
End Sub

Private Sub mnuPaste_Click()
    Screen.ActiveControl.SelText = Clipboard.GetText 'Set the currently selected text in the active                                                    'control to the text on the clipboard
End Sub

Private Sub mnuPaste1_Click()
    Screen.ActiveControl.SelText = Clipboard.GetText 'Set the currently selected text in the active                                                    'control to the text on the clipboard
End Sub

Private Sub mnuPrint_Click()
    Printer.Print txtMain.Text
End Sub

Private Sub mnuPrint1_Click()
    Printer.Print txtMain.Text
End Sub

Private Sub mnuPrinterSetup_Click()
    On Error GoTo ErrorHandler 'Set cancel to true
    CommonDialog1.ShowPrinter
ErrorHandler: ' User Pressed The Cancel Button
 Exit Sub
End Sub


Private Sub mnuRestore_Click()
    ApplySkin SkinStorageOriginal.SkinSource
    
End Sub

Private Sub mnuSave_Click()
    Dim strNewFile As String
    On Error GoTo ErrorHandler 'Set cancel to true
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt 'Set flags
        CommonDialog1.Filter = "Text Documents (*.txt)|*.txt" 'Set filters
        CommonDialog1.FilterIndex = 1 'Specify default filter
        CommonDialog1.ShowSave
        strNewFile = CommonDialog1.FileName
        txtMain.SaveFile strNewFile
ErrorHandler:     'User pressed the cancel button
  Exit Sub
End Sub

Private Sub mnuSaveAs_Click()
    Dim strNewFile As String
    On Error GoTo ErrorHandler 'Set cancel to true
        CommonDialog1.CancelError = True
        CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt 'Set flags
        CommonDialog1.Filter = "All Files (*.*)|*.*|Rich Text Files (*.rtf)|*.rtf|Text Documents (*.txt)|*.txt" 'Set filters
        CommonDialog1.FilterIndex = 1 'Specify default filter
        CommonDialog1.ShowSave
        strNewFile = CommonDialog1.FileName
        txtMain.SaveFile strNewFile
ErrorHandler:   'User pressed the cancel button
  Exit Sub
End Sub

Private Sub mnuSelectAll_Click()
   Dim a As String
       txtMain.SelStart = 0

       txtMain.SelLength = Len(txtMain.Text)
End Sub

Private Sub mnuSelectAll1_Click()
    Dim a As String
       txtMain.SelStart = 0

       txtMain.SelLength = Len(txtMain.Text)
End Sub

Private Sub mnuSetFont_Click()
    On Error GoTo ErrorHandler 'Set cancel to true
        CommonDialog1.Flags = cdlCFScreenFonts
        CommonDialog1.ShowFont
        txtMain.SelFontName = CommonDialog1.FontName
        txtMain.SelBold = CommonDialog1.FontBold
        txtMain.SelItalic = CommonDialog1.FontItalic
        txtMain.SelFontSize = CommonDialog1.FontSize
        txtMain.SelStrikeThru = CommonDialog1.FontStrikethru
        txtMain.SelUnderline = CommonDialog1.FontUnderline
ErrorHandler:      'User pressed the cancel button
Exit Sub
End Sub

Private Sub mnuSkin1_Click()
    ApplySkin SkinStorage1.SkinSource
    
End Sub

Private Sub mnuSkin10_Click()
    ApplySkin SkinStorage10.SkinSource
    
End Sub

Private Sub mnuSkin2_Click()
    ApplySkin SkinStorage2.SkinSource
    
End Sub

Private Sub mnuSkin3_Click()
    ApplySkin SkinStorage3.SkinSource
    
End Sub

Private Sub mnuSkin4_Click()
    ApplySkin SkinStorage4.SkinSource
    
End Sub

Private Sub mnuSkin5_Click()
    ApplySkin SkinStorage5.SkinSource
    
End Sub

Private Sub mnuSkin6_Click()
    ApplySkin SkinStorage6.SkinSource
    
End Sub

Private Sub mnuSkin7_Click()
    ApplySkin SkinStorage7.SkinSource
    
End Sub

Private Sub mnuSkin8_Click()
    ApplySkin SkinStorage8.SkinSource
    
End Sub

Private Sub mnuSkin9_Click()
    ApplySkin SkinStorage9.SkinSource
    
End Sub

Private Sub mnuTimeDate_Click()
   SendKeys (Now)
End Sub

Private Sub mnuUndo_Click()
    SendKeys ("^z")
End Sub

Private Sub mnuWordCount_Click()
    Dim a() As String
    Dim b() As String
    Dim wordcount As Long
    Dim x As Long
        frmMainForm.Caption = "Skin Pad Plus - Counting Words..."

        a() = Split(txtMain.Text, " ")
        wordcount = UBound(a)
        For x = 0 To UBound(a)
        If a(x) = "" Then
        wordcount = wordcount - 1
    End If
  Next

        b() = Split(txtMain.Text, Chr$(10))
        wordcount = wordcount + UBound(b)
        For x = 0 To UBound(b)
        If b(x) = "" Then
        wordcount = wordcount - 1
    End If
  Next
        If wordcount = -2 Then wordcount = -1
        frmMainForm.Caption = "Skin Pad Plus"
        MsgBox "Words In Text Area: " & wordcount + 1, vbOKOnly, "Skin Pad Plus Word Counter"

End Sub

Private Sub Chars_Lines()
    If CharCount = True Then
    Dim Lines, Chars As String
    Dim blah() As String
    Dim bleh() As String
    Dim Curline As String
    Dim CurChar, TotalChar As String

        Curline = Mid(txtMain.Text, 1, txtMain.SelStart)
        blah() = Split(Curline, Chr$(10))
        bleh() = Split(txtMain.Text, Chr$(10))

       If txtMain.SelStart = 0 Then
       CurChar = 0
       Curline = 1

       If Len(txtMain.Text) = 0 Then
       TotalChar = 0
    Else
       TotalChar = Len(txtMain.Text) - (UBound(bleh) * 2)

  End If

    Else
        CurChar = txtMain.SelStart - (UBound(blah) * 2)
        Curline = UBound(blah) + 1
        TotalChar = Len(txtMain.Text) - (UBound(bleh) * 2)
  End If



        Lines = "Line:" & Curline & "/" & SendMessage(txtMain.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
        Chars = "Char:" & CurChar & "/" & TotalChar

        stasb1.SimpleText = Chars & "  " & Lines

    Else
        stasb1.SimpleText = "Off"
End If
End Sub

Private Sub SkinButton1_Click()

    ' Set the focus so the selected text can be directly edited.
txtMain.SetFocus

    ' Finds the next instance of the word, starting from
    ' the selected text.
txtMain.Find (Text1.Text), txtMain.SelStart + 1
End Sub

Private Sub Text1_Change()

    ' Enables the find button when the search text is entered.
cmdFind.Enabled = True

End Sub

Private Sub ApplySkin(SkSrc As SkinSource)
    Set SkinForm1.SkinSource = SkSrc
    Text1.BackColor = SkSrc.ClrWindow
    Text1.ForeColor = SkSrc.ClrWindowText
    Text1.BackColor = SkSrc.ClrWindow
    Text1.ForeColor = SkSrc.ClrWindowText
    cmdFind.ApplySkin SkinForm1
    cmdFindNext.ApplySkin SkinForm1
    Refresh
End Sub

Private Sub txtMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = vbRightButton Then
     Beep
     PopupMenu Popup
    End If
End Sub
