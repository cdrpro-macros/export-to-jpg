VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} exportToJpeg_Form 
   Caption         =   "expJpg"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   3600
   OleObjectBlob   =   "exportToJpeg_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "exportToJpeg_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myWidth&, myBatExe$, sAspectRat$
Public wBool As Boolean, hBool As Boolean

Private sanUI&
Private macroAdURL$, IsOwer As Boolean
Private r0 As wRECT, cp0 As wPOINT, sx&, sy&

#If VBA7 Then
  Private Declare PtrSafe Function FindWindowExW& Lib "user32" (ByVal hParent&, ByVal hChildAfter&, ByVal lpClassW&, ByVal lpTitleW&)
  Private Declare PtrSafe Function SendMessageW& Lib "user32" (ByVal hwnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
  Private Declare PtrSafe Function GetSystemMetrics& Lib "user32" (ByVal Index As wEnumSystemMetrics)
  Private Declare PtrSafe Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags As wEnumSetWindowPos)
  Private Declare PtrSafe Function GetWindowRect& Lib "user32" (ByVal hwnd&, r As wRECT)
  Private Declare PtrSafe Function GetCursorPos& Lib "user32" (p As wPOINT)
  Private Declare PtrSafe Function SetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%, ByVal dwNewLong&)
  Private Declare PtrSafe Function GetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%)
  Private Declare PtrSafe Function SetLayeredWindowAttributes& Lib "user32" (ByVal hwnd&, ByVal crKey%, ByVal bAlpha As Byte, ByVal dwFlags&)
  Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  Private Declare PtrSafe Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  Private Declare PtrSafe Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
  Private Declare PtrSafe Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
#Else
  Private Declare Function FindWindowExW& Lib "user32" (ByVal hParent&, ByVal hChildAfter&, ByVal lpClassW&, ByVal lpTitleW&)
  Private Declare Function SendMessageW& Lib "user32" (ByVal hwnd&, ByVal msg&, ByVal wParam&, ByVal lParam&)
  Private Declare Function GetSystemMetrics& Lib "user32" (ByVal Index As wEnumSystemMetrics)
  Private Declare Function SetWindowPos& Lib "user32" (ByVal hwnd&, ByVal hAfter&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal wFlags As wEnumSetWindowPos)
  Private Declare Function GetWindowRect& Lib "user32" (ByVal hwnd&, r As wRECT)
  Private Declare Function GetCursorPos& Lib "user32" (p As wPOINT)
  Private Declare Function SetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%, ByVal dwNewLong&)
  Private Declare Function GetWindowLongW& Lib "user32" (ByVal hwnd&, ByVal nIndex%)
  Private Declare Function SetLayeredWindowAttributes& Lib "user32" (ByVal hwnd&, ByVal crKey%, ByVal bAlpha As Byte, ByVal dwFlags&)
  Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
  Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
  Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
  Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
  Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
#End If



Private Enum wEnumSystemMetrics: SM_CXScreen = 0: SM_CYScreen = 1: End Enum
Private Enum wEnumSetWindowPos: SWP_NOSIZE = 1: SWP_NOMOVE = 2: SWP_NOZORDER = 4: SWP_NOREDRAW = 8: SWP_NOACTIVATE = 16: _
              SWP_DRAWFRAME = 32: SWP_SHOWWINDOW = 64: SWP_HIDEWINDOW = 128: SWP_NOCOPYBITS = 256: SWP_NOREPOSITION = 512: _
              SWP_NOSENDCHANGING = 1024: SWP_DEFERERASE = 8192: SWP_ASYNCWINDOWPOS = 16384: End Enum
Private Type wPOINT: X As Long: Y As Long: End Type
Private Type wRECT: l As Long: t As Long: r As Long: b As Long: End Type

Const REG_SZ = 1 ' Unicode nul terminated string
Const REG_BINARY = 3 ' Free form binary

Private Sub cm_close_Click()
  Unload Me
End Sub

Private Sub Command_OK_Click()
  Dim d As Document, opt As New StructExportOptions, p$, p2$
  Dim myColorType&, sn$, myPage As Page, oldPage As Page, myFold$, c&, c2&
  Dim myStrPage$
  
  Me.meHead.Caption = Me.meHead.Caption & " (%)"
  IsOwer = False
  
  Set d = ActiveDocument: c2 = 1
  Set oldPage = ActivePage
  
  If myWorkMetod = "Selection" Then _
      If ActiveShape Is Nothing Then MsgBox "Not selected!  ", vbCritical, "Exp2JPG": Exit Sub
  
  If d.Pages.Count = 1 And myWorkMetod = "All Page" Then myWorkMetod = "Active Page": Exit Sub
  
  '=============================================================================
  sn = nameprefix_TextBox.Text
  If myWorkMetod = "Selection" Then sn = sn & "_" & tb_AfterName.Text
  'If myWorkMetod <> "All Page" Then sn = tb_AfterName.Text Else sn = ""
  'If prefix_CheckBox.Value = True Then  & sn
  '=============================================================================
  
  If myCmpress.Value < 0 Or myCmpress.Value > 100 Then MsgBox "Wrong values are set!": Exit Sub
  If Smoot.Value < 0 Or Smoot.Value > 100 Then MsgBox "Wrong values are set!": Exit Sub
  
  Select Case myColorMode.ListIndex
      Case 0: myColorType = 2
      Case 1: myColorType = 5
      Case 2: myColorType = 4
  End Select
  

  With opt
      .AntiAliasingType = AntiAliasing_CheckBox.Value
      '.Compression = cdrCompressionJPEG
      .ImageType = myColorType
      .MaintainAspect = CLng(sAspectRat)
      .MaintainLayers = False
      .ResolutionX = CLng(Dpi_TextBox.Text)
      .ResolutionY = CLng(Dpi_TextBox.Text)
      .SizeX = CLng(bwidth.Text)
      .SizeY = CLng(bheight.Text)
      .Transparent = False
      .UseColorProfile = profile_CheckBox.Value
      .Overwrite = False
  End With
  
  
  myFold = folder_TextBox.Text
  If Right$(myFold, 1) <> "\" Then myFold = myFold & "\"
  
  '=============================================================================
  Select Case myWorkMetod
  
  Case "Selection"
      p = myFold & sn & ".jpg"
      If exportToJpeg2(p, opt) Then
        If mUse = True Then mySendMail p
      End If
  
  Case "Active Page"
      If ActivePage.Shapes.Count = 0 Then MsgBox "Shapes Count = 0": Exit Sub
      p = myFold & sn & "_" & ActivePage.Index & ".jpg"
      If exportToJpeg2(p, opt) Then
        If mUse = True Then mySendMail p
      End If
  
  Case "All Page"
      If d.Pages.Count < 2 Then Exit Sub
      For c = 1 To d.Pages.Count
        If d.Pages(c).Shapes.Count > 0 Then
          d.Pages(c).Activate
          p = myFold & sn & "_" & c2 & ".jpg": c2 = c2 + 1
          If exportToJpeg2(p, opt) Then
            If mUse = True And myMultSend = False Then
              If c = 1 Then p2 = p
              If c > 1 Then p2 = p2 & "|" & p
            End If
            If mUse = True And myMultSend = True Then mySendMail p
          End If
        End If
      Next c
      If mUse = True And myMultSend = False Then mySendMail p2
      
  Case "Page:"
      myStrPage = getNumPage(myPageIndex.Text)
      
      For Each a In Split(myStrPage, ",")
          c = CLng(a)
          If d.Pages(c).Shapes.Count > 0 Then
            d.Pages(c).Activate
            p = myFold & sn & "_" & c & ".jpg"
            If exportToJpeg2(p, opt) Then
              If mUse = True And myMultSend = False Then
                If c = 1 Then p2 = p
                If c > 1 Then p2 = p2 & "|" & p
              End If
              If mUse = True And myMultSend = True Then mySendMail p
            End If
          End If
      Next a
      If mUse = True And myMultSend = False Then mySendMail p2
      
  End Select
  '=============================================================================

  oldPage.Activate
  myListsSave
  Me.meHead.Caption = Left$(Me.meHead.Caption, Len(Me.meHead.Caption) - 4)
End Sub

Private Function exportToJpeg2(p$, opt As StructExportOptions) As Boolean
  Dim ex As ExportFilter, d As Document, myMsg&
  Set d = ActiveDocument
  
  If FileSystem.Dir(p) <> "" And IsOwer = False Then
    myMsg = MsgBox("File already exists" & vbCr & "Overwrite (all)?", vbOKCancel + vbExclamation, "Warning...")
    If myMsg = 1 Then
      opt.Overwrite = True
      IsOwer = True
    Else
      exportToJpeg2 = False
      Exit Function
    End If
  End If
  
  If myWorkMetod = "Selection" Then
    Set ex = d.ExportEx(p, cdrJPEG, cdrSelection, opt)
  Else
    Set ex = d.ExportEx(p, cdrJPEG, cdrCurrentPage, opt)
  End If
  
  With ex
    If jpgEncod.Text = "Progressive" Then
      .Progressive = True
      .Optimized = False
    ElseIf jpgEncod.Text = "Optimized" Then
      .Progressive = False
      .Optimized = True
    End If
    .SubFormat = 0
    .Compression = CLng(myCmpress.Text)
    .Smoothing = CLng(Smoot.Text)
    .Finish
  End With
  
  exportToJpeg2 = True
End Function

Private Function getNumPage$(myPageIn$)
    Dim a$(), a2$(), i&, i2&
    getNumPage = ""
    a = Split(myPageIn, ",", , vbTextCompare)
    For i = 0 To UBound(a)
        a2 = Split(a(i), "-", , vbTextCompare)
        If UBound(a2) > 0 Then
            If a2(UBound(a2)) > a2(0) Then
                For i2 = CLng(a2(0)) To CLng(a2(UBound(a2)))
                    If getNumPage = "" Then getNumPage = i2 _
                    Else getNumPage = getNumPage & "," & i2
                Next i2
            Else
                MsgBox "incorrect number order!   ", vbCritical, "Exp2JPG"
            End If
        Else
            If getNumPage = "" Then getNumPage = CLng(a(i)) _
            Else getNumPage = getNumPage & "," & CLng(a(i))
        End If
    Next
End Function

Private Sub mySendMail(txtAttachementFileLocation As String)
  Dim txtMainAddresses As String, txtSubject As String, txtUser$
  
  s = txtAttachementFileLocation
  s = Replace(s, "|", Chr(34) & ";ATTACH=" & Chr(34))
  txtAttachementFileLocation = s
  
  '"c:\thebat\thebat.exe" /MAIL;TO=other@mail.com;USER="Название ящика";SUBJECT="Hello";ATTACH="c:\thebat\tips.ini"
  
  Dim sText As String
  Dim sAddedText As String
  
  If myBatExe = "" Then
  MsgBox "TheBat not found", vbCritical
  Exit Sub
  End If
  
  txtMainAddresses = mAddres.Text
  txtSubject = mSubject.Text
  txtUser = myAccount.Text
  
  If Len(txtMainAddresses) Then
    sText = Chr(34) & txtMainAddresses & Chr(34)
  End If
  
  If Len(txtUser) Then _
    sAddedText = sAddedText & ";USER=" & Chr(34) & txtUser & Chr(34)
  If Len(txtSubject) Then _
    sAddedText = sAddedText & ";SUBJECT=" & Chr(34) & txtSubject & Chr(34)
  If Len(txtAttachementFileLocation) Then _
    sAddedText = sAddedText & ";ATTACH=" & Chr(34) & _
    txtAttachementFileLocation & Chr(34)
  
  sText = Chr(34) & myBatExe & Chr(34) & " /MAIL;TO=" & sText
  sText = sText & sAddedText
  
  If Len(sText) Then
    ReturnValue = Shell(sText, vbMaximizedFocus)
  End If
End Sub

Private Sub Command_OK_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Command_OK.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub Command_OK_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Command_OK.BackColor = vbWhite
  Command_OK.SpecialEffect = fmSpecialEffectRaised
End Sub

Private Sub folder_TextBox_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "folder2", folder_TextBox
  folder_TextBox.ControlTipText = "Folder: " & folder_TextBox.Text
End Sub

Private Sub folder_TextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  folder_TextBox.Text = ActiveDocument.FilePath
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "folder2", folder_TextBox
End Sub

Private Sub mAddres_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "mAddres2", mAddres
End Sub

Private Sub mSubject_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "mSubject2", mSubject
End Sub

Private Sub myPageIndex_Change()
  s = myPageIndex.Text
  s = Replace(s, ".", ",")
  myPageIndex.Text = s
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myPageIndex", myPageIndex
  myPageIndex.ControlTipText = "Pages: " & myPageIndex.Text
End Sub

Private Sub tb_AfterName_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "name", tb_AfterName
End Sub


Private Sub Command_folder_Click()
  Dim myFolder$
  myFolder = Application.CorelScriptTools.GetFolder(folder_TextBox.Text, "Export to...")
  If myFolder = "" Then Exit Sub
  folder_TextBox.Text = myFolder
End Sub


Private Sub Dpi_TextBox_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "dpi", Dpi_TextBox.Text
  If Dpi_TextBox.Value < 0 Or Dpi_TextBox.Value > 600 Then
    Dpi_TextBox.ForeColor = vbRed
  Else
    Dpi_TextBox.ForeColor = &H80000012
  End If
  If bwidth.Text = "0" Then myColculateSizeNull
End Sub


Private Sub nameprefix_TextBox_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "name prefix text", nameprefix_TextBox
End Sub

Private Sub nameprefix_TextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
  nameprefix_TextBox.Text = Replace(CorelDRAW.ActiveDocument.Name, ".cdr", "")
End Sub

'Private Sub prefix_CheckBox_Click()
'  SaveSetting "SanchoCorelVBA", "ExportToJpg", "name prefix", IIf(prefix_CheckBox, "1", "0")
'  namePrefixEn
'End Sub

'Private Sub namePrefixEn()
'  If prefix_CheckBox = False Then
'    nameprefix_TextBox.Enabled = False
'  Else
'    nameprefix_TextBox.Enabled = True
'  End If
'End Sub

Private Sub myColorMode_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "ColorMode", myColorMode
End Sub

Private Sub bwidth_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "width", bwidth
  If bwidth.Value < 0 Or bwidth.Value > 10000 Then
  bwidth.ForeColor = vbRed
  Else
  bwidth.ForeColor = &H80000012
  End If
  If sAspectRat = "1" Then _
      If wBool Then If bwidth.Text <> "" Then myColculateSize "w"
  If bwidth.Text = "0" Then myColculateSizeNull Else bwidth.ControlTipText = ""
End Sub

Private Sub bwidth_Enter()
  wBool = True
End Sub

Private Sub bwidth_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  wBool = False
End Sub
            
Private Sub bheight_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "height", bheight
  If bheight.Value < 0 Or bheight.Value > 10000 Then
  bheight.ForeColor = vbRed
  Else
  bheight.ForeColor = &H80000012
  End If
  If sAspectRat = "1" Then _
      If hBool Then If bheight.Text <> "" Then myColculateSize "h"
  If bheight.Text = "0" Then myColculateSizeNull Else bheight.ControlTipText = ""
End Sub

Private Sub bheight_Enter()
  hBool = True
End Sub

Private Sub bheight_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  hBool = False
End Sub

Private Sub myColculateSize(st$)
  Dim sr As ShapeRange, w#, h#, d As Document, oldUn As cdrUnit, myPes#
  Set sr = New ShapeRange
  Set d = ActiveDocument
  oldUn = d.Unit: d.Unit = cdrInch
  
  Select Case myWorkMetod
  Case "Selection"
      Set sr = ActiveSelectionRange
      If sr.Count = 0 Then Exit Sub
      sr.GetSize w, h
      
      w = w * Dpi_TextBox.Value
      h = h * Dpi_TextBox.Value
      If st = "w" Then
          myPes = bwidth.Value * 100 / w
          bheight.Value = CLng((myPes * h) / 100)
      ElseIf st = "h" Then
          myPes = bheight.Value * 100 / h
          bwidth.Value = CLng((myPes * w) / 100)
      End If

  'Case "Active Page"
  'Case "All Page"
  'Case "Page:"
  End Select
  d.Unit = oldUn
End Sub

Private Sub myColculateSizeNull()
  Dim sr As ShapeRange, w#, h#, d As Document, oldUn As cdrUnit
  Set sr = New ShapeRange
  Set d = ActiveDocument
  oldUn = d.Unit: d.Unit = cdrInch
  Set sr = ActiveSelectionRange
  If sr.Count = 0 Then Exit Sub
  sr.GetSize w, h
  bwidth.ControlTipText = CLng(w * Dpi_TextBox.Value)
  bheight.ControlTipText = CLng(h * Dpi_TextBox.Value)
  d.Unit = oldUn
End Sub

Private Sub cm_AspectRat_dis_Click()
  cm_AspectRat_dis.Move 137, 79
  cm_AspectRat_en.Move 90, 79
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "sAspectRat", "0"
  sAspectRat = "0"
End Sub

Private Sub cm_AspectRat_en_Click()
  cm_AspectRat_en.Move 137, 79
  cm_AspectRat_dis.Move 90, 79
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "sAspectRat", "1"
  sAspectRat = "1"
  If bwidth.Text <> "" Then myColculateSize "w"
End Sub

Private Sub cm_AspectRatMove()
  If sAspectRat = "1" Then
      cm_AspectRat_en.Move 137, 79
      cm_AspectRat_dis.Move 90, 79
  Else
      cm_AspectRat_dis.Move 137, 79
      cm_AspectRat_en.Move 90, 79
  End If
End Sub

Private Sub profile_CheckBox_Click()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "profile", IIf(profile_CheckBox, "1", "0")
End Sub

Private Sub AntiAliasing_CheckBox_Click()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "AntiAliasing", IIf(AntiAliasing_CheckBox, "1", "0")
End Sub

Private Sub jpgEncod_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "EncodingMethod", jpgEncod
End Sub

Private Sub myCmpress_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myCmpress", myCmpress
  
  If myCmpress.Value < 0 Or myCmpress.Value > 100 Then
      myCmpress.ForeColor = vbRed
  Else
      myCmpress.ForeColor = &H80000012
  End If
End Sub

Private Sub Smoot_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "Smoot", Smoot
  
  If Smoot.Value < 0 Or Smoot.Value > 100 Then
      Smoot.ForeColor = vbRed
  Else
      Smoot.ForeColor = &H80000012
  End If
End Sub

Private Sub myWorkMetod_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myWorkMetod", myWorkMetod
  myWorkMetod2
End Sub

Private Sub myWorkMetod2()
  If myWorkMetod.Text = "All Page" Then
      If mUse = 0 Then
      myMultSend.Enabled = False
      Else
      myMultSend.Enabled = True
      End If
      myPageIndex.Enabled = False
      myPageIndex.BackColor = &H8000000F
  ElseIf myWorkMetod.Text = "Page:" Then
      If mUse = 0 Then
      myMultSend.Enabled = False
      Else
      myMultSend.Enabled = True
      End If
      myPageIndex.Enabled = True
      myPageIndex.BackColor = &H80000005
  ElseIf myWorkMetod.Text = "Selection" Then
      myMultSend.Enabled = False
      myPageIndex.Enabled = False
      myPageIndex.BackColor = &H8000000F
      If bwidth.Text <> "" Then myColculateSize "w"
      If bwidth.Text = "0" Then myColculateSizeNull
  ElseIf myWorkMetod.Text = "Active Page" Then
      myMultSend.Enabled = False
      myPageIndex.Enabled = False
      myPageIndex.BackColor = &H8000000F
  End If
End Sub

Private Sub mTheBatEXE_Click()
  Dim myFolder$
  myTheBat_exe = Application.CorelScriptTools.GetFileBox("TheBat (*.exe)|*.exe", "TheBat.exe...", 1, , , , "OK")
  If myTheBat_exe = "" Then Exit Sub
  myBatExe = myTheBat_exe
  
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myBatExe", myBatExe
End Sub

Private Sub mUse_Click()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "mUse", IIf(mUse, "1", "0")
  
  If mUse = 0 Then
      mAddres.Enabled = False
      mSubject.Enabled = False
      mTheBatEXE.Enabled = False
      myAccount.Enabled = False
      myMultSend.Enabled = False
  Else
      mAddres.Enabled = True
      mSubject.Enabled = True
      mTheBatEXE.Enabled = True
      myAccount.Enabled = True
      If myWorkMetod.Text = "All Page" Or myWorkMetod.Text = "Page:" Then myMultSend.Enabled = True
  End If
End Sub


Private Sub myMultSend_Click()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myMultSend", IIf(myMultSend, "1", "0")
End Sub

Private Sub myAccount_Change()
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "myAccount", myAccount
End Sub

Private Sub UserForm_Initialize()
  Dim meCap$, s$, lng1&, lng2&
  On Error Resume Next
  Me.DesignMode = fmModeOff
      meCap = Me.Caption
      Me.Caption = Me.Caption & Hex$(Now) & Hex$(Timer)
      lng1 = CLng(StrPtr("ThunderDFrame"))
      lng2 = CLng(StrPtr(Me.Caption))
      sanUI = FindWindowExW(0, 0, lng1, lng2)
      Me.Caption = meCap
  windowSetStyle
  Me.Height = 245: Me.Width = 122.25
  windowTransparent 220

  Me.meHead.Caption = "  Exp2jpg " & jpgVer
  Me.meHead.ControlTipText = "Exp2jpg " & jpgVer & " " & jpgDate & "  " & Chr(169) & " Sancho"
  s = GetSetting("SanchoCorelVBA", "ExportToJpg", "Pos")
  If Len(s) Then
    StartUpPosition = 0
    'Move CSng(Split(s, " ")(0)), CSng(Split(s, " ")(1))
    Me.Left = CSng(Split(s, " ")(0))
    Me.Top = CSng(Split(s, " ")(1))
  End If
  
  'namePrefixEn
  
  wBool = False
  hBool = False
  
  myWorkMetod.AddItem "Selection"
  myWorkMetod.AddItem "Active Page"
  myWorkMetod.AddItem "All Page"
  myWorkMetod.AddItem "Page:"
  myWorkMetod = GetSetting("SanchoCorelVBA", "ExportToJpg", "myWorkMetod", "Selection")
  myWorkMetod2
  
  Dpi_TextBox.AddItem "300"
  Dpi_TextBox.AddItem "200"
  Dpi_TextBox.AddItem "150"
  Dpi_TextBox.AddItem "100"
  Dpi_TextBox.AddItem "96"
  Dpi_TextBox.AddItem "72"
  Dpi_TextBox = GetSetting("SanchoCorelVBA", "ExportToJpg", "dpi", "100")
  tb_AfterName = GetSetting("SanchoCorelVBA", "ExportToJpg", "name", "")

  folder_TextBox.List = Split(GetSetting("SanchoCorelVBA", "ExportToJpg", "folder"), "|", 20)
  If folder_TextBox.ListCount = 20 Then folder_TextBox.RemoveItem 19
  If folder_TextBox.ListCount = 0 Then
      folder_TextBox.Text = "C:\"
  Else
      folder_TextBox.Text = GetSetting("SanchoCorelVBA", "ExportToJpg", "folder2")
  End If

  'prefix_CheckBox = (GetSetting("SanchoCorelVBA", "ExportToJpg", "name prefix", "0") = "1")
  nameprefix_TextBox = GetSetting("SanchoCorelVBA", "ExportToJpg", "name prefix text", "maket_")
  myColorMode.List = Array("Grayscale", "CMYKColor", "RGBColor")
  myColorMode = GetSetting("SanchoCorelVBA", "ExportToJpg", "ColorMode", "RGBColor")
  bwidth = GetSetting("SanchoCorelVBA", "ExportToJpg", "width", "0")
  bheight = GetSetting("SanchoCorelVBA", "ExportToJpg", "height", "0")
  sAspectRat = GetSetting("SanchoCorelVBA", "ExportToJpg", "sAspectRat", "1")
      cm_AspectRatMove
  profile_CheckBox = (GetSetting("SanchoCorelVBA", "ExportToJpg", "profile", "1") = "1")
  AntiAliasing_CheckBox = (GetSetting("SanchoCorelVBA", "ExportToJpg", "AntiAliasing", "1") = "1")
  jpgEncod.List = Array("Progressive", "Optimized")
  jpgEncod = GetSetting("SanchoCorelVBA", "ExportToJpg", "EncodingMethod", "Optimized")
  myCmpress = GetSetting("SanchoCorelVBA", "ExportToJpg", "myCmpress", "50")
  Smoot = GetSetting("SanchoCorelVBA", "ExportToJpg", "Smoot", "10")
  

  mUse = (GetSetting("SanchoCorelVBA", "ExportToJpg", "mUse", "0") = "1")
  myMultSend = (GetSetting("SanchoCorelVBA", "ExportToJpg", "myMultSend", "0") = "1")
  myBatExe = GetSetting("SanchoCorelVBA", "ExportToJpg", "myBatExe", "")
  mTheBatEXE.ControlTipText = "Find TheBat.exe " & Chr(40) & myBatExe & Chr(41)
  myAccount = GetSetting("SanchoCorelVBA", "ExportToJpg", "myAccount", "what's your mail account name?")


  mAddres.List = Split(GetSetting("SanchoCorelVBA", "ExportToJpg", "mAddres"), "|", 20)
  If mAddres.ListCount = 20 Then mAddres.RemoveItem 19
  If mAddres.ListCount = 0 Then
      mAddres = ""
  Else
      mAddres = GetSetting("SanchoCorelVBA", "ExportToJpg", "mAddres2", mAddres)
  End If

  mSubject.List = Split(GetSetting("SanchoCorelVBA", "ExportToJpg", "mSubject"), "|", 20)
  If mSubject.ListCount = 20 Then mSubject.RemoveItem 19
  If mSubject.ListCount = 0 Then
      mSubject = ""
  Else
      mSubject = GetSetting("SanchoCorelVBA", "ExportToJpg", "mSubject2", mSubject)
  End If

  If mUse = 0 Then
      mAddres.Enabled = False
      mSubject.Enabled = False
      mTheBatEXE.Enabled = False
      myAccount.Enabled = False
  End If
  
  myPageIndex.Text = GetSetting("SanchoCorelVBA", "ExportToJpg", "myPageIndex")
  myPageIndex.ControlTipText = "Pages: " & myPageIndex.Text
  
  myLoadPresets

  'MY ===============================
  If GetSetting("SanchoCorelVBA", "ExportToJpg", "AutoChange", "0") = "1" Then
      nameprefix_TextBox.Text = Replace(CorelDRAW.ActiveDocument.Name, ".cdr", "")
      folder_TextBox.Text = ActiveDocument.FilePath
      If ActiveDocument.Pages.Count > 1 Then myWorkMetod = "All Page" Else myWorkMetod = "Selection"
  End If
  
  Select Case Random(0, 3)
  Case 0
      cm_aboutCdrPreflight.Caption = "CdrPreflight. Read more..."
      macroAdURL = "http://cdrpro.ru/en/macros/cdrpreflight/"
  Case 1
      cm_aboutCdrPreflight.Caption = "CardGenerator. Read more..."
      macroAdURL = "http://cdrpro.ru/en/macros/cardgenerator/"
  Case Else
      cm_aboutCdrPreflight.Caption = "CorelDRAW Macros"
      macroAdURL = "http://cdrpro.ru/en/"
  End Select
  
End Sub

Private Sub windowSetStyle()
  Const WS_EX_LAYERED& = &H80000, GWL_STYLE& = -16
  SetWindowLongW sanUI, GWL_STYLE, GetWindowLongW(sanUI, GWL_STYLE) And Not &HCF0000
End Sub

Private Sub windowTransparent(Optional ByVal level As Byte)
  Const WS_EX_LAYERED& = &H80000, LWA_COLORKEY& = &H1, LWA_ALPHA& = &H2, GWL_EXSTYLE& = -20
  If level Then
    SetWindowLongW sanUI, GWL_EXSTYLE, GetWindowLongW(sanUI, GWL_EXSTYLE) Or WS_EX_LAYERED
    SetLayeredWindowAttributes sanUI, 0, level, LWA_ALPHA
  Else
    SetWindowLongW sanUI, GWL_EXSTYLE, GetWindowLongW(sanUI, GWL_EXSTYLE) And Not WS_EX_LAYERED
  End If
End Sub

Private Sub meHead_MouseDown(ByVal Button%, ByVal Shift%, ByVal X!, ByVal Y!)
  If Button = 1 Then
      On Error Resume Next
      Tag = "1"
      GetWindowRect sanUI, r0
      GetCursorPos cp0
      sx = GetSystemMetrics(SM_CXScreen)
      sy = GetSystemMetrics(SM_CYScreen)
  End If
End Sub

Private Sub meHead_MouseMove(ByVal Button%, ByVal Shift%, ByVal X!, ByVal Y!)
  If Len(Tag) = 0 Then Exit Sub
  On Error Resume Next
  Dim cp As wPOINT
  GetCursorPos cp
  cp.X = r0.l + cp.X - cp0.X: If cp.X < 0 Then cp.X = 0 Else If cp.X > sx - 26 Then cp.X = sx - 26
  cp.Y = r0.t + cp.Y - cp0.Y: If cp.Y < 0 Then cp.Y = 0 Else If cp.Y > sy - 26 Then cp.Y = sy - 26
  SetWindowPos sanUI, 0, cp.X, cp.Y, 0, 0, SWP_NOACTIVATE Or SWP_NOSIZE Or SWP_NOSENDCHANGING
End Sub

Private Sub meHead_MouseUp(ByVal Button%, ByVal Shift%, ByVal X!, ByVal Y!)
  On Error Resume Next
  Select Case Button
      Case 1:
      If Len(Tag) Then
      Tag = vbNullString
      GetWindowRect sanUI, r0
      Exit Sub
      End If
  End Select
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  Command_OK.BackColor = &H8000000F
  Command_OK.SpecialEffect = fmSpecialEffectEtched
  Command_OK.ForeColor = &H80000012
  cm_aboutCdrPreflight.SpecialEffect = fmSpecialEffectEtched
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "Pos", Left & " " & Top
End Sub

Private Function Random(Lowerbound As Long, Upperbound As Long) As Long
  Randomize
  Random = Int(Rnd * Upperbound) + Lowerbound
End Function

Private Sub myListsSave()
  If folder_TextBox.ListCount > 0 Then
      For Count = 0 To folder_TextBox.ListCount - 1
          If folder_TextBox = folder_TextBox.List(Count) Then yn = True
      Next Count

      If yn = False Then
          For i = 1 To folder_TextBox.ListCount
          s = s + folder_TextBox.List(i - 1) + "|"
          Next
          s = Left$(s, Len(s) - 1)

          If InStr(1, "|" + s + "|", "|" + folder_TextBox + "|", vbTextCompare) = 0 Then _
          folder_TextBox.AddItem folder_TextBox, 0

          s = folder_TextBox + IIf(s = sEmpty, sEmpty, "|") + s
          SaveSetting "SanchoCorelVBA", "ExportToJpg", "folder", s
      End If
  Else
      SaveSetting "SanchoCorelVBA", "ExportToJpg", "folder", folder_TextBox
  End If

  s = ""

  If mAddres.ListCount > 0 Then
      For Count = 0 To mAddres.ListCount - 1
          If mAddres = mAddres.List(Count) Then yn2 = True
      Next Count

      If yn2 = False Then
          For i = 1 To mAddres.ListCount
          s = s + mAddres.List(i - 1) + "|"
          Next
          s = Left$(s, Len(s) - 1)

          If InStr(1, "|" + s + "|", "|" + mAddres + "|", vbTextCompare) = 0 Then _
          mAddres.AddItem mAddres, 0

          s = mAddres + IIf(s = sEmpty, sEmpty, "|") + s
          SaveSetting "SanchoCorelVBA", "ExportToJpg", "mAddres", s
      End If
  Else
      SaveSetting "SanchoCorelVBA", "ExportToJpg", "mAddres", mAddres
  End If

  s = ""

  If mSubject.ListCount > 0 Then
      For Count = 0 To mSubject.ListCount - 1
          If mSubject = mSubject.List(Count) Then yn3 = True
      Next Count

      If yn3 = False Then
          For i = 1 To mSubject.ListCount
          s = s + mSubject.List(i - 1) + "|"
          Next
          s = Left$(s, Len(s) - 1)

          If InStr(1, "|" + s + "|", "|" + mSubject + "|", vbTextCompare) = 0 Then _
          mSubject.AddItem mSubject, 0

          s = mSubject + IIf(s = sEmpty, sEmpty, "|") + s
          SaveSetting "SanchoCorelVBA", "ExportToJpg", "mSubject", s
      End If
  Else
      SaveSetting "SanchoCorelVBA", "ExportToJpg", "mSubject", mSubject
  End If
End Sub

Private Sub myLoadPresets()
  Dim c&, i&, presName$
  c = GetSetting("SanchoCorelVBA", "ExportToJpg", "PresetsCount", 0)
  For i = 1 To c
      presName = GetSetting("SanchoCorelVBA", "ExportToJpg", "Presets" & i & "Name")
      If presName <> "" Then cb_presList.AddItem i & "| " & presName
  Next i
End Sub

Private Sub cb_presList_Change()
  Dim i&, a$(), c1&, a2$()

  If cb_presList.SelLength = 0 Then Exit Sub
  For c1 = 0 To cb_presList.ListCount - 1 Step 1
      If cb_presList.SelText = cb_presList.List(c1) Then
          a2 = Split(cb_presList.SelText, "|")
          i = CLng(a2(0))
          Exit For
      End If
  Next c1
  
  a = Split(GetSetting("SanchoCorelVBA", "ExportToJpg", "Presets" & i), "|")
  
  nameprefix_TextBox = a(1)
  'prefix_CheckBox = a(2)
  tb_AfterName = a(3)
  folder_TextBox = a(4)
  Dpi_TextBox = a(5)
  profile_CheckBox = a(6)
  AntiAliasing_CheckBox = a(7)
  bwidth = a(8)
  bheight = a(9)
  sAspectRat = a(10): cm_AspectRatMove
  myColorMode = a(11)
  jpgEncod = a(12)
  myCmpress = a(13)
  Smoot = a(14)
  mUse = a(15)
  If a(15) = "1" Then
      mAddres = a(16)
      mSubject = a(17)
      myMultSend = a(18)
      myAccount = a(19)
  End If
  
  'If a(0) = "2" Then
      'FtFillToBit = a(58)
  'End If
End Sub

Private Sub cm_presAdd_Click()
  Dim strPres$, strPresN$, c&
  c = GetSetting("SanchoCorelVBA", "ExportToJpg", "PresetsCount", 0)
  c = c + 1
  
  strPresN = InputBox("Name for Preset", "Name...")
  If strPresN = "" Then Exit Sub

  strPres = jpgPresVer & "|" & nameprefix_TextBox.Text & "|" & "1" & "|" & tb_AfterName.Text & "|" & folder_TextBox.Text & "|" & Dpi_TextBox.Text & "|" & _
      IIf(profile_CheckBox, "1", "0") & "|" & IIf(AntiAliasing_CheckBox, "1", "0") & "|" & bwidth.Text & "|" & bheight.Text & "|" & sAspectRat & "|" & myColorMode.Text & "|" & _
      jpgEncod.Text & "|" & myCmpress.Text & "|" & Smoot.Text & "|" & IIf(mUse, "1", "0") & "|" & mAddres.Text & "|" & mSubject.Text & "|" & IIf(myMultSend, "1", "0") & "|" & myAccount.Text
  
          
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & c, strPres
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & c & "Name", strPresN
  
  cb_presList.AddItem c & "| " & strPresN
  cb_presList.Text = c & "| " & strPresN
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "PresetsCount", c
End Sub

Private Sub cm_presDel_Click()
  Dim i&, c&, i2&, a$()
  
  If cb_presList.SelLength = 0 Then Exit Sub
  For c1 = 0 To cb_presList.ListCount - 1 Step 1
      If cb_presList.SelText = cb_presList.List(c1) Then
          a = Split(cb_presList.SelText, "|")
          i = CLng(a(0))
          Exit For
      End If
  Next c1
  
  cb_presList.Clear
  c = CLng(GetSetting("SanchoCorelVBA", "ExportToJpg", "PresetsCount", 0))
  If i < c Then
      For i2 = i + 1 To c Step 1
          SaveSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & i, _
          GetSetting("SanchoCorelVBA", "ExportToJpg", "Presets" & i2)
          SaveSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & i & "Name", _
          GetSetting("SanchoCorelVBA", "ExportToJpg", "Presets" & i2 & "Name")
          i = i + 1
      Next i2
      DeleteSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & c
      DeleteSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & c & "Name"
  Else
      DeleteSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & i
      DeleteSetting "SanchoCorelVBA", "ExportToJpg", "Presets" & i & "Name"
  End If
  SaveSetting "SanchoCorelVBA", "ExportToJpg", "PresetsCount", c - 1
  myLoadPresets
End Sub

Private Sub Image3_Click()
  uf_about.Show
End Sub

Private Sub cm_aboutCdrPreflight_Click()
  On Error Resume Next
  Dim RetVal#
  RetVal = Shell(GetString(HKEY_CLASSES_ROOT, "HTTP\shell\open\command", "") & _
  " " & macroAdURL, 1)
  '"http://cdrpro.ru/en/macros/cdrpreflight/"
  AppActivate RetVal
End Sub

Private Sub cm_aboutCdrPreflight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  cm_aboutCdrPreflight.SpecialEffect = fmSpecialEffectSunken
End Sub

Private Sub cm_aboutCdrPreflight_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  cm_aboutCdrPreflight.SpecialEffect = fmSpecialEffectRaised
End Sub

Function GetString(hKey As Long, strPath As String, strValue As String)
  Dim Ret
  RegOpenKey hKey, strPath, Ret
  GetString = RegQueryStringValue(Ret, strValue)
  RegCloseKey Ret
End Function

Function RegQueryStringValue(ByVal hKey As Long, ByVal strValueName As String) As String
  Dim lResult As Long, lValueType As Long, strBuf As String, lDataBufSize As Long
  
  lResult = RegQueryValueEx(hKey, strValueName, 0, lValueType, ByVal 0, lDataBufSize)
  If lResult = 0 Then
      If lValueType = REG_SZ Then
          strBuf = String(lDataBufSize, Chr$(0))
          lResult = RegQueryValueEx(hKey, strValueName, 0, 0, ByVal strBuf, lDataBufSize)
          If lResult = 0 Then
              RegQueryStringValue = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
          End If
      ElseIf lValueType = REG_BINARY Then
          Dim strData As Integer
          lResult = RegQueryValueEx(hKey, strValueName, 0, 0, strData, lDataBufSize)
          If lResult = 0 Then
              RegQueryStringValue = strData
          End If
      End If
  End If
End Function
