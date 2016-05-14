VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_about 
   Caption         =   "About..."
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   3660
   OleObjectBlob   =   "uf_about.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cb_OK_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
  Label1.Caption = "Export to JPG, version " & jpgVer & " " & jpgDate & Chr(10) & _
                   "Copyright " & Chr(169) & " 2006-2009 by Sancho" & Chr(10) & _
                   Chr(10) & _
                   "http://" & myWebSite & Chr(10) & _
                   "e-mail: " & myEmail & Chr(10) & _
                   Chr(10) & _
                   "CorelDRAW " & CorelDRAW.Version
End Sub
