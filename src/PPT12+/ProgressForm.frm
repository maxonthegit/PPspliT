VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressForm 
   Caption         =   "PPspliT progress"
   ClientHeight    =   2430
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7260
   OleObjectBlob   =   "ProgressForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    If MsgBox("Are you sure you want to cancel the operation?", vbYesNo, "PPspliT question") = vbYes Then
        PPspliT.cancelStatus = True
    End If
End Sub
