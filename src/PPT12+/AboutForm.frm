VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AboutForm 
   Caption         =   "About PPspliT"
   ClientHeight    =   6930
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9945
   OleObjectBlob   =   "AboutForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AboutForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CloseButton_Click()
    AboutForm.Hide
End Sub

Private Sub PaypalLabel_Click()
    ActivePresentation.FollowHyperlink "https://www.paypal.com/donate/?business=6JYP92W3XHXVY&no_recurring=1&currency_code=EUR"
End Sub

Private Sub WebsiteLabel_Click()
    ActivePresentation.FollowHyperlink "https://www.maxonthenet.altervista.org/"
End Sub

