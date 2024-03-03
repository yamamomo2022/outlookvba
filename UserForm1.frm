VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7660
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11070
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton_èoãŒïÒçê_éññ±èä_Click()
    Call sendmail
    Unload Me
    Application.ActiveExplorer.WindowState = olMaximized

End Sub

Private Sub UserForm_Click()

End Sub
