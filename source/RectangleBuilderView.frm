VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RectangleBuilderView 
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7425
   OleObjectBlob   =   "RectangleBuilderView.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RectangleBuilderView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'===============================================================================

Public ULeftWidth As TextBoxHandler
Public ULeftHeight As TextBoxHandler
Public ULeftOffsetX As TextBoxHandler
Public ULeftOffsetY As TextBoxHandler

Public URightWidth As TextBoxHandler
Public URightHeight As TextBoxHandler
Public URightOffsetX As TextBoxHandler
Public URightOffsetY As TextBoxHandler

Public LLeftWidth As TextBoxHandler
Public LLeftHeight As TextBoxHandler
Public LLeftOffsetX As TextBoxHandler
Public LLeftOffsetY As TextBoxHandler

Public LRightWidth As TextBoxHandler
Public LRightHeight As TextBoxHandler
Public LRightOffsetX As TextBoxHandler
Public LRightOffsetY As TextBoxHandler

Public MainWidth As TextBoxHandler
Public MainHeight As TextBoxHandler

Public IsOk As Boolean
Public IsCancel As Boolean

'===============================================================================

Private Sub UserForm_Initialize()
    Me.Caption = APP_NAME
    
    Set ULeftWidth = _
        TextBoxHandler.New_(TextBoxULeftWidth, TextBoxTypeDouble, 0)
    Set ULeftHeight = _
        TextBoxHandler.New_(TextBoxULeftHeight, TextBoxTypeDouble, 0)
    Set ULeftOffsetX = _
        TextBoxHandler.New_(TextBoxULeftOffsetX, TextBoxTypeDouble, 0)
    Set ULeftOffsetY = _
        TextBoxHandler.New_(TextBoxULeftOffsetY, TextBoxTypeDouble, 0)
        
    Set URightWidth = _
        TextBoxHandler.New_(TextBoxURightWidth, TextBoxTypeDouble, 0)
    Set URightHeight = _
        TextBoxHandler.New_(TextBoxURightHeight, TextBoxTypeDouble, 0)
    Set URightOffsetX = _
        TextBoxHandler.New_(TextBoxURightOffsetX, TextBoxTypeDouble, 0)
    Set URightOffsetY = _
        TextBoxHandler.New_(TextBoxURightOffsetY, TextBoxTypeDouble, 0)
        
    Set LLeftWidth = _
        TextBoxHandler.New_(TextBoxLLeftWidth, TextBoxTypeDouble, 0)
    Set LLeftHeight = _
        TextBoxHandler.New_(TextBoxLLeftHeight, TextBoxTypeDouble, 0)
    Set LLeftOffsetX = _
        TextBoxHandler.New_(TextBoxLLeftOffsetX, TextBoxTypeDouble, 0)
    Set LLeftOffsetY = _
        TextBoxHandler.New_(TextBoxLLeftOffsetY, TextBoxTypeDouble, 0)
        
    Set LRightWidth = _
        TextBoxHandler.New_(TextBoxLRightWidth, TextBoxTypeDouble, 0)
    Set LRightHeight = _
        TextBoxHandler.New_(TextBoxLRightHeight, TextBoxTypeDouble, 0)
    Set LRightOffsetX = _
        TextBoxHandler.New_(TextBoxLRightOffsetX, TextBoxTypeDouble, 0)
    Set LRightOffsetY = _
        TextBoxHandler.New_(TextBoxLRightOffsetY, TextBoxTypeDouble, 0)
        
    Set MainWidth = _
        TextBoxHandler.New_(TextBoxMainWidth, TextBoxTypeDouble, 0.0001)
    Set MainHeight = _
        TextBoxHandler.New_(TextBoxMainHeight, TextBoxTypeDouble, 0.0001)
End Sub

Private Sub UserForm_Activate()
    '
End Sub

Private Sub btnOk_Click()
    FormŒ 
End Sub

Private Sub btnCancel_Click()
    FormCancel
End Sub

'===============================================================================

Private Sub FormŒ ()
    Me.Hide
    IsOk = True
End Sub

Private Sub FormCancel()
    Me.Hide
    IsCancel = True
End Sub

'===============================================================================

Private Sub UserForm_QueryClose(—ancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        —ancel = True
        FormCancel
    End If
End Sub
