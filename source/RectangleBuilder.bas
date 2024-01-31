Attribute VB_Name = "RectangleBuilder"
'===============================================================================
'   Макрос          : RectangleBuilder
'   Версия          : 2024.01.31
'   Сайты           : https://vk.com/elvin_macro
'                     https://github.com/elvin-nsk
'   Автор           : elvin-nsk (me@elvin.nsk.ru)
'===============================================================================

Option Explicit

Public Const RELEASE As Boolean = True

Public Const APP_NAME As String = "RectangleBuilder"

'===============================================================================

Private Enum EnumCorner
    UpperLeft
    UpperRight
    LowerLeft
    LowerRight
End Enum

'===============================================================================

Sub Start()

    If RELEASE Then On Error GoTo Catch
    
    With InputData.RequestDocumentOrPage(PageCanBeEmpty:=True)
        If .IsError Then GoTo Finally
    End With
    
    Dim MainShape As Shape
    Dim TrimShape As Shape
    With New RectangleBuilderView
    
        .Show vbModal
        If .IsCancel Then Exit Sub
        
        BoostStart APP_NAME, RELEASE
        
        Set MainShape = _
            ActiveLayer.CreateRectangle2(0, 0, .MainWidth, .MainHeight)
        
        TryCreateAndTrim _
            MainShape, UpperLeft, _
            .ULeftWidth, .ULeftHeight, _
            .ULeftOffsetX, .ULeftOffsetY
        
        TryCreateAndTrim _
            MainShape, UpperRight, _
            .URightWidth, .URightHeight, _
            .URightOffsetX, .URightOffsetY
       
        TryCreateAndTrim _
            MainShape, LowerLeft, _
            .LLeftWidth, .LLeftHeight, _
            .LLeftOffsetX, .LLeftOffsetY
        
        TryCreateAndTrim _
            MainShape, LowerRight, _
            .LRightWidth, .LRightHeight, _
            .LRightOffsetX, .LRightOffsetY
        
    End With
    
    Align PackShapes(MainShape), ActivePage.BoundingBox, cdrCenter
    
Finally:
    BoostFinish
    Exit Sub

Catch:
    VBA.MsgBox VBA.Err.Source & ": " & VBA.Err.Description, vbCritical, "Error"
    Resume Finally

End Sub

'===============================================================================

Private Sub TryCreateAndTrim( _
                     ByRef MainShape As Shape, _
                     ByVal Corner As EnumCorner, _
                     ByVal Width As Double, _
                     ByVal Height As Double, _
                     ByVal OffsetX As Double, _
                     ByVal OffsetY As Double _
                 )
    If Width = 0 Or Height = 0 Then Exit Sub
    Dim TrimShape As Shape
    Set TrimShape = _
            CreateAjacentRectangle( _
                MainShape, UpperLeft, _
                Width, Height, _
                OffsetX, OffsetY _
            )
    TrimAndDelete TrimShape, MainShape
End Sub

Private Function CreateAjacentRectangle( _
                     ByVal Anchor As Shape, _
                     ByVal Corner As EnumCorner, _
                     ByVal Width As Double, _
                     ByVal Height As Double, _
                     ByVal OffsetX As Double, _
                     ByVal OffsetY As Double _
                 ) As Shape
    Dim Result As Shape
    Set Result = _
        Anchor.Layer.CreateRectangle2(0, 0, Width, Height)
    Select Case Corner
        Case EnumCorner.UpperLeft
            Result.TopY = Anchor.TopY - OffsetY
            Result.LeftX = Anchor.LeftX + OffsetX
        Case EnumCorner.UpperRight
            Result.TopY = Anchor.TopY - OffsetY
            Result.RightX = Anchor.RightX - OffsetX
        Case EnumCorner.LowerLeft
            Result.BottomY = Anchor.BottomY + OffsetY
            Result.LeftX = Anchor.LeftX + OffsetX
        Case EnumCorner.LowerRight
            Result.BottomY = Anchor.BottomY + OffsetY
            Result.RightX = Anchor.RightX - OffsetX
    End Select
    Set CreateAjacentRectangle = Result
End Function

Private Sub TrimAndDelete( _
               ByVal TrimmerShape As Shape, _
               ByRef TargetShape As Shape _
           )
    Trim TrimmerShape, TargetShape
    TrimmerShape.Delete
End Sub


'===============================================================================
' # тесты

Private Sub testSomething()
'
End Sub
