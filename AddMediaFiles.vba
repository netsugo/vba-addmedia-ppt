' 連番のメディアファイルを各スライドに挿入します
Sub AddMediaFiles()
    
    Const mediaTargetDir = "voiceroid" ' change it
    Const filePattern = "voiceroid-%d.wav" ' change it
    
    Dim cd As String
    cd = ActivePresentation.Path
    
    Dim i As Long
    With ActivePresentation
        For i = 1 To .Slides.Count
            Dim mediaTargetPath As String
            mediaTargetPath = cd & "\" & mediaTargetDir & "\" & Replace(filePattern, "%d", (i - 1))
            Dim oShp As Shape
            Set oShp = ActivePresentation.Slides(i).Shapes.AddMediaObject2(mediaTargetPath)
            With oShp.AnimationSettings.PlaySettings
                .PlayOnEntry = True
                .HideWhileNotPlaying = True
            End With
        Next i
    End With

    MsgBox "メディアの埋め込みが完了しました。"

End Sub
