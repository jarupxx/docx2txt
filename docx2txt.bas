Sub docx2txt()
'
' dicx2txt Macro
' SourcePathにWord(.docx) ファイルのフォルダー
' DestinationPathに出力先フォルダーを設定してください
' 出力先フォルダーは作成しておいてください
'
Dim wordFile As String, msg As String, txtFileName As String, hoge As String, SourcePath As String, DestinationPath As String
SourcePath = "C:\before\"
DestinationPath = "C:\after\"
wordFile = Dir(SourcePath & "*.docx")
Do While wordFile <> ""
    txtFileName = DestinationPath & Replace(wordFile, ".docx", ".txt")
    wordFile = SourcePath & wordFile
    Documents.Open wordFile
    Documents(wordFile).Activate
    ActiveDocument.SaveAs2 FileName:=txtFileName, FileFormat:=wdFormatText, Encoding:=msoEncodingUTF8, LineEnding:=wdCRLF
    ActiveDocument.Close
    msg = msg & txtFileName & vbCrLf
    wordFile = Dir()
Loop
MsgBox msg

End Sub
