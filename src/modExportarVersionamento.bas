Attribute VB_Name = "modExportarVersionamento"
Sub ExportarVBA()

    Dim comp As Object
    Dim caminho As String
    
    ' pasta src ao lado da pasta workbook
    caminho = ThisWorkbook.Path & "\..\src\"

    If Dir(caminho, vbDirectory) = "" Then
        MkDir caminho
    End If

    For Each comp In ThisWorkbook.VBProject.VBComponents

        Select Case comp.Type
        
            Case 1 ' módulo .bas
                comp.Export caminho & comp.Name & ".bas"
                
            Case 2 ' classe .cls
                comp.Export caminho & comp.Name & ".cls"
                
            Case 3 ' form .frm (+ .frx automático)
                comp.Export caminho & comp.Name & ".frm"
                
        End Select

    Next comp

    MsgBox "VBA exportado para /src com sucesso!", vbInformation

End Sub
