VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Login"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5400
   OleObjectBlob   =   "frmLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnLogin_Click()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long

    Dim usuario As String
    Dim senha As String

    usuario = cmbUsuario.Value
    senha = txtSenha.Value

    If usuario = "" Then
        MsgBox "Selecione um usuário.", vbExclamation
        Exit Sub
    End If

    Set ws = Sheets("Base_Supervisores")
    Set tbl = ws.ListObjects("tblSupervisores")

    For i = 1 To tbl.DataBodyRange.Rows.Count

        If tbl.DataBodyRange(i, 2).Value = usuario And _
           tbl.DataBodyRange(i, 4).Value = senha Then

            gSupervisorID = tbl.DataBodyRange(i, 1).Value
            gSupervisorNome = usuario

            If usuario = "ADM" Then
                gIsAdmin = True
            Else
                gIsAdmin = False
            End If
            
            Application.Visible = True
            Me.Hide
            Sheets("Home").Activate

            Exit Sub

        End If

    Next i

    MsgBox "Usuário ou senha inválidos.", vbCritical

End Sub

Private Sub UserForm_Initialize()

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim cel As Range

    Set ws = Sheets("Base_Supervisores")
    Set tbl = ws.ListObjects("tblSupervisores")

    cmbUsuario.Clear

    For Each cel In tbl.ListColumns("Nome").DataBodyRange
        cmbUsuario.AddItem cel.Value
    Next cel

End Sub

Private Sub btnCancelar_Click()

    Application.DisplayAlerts = False
    ThisWorkbook.Close

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If CloseMode = 0 Then ' clicou no X
        Application.DisplayAlerts = False
        ThisWorkbook.Close
    End If

End Sub


