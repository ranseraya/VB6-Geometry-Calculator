Attribute VB_Name = "modTema"
Public Sub TerapkanTema(ByVal frm As Form)
    Dim ctrl As Control
    
    frm.BackColor = &H25201A
    
    For Each ctrl In frm.Controls
        Select Case TypeName(ctrl)
            
            Case "TextBox"
                ctrl.BackColor = &H333333
                ctrl.ForeColor = &HE0E0E0
                ctrl.BorderStyle = 1
                ctrl.Font.Name = "Consolas"
                ctrl.Font.Size = 10
                
            Case "CommandButton"
                    ctrl.BackColor = &H333333
                    ctrl.Font.Name = "Consolas"
                    ctrl.Font.Size = 10
                    ctrl.Font.Bold = True

            Case "Label"
                If ctrl.Tag = "Tombol" Then
                    ctrl.BackColor = &H333333
                    ctrl.ForeColor = &HFF7F&
                    ctrl.Font.Name = "Consolas"
                    ctrl.Font.Size = 10
                    ctrl.Font.Bold = True
                    ctrl.BorderStyle = 1
                    ctrl.Alignment = 2
                
                ElseIf ctrl.Tag = "Hasil" Then
                    ctrl.ForeColor = &HFF7F&
                    ctrl.BackStyle = 0
                    ctrl.Font.Name = "Consolas"
                    ctrl.Font.Size = 11
                    ctrl.Font.Bold = True
                    
                ElseIf ctrl.Tag = "Judul" Then
                    ctrl.ForeColor = &HFFD300
                    ctrl.BackStyle = 0
                    ctrl.Font.Name = "Consolas"
                    ctrl.Font.Size = 12
                    ctrl.Font.Bold = True
                    
                Else
                    ctrl.ForeColor = &HFFD300
                    ctrl.BackStyle = 0
                    ctrl.Font.Name = "Consolas"
                    ctrl.Font.Size = 10
                End If
        
        End Select
    Next ctrl
End Sub
