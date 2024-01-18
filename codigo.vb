Module codigo
    Public galo(2, 2) As Integer
    Public jog1, jog2 As String
    Public jogada, posicoesNaoPreenchidas, num_jogo, aux, ptj1, ptj2 As Integer

    Sub final_jogo()
        FrmGalo.Btn00.Enabled = False
        FrmGalo.Btn01.Enabled = False
        FrmGalo.Btn02.Enabled = False
        FrmGalo.Btn10.Enabled = False
        FrmGalo.Btn11.Enabled = False
        FrmGalo.Btn12.Enabled = False
        FrmGalo.Btn20.Enabled = False
        FrmGalo.Btn21.Enabled = False
        FrmGalo.Btn22.Enabled = False
        FrmGalo.BtnNovoJogo.Visible = True
    End Sub

    Sub verifica()
        'IFS RELATIVOS AO JOGADOR 1
        If galo(0, 0) = 1 And galo(0, 1) = 1 And galo(0, 2) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(1, 0) = 1 And galo(1, 1) = 1 And galo(1, 2) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(2, 0) = 1 And galo(2, 1) = 1 And galo(2, 2) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(0, 0) = 1 And galo(1, 0) = 1 And galo(2, 0) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(0, 1) = 1 And galo(1, 1) = 1 And galo(2, 1) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(0, 2) = 1 And galo(1, 2) = 1 And galo(2, 2) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(0, 0) = 1 And galo(1, 1) = 1 And galo(2, 2) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        If galo(0, 2) = 1 And galo(1, 1) = 1 And galo(2, 0) = 1 Then
            MsgBox("Ganhou o " & jog1)
            ptj1 = ptj1 + 1
            FrmGalo.LblPontosJogador1.Text = ptj1
            final_jogo()
            Exit Sub
        End If
        'JOGADOR 2
        If galo(0, 0) = 2 And galo(0, 1) = 2 And galo(0, 2) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(1, 0) = 2 And galo(1, 1) = 2 And galo(1, 2) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(2, 0) = 2 And galo(2, 1) = 2 And galo(2, 2) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(0, 0) = 2 And galo(1, 0) = 2 And galo(2, 0) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(0, 1) = 2 And galo(1, 1) = 2 And galo(2, 1) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(0, 2) = 2 And galo(1, 2) = 2 And galo(2, 2) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj1 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(0, 0) = 2 And galo(1, 1) = 2 And galo(2, 2) = 2 Then
            MsgBox("Ganhou o " & jog2)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        If galo(0, 2) = 2 And galo(1, 1) = 2 And galo(2, 0) = 2 Then
            MsgBox("Ganhou o " & jog1)
            ptj2 = ptj2 + 1
            FrmGalo.LblPontosJogador2.Text = ptj2
            final_jogo()
            Exit Sub
        End If
        posicoesNaoPreenchidas = 0
        For i = 0 To 2
            For j = 0 To 2
                If galo(i, j) = 0 Then
                    posicoesNaoPreenchidas += 1
                End If
            Next
        Next
        If posicoesNaoPreenchidas = 0 Then
            MsgBox("Empate")
            final_jogo()
        End If

    End Sub

    Sub novo_jogo()
        Dim i, j As Integer
        'Limpar o ARRAY
        For i = 0 To 2
            For j = 0 To 2
                galo(i, j) = 0
            Next
        Next
        'ACTIVAR OS BOTÕES
        FrmGalo.Btn00.Enabled = True
        FrmGalo.Btn01.Enabled = True
        FrmGalo.Btn02.Enabled = True
        FrmGalo.Btn10.Enabled = True
        FrmGalo.Btn11.Enabled = True
        FrmGalo.Btn12.Enabled = True
        FrmGalo.Btn20.Enabled = True
        FrmGalo.Btn21.Enabled = True
        FrmGalo.Btn22.Enabled = True
        'Limpar botões
        FrmGalo.Btn00.Text = ""
        FrmGalo.Btn01.Text = ""
        FrmGalo.Btn02.Text = ""
        FrmGalo.Btn10.Text = ""
        FrmGalo.Btn11.Text = ""
        FrmGalo.Btn12.Text = ""
        FrmGalo.Btn20.Text = ""
        FrmGalo.Btn21.Text = ""
        FrmGalo.Btn22.Text = ""
        FrmGalo.BtnNovoJogo.Visible = False
    End Sub

End Module
