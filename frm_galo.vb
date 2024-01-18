Public Class FrmGalo


















    Private Sub Btn00_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn00.Click
        If galo(0, 0) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 = 1 Then
                galo(0, 0) = 1
                Btn00.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(0, 0) = 2
                Btn00.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Posição já preenchida!")
        End If
    End Sub

    Private Sub Btn01_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn01.Click
        If galo(0, 1) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 = 1 Then
                galo(0, 1) = 1
                Btn01.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(0, 1) = 2
                Btn01.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente.")
        End If
    End Sub

    Private Sub btn_13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn02.Click
        If galo(0, 2) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(0, 2) = 1
                Btn02.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(0, 2) = 2
                Btn02.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn10.Click
        If galo(1, 0) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(1, 0) = 1
                Btn10.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(1, 0) = 2
                Btn10.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn11.Click
        If galo(1, 1) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(1, 1) = 1
                Btn11.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(1, 1) = 2
                Btn11.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn12.Click
        If galo(1, 2) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(1, 2) = 1
                Btn12.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(1, 2) = 2
                Btn12.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn20.Click
        If galo(2, 0) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(2, 0) = 1
                Btn20.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(2, 0) = 2
                Btn20.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn21.Click
        If galo(2, 1) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(2, 1) = 1
                Btn21.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(2, 1) = 2
                Btn21.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub

    Private Sub btn_33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn22.Click
        If galo(2, 2) = 0 Then
            jogada = jogada + 1
            If jogada Mod 2 <> 0 Then
                galo(2, 2) = 1
                Btn22.Text = "X"
                LblJogador1.ForeColor = Color.Black
                LblJogador2.ForeColor = Color.Red
            Else
                galo(2, 2) = 2
                Btn22.Text = "O"
                LblJogador2.ForeColor = Color.Black
                LblJogador1.ForeColor = Color.Red
            End If
            verifica()
        Else
            MsgBox("Esta Posição já foi preenchida anteriormente")
        End If
    End Sub



    Private Sub frm_galo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        jog1 = InputBox("Jogador 1:")
        jog2 = InputBox("Jogador 2:")
        jogada = 0
        ptj1 = 0
        ptj2 = 0
        num_jogo = 1
        '#########################################
        'Inicializar o array
        '#########################################
        For i = 0 To 2
            For j = 0 To 2
                galo(i, j) = 0
            Next
        Next
        LblJogador1.Text = jog1
        LblJogador2.Text = jog2
        LblJogador1.ForeColor = Color.Red
        LblPontosJogador1.Text = ptj1
        LblPontosJogador2.Text = ptj2
    End Sub




    Private Sub NovoJogoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NovoJogoToolStripMenuItem.Click
        novo_jogo()
    End Sub

    Private Sub NovaSérieToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NovaSérieToolStripMenuItem.Click
        jog1 = InputBox("Jogador 1:")
        jog2 = InputBox("Jogador 2:")
        LblJogador1.Text = jog1
        LblJogador2.Text = jog2
        For i = 0 To 2
            For j = 0 To 2
                galo(i, j) = 0
            Next
        Next
        jogada = 0
        LblJogador1.ForeColor = Color.Red
        'num_jogo = 1
        ptj1 = 0
        ptj2 = 0
        LblPontosJogador1.Text = ptj1
        LblPontosJogador2.Text = ptj2
        Btn00.Enabled = True
        Btn01.Enabled = True
        Btn02.Enabled = True
        Btn10.Enabled = True
        Btn11.Enabled = True
        Btn12.Enabled = True
        Btn20.Enabled = True
        Btn21.Enabled = True
        Btn22.Enabled = True
        'Limpar botões
        Btn00.Text = ""
        Btn01.Text = ""
        Btn02.Text = ""
        Btn10.Text = ""
        Btn11.Text = ""
        Btn12.Text = ""
        Btn20.Text = ""
        Btn21.Text = ""
        Btn22.Text = ""
    End Sub

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label1.Click

    End Sub


    Private Sub cmd_novojogo_Click(sender As System.Object, e As System.EventArgs) Handles BtnNovoJogo.Click
        novo_jogo()
    End Sub
End Class