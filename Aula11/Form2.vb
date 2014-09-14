Imports System.Data.OleDb

Public Class Form2
    Dim acao As String
    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Desabilita()
        ToolStrip2.Visible = False
    End Sub
    Sub Habilita()
        'Me.TxtCodigo.Enabled = True
        Me.Txtnm.Enabled = True
        Me.TxtCPF.Enabled = True
        Me.TxtRG.Enabled = True
        Me.Dtpdtnscm.Enabled = True
        Me.TxtlocalNscm.Enabled = True
        Me.CmbUF.Enabled = True
        Me.Cmbsx.Enabled = True
        Me.Txttel.Enabled = True
        Me.Txtcel.Enabled = True
        Me.TxtopCel.Enabled = True
        Me.Txtemail.Enabled = True
        Me.Txtface.Enabled = True
        Me.TxtlgdrTp.Enabled = True
        Me.TxtlgdrNm.Enabled = True
        Me.TxtlgdrNum.Enabled = True
        Me.TxtlgdrCmpl.Enabled = True
        Me.TxtlgdrBairro.Enabled = True
        Me.TxtlgdrCid.Enabled = True
        Me.CmblgdrUF.Enabled = True
        Me.TxtlgdrCEP.Enabled = True
        'Me.DtpDataCadastro.Enabled = True
        'Me.TxtSituacao.Enabled = True
        'Me.DtpDataSituacao.Enabled = True
        Me.Txtnm.Focus()
    End Sub

    Sub Desabilita()
        'Me.TxtCodigo.Enabled = False
        Me.Txtnm.Enabled = False
        Me.TxtCPF.Enabled = False
        Me.TxtRG.Enabled = False
        Me.Dtpdtnscm.Enabled = False
        Me.TxtlocalNscm.Enabled = False
        Me.CmbUF.Enabled = False
        Me.Cmbsx.Enabled = False
        Me.Txttel.Enabled = False
        Me.Txtcel.Enabled = False
        Me.TxtopCel.Enabled = False
        Me.Txtemail.Enabled = False
        Me.Txtface.Enabled = False
        Me.TxtlgdrTp.Enabled = False
        Me.TxtlgdrNm.Enabled = False
        Me.TxtlgdrNum.Enabled = False
        Me.TxtlgdrCmpl.Enabled = False
        Me.TxtlgdrBairro.Enabled = False
        Me.TxtlgdrCid.Enabled = False
        Me.CmblgdrUF.Enabled = False
        Me.TxtlgdrCEP.Enabled = False
        Me.Txtcdg.Focus()
        'Me.DtpDataCadastro.Enabled = False
        'Me.TxtSituacao.Enabled = False
        'Me.DtpDataSituacao.Enabled = False
    End Sub
    Sub Limpa_total()
        Me.Txtcdg.Clear()
        Me.Txtnm.Clear()
        Me.TxtCPF.Clear()
        Me.TxtRG.Clear()
        Me.Dtpdtnscm.Value = Now
        Me.TxtlocalNscm.Clear()
        Me.CmbUF.Text = vbNullString
        Me.Cmbsx.Text = vbNullString
        Me.Txttel.Clear()
        Me.Txtcel.Clear()
        Me.TxtopCel.Clear()
        Me.Txtemail.Clear()
        Me.Txtface.Clear()
        Me.TxtlgdrTp.Clear()
        Me.TxtlgdrNm.Clear()
        Me.TxtlgdrNum.Clear()
        Me.TxtlgdrCmpl.Clear()
        Me.TxtlgdrBairro.Clear()
        Me.TxtlgdrCid.Clear()
        Me.CmblgdrUF.Text = vbNullString
        Me.TxtlgdrCEP.Clear()
        Me.Dtpcad.Value = Now
        Me.Txtsit().Clear()
        Me.Dtpsit.Value = Now
    End Sub
    Sub Limpa()
        Me.Txtnm.Clear()
        Me.TxtCPF.Clear()
        Me.TxtRG.Clear()
        Me.Dtpdtnscm.Value = Now
        Me.TxtlocalNscm.Clear()
        Me.CmbUF.Text = vbNullString
        Me.Cmbsx.Text = vbNullString
        Me.Txttel.Clear()
        Me.Txtcel.Clear()
        Me.TxtopCel.Clear()
        Me.Txtemail.Clear()
        Me.Txtface.Clear()
        Me.TxtlgdrTp.Clear()
        Me.TxtlgdrNm.Clear()
        Me.TxtlgdrNum.Clear()
        Me.TxtlgdrCmpl.Clear()
        Me.TxtlgdrBairro.Clear()
        Me.TxtlgdrCid.Clear()
        Me.CmblgdrUF.Text = vbNullString
        Me.TxtlgdrCEP.Clear()
        Me.Dtpcad.Value = Now
        Me.Txtsit().Clear()
        Me.Dtpsit.Value = Now
    End Sub

    Private Sub Tstincluir_Click(sender As Object, e As EventArgs) Handles Tstincluir.Click
        acao = "INCLUIR"
        Habilita()
        Limpa()
        ToolStrip1.Visible = False
        ToolStrip2.Visible = True
        Txtnm.Focus()
    End Sub

    Private Sub Tstcancel_Click(sender As Object, e As EventArgs) Handles Tstcancel.Click
        Me.Dispose()
        Form1.CadastrosToolStripMenuItem.Checked = False
        Form1.CadastrosToolStripMenuItem.Enabled = True
    End Sub

    Private Sub Tsbtncancelar_Click(sender As Object, e As EventArgs) Handles Tsbtncancelar.Click
        ToolStrip2.Visible = False
        ToolStrip1.Visible = True
        Limpa_total()
        Desabilita()
    End Sub

    Private Sub Tstedit_Click(sender As Object, e As EventArgs) Handles Tstedit.Click
        Dim codigo As Long = Convert.ToInt64(InputBox("Digite o código do cliente:", "Atenção", "0"))
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim reader As OleDbDataReader
        Dim cont As Integer = 0
        Dim sql As String
        acao = "EDITAR"

        sql = "SELECT Nome, Codigo, DataDeNascimento, CidadeDeNascimento, UFDeNascimento, RG, CPF, Sexo, TelefoneResidencial, TelefoneCelular, OperadoraCelular, Email, Facebook, TipoDeLogadouroResidencial, LogadouroResidencial, NumeroResidencial, ComplementoResidencial, BairroResidencial, CidadeResidencial, EstadoResidencial, CEPResidencial, DataDeCadastramento, DataDaSituacao, Situacao FROM Cliente WHERE Codigo =" & codigo.ToString

        cn.ConnectionString = My.Settings.banco201302ConnectionString
        cn.Open()

        cmd.CommandType = CommandType.Text
        cmd.CommandText = sql
        cmd.Connection = cn

        reader = cmd.ExecuteReader

        'carga das informações
        If reader.Read Then
            Habilita()
            ToolStrip1.Visible = False
            ToolStrip2.Visible = True
            Me.Txtcdg.Text = reader.Item("Codigo")
            Me.Txtnm.Text = reader.Item("Nome")
            Me.TxtCPF.Text = reader.Item("CPF")
            Me.TxtRG.Text = reader.Item("RG")
            Me.Dtpdtnscm.Value = reader.Item("DataDeNascimento")
            Me.TxtlocalNscm.Text = reader.Item("CidadeDeNascimento")
            Me.CmbUF.Text = reader.Item("UFDeNascimento")
            Me.Cmbsx.Text = reader.Item("Sexo")
            Me.Txttel.Text = reader.Item("TelefoneResidencial")
            Me.Txtcel.Text = reader.Item("TelefoneCelular")
            Me.TxtopCel.Text = reader.Item("OperadoraCelular")
            Me.Txtemail.Text = reader.Item("Email")
            Me.Txtface.Text = reader.Item("Facebook")
            Me.TxtlgdrTp.Text = reader.Item("TipodeLogadouroResidencial")
            Me.TxtlgdrNm.Text = reader.Item("LogadouroResidencial")
            Me.TxtlgdrNum.Text = reader.Item("NumeroResidencial")
            Me.TxtlgdrCmpl.Text = reader.Item("ComplementoResidencial")
            Me.TxtlgdrBairro.Text = reader.Item("BairroResidencial")
            Me.TxtlgdrCid.Text = reader.Item("CidadeResidencial")
            Me.CmblgdrUF.Text = reader.Item("EstadoResidencial")
            Me.TxtlgdrCEP.Text = reader.Item("CEPResidencial")
            Me.Dtpcad.Value = reader.Item("DataDeCadastramento")
            Me.Dtpsit.Value = reader.Item("DataDaSituacao")
            Me.Txtsit.Text = reader.Item("Situacao")
        Else
            MessageBox.Show("Registro não localizado")
        End If
    End Sub

    Private Sub Tstexcluir_Click(sender As Object, e As EventArgs) Handles Tstexcluir.Click
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        Dim reader As OleDbDataReader
        Dim cont As Integer = 0
        Dim sql As String
        Dim codigo As Long = Convert.ToInt64(InputBox("Digite o código do cliente:", "Atenção", "0"))
        acao = "EXCLUIR"

        sql = "SELECT Nome, Codigo, DataDeNascimento, CidadeDeNascimento, UFDeNascimento, RG, CPF, Sexo, TelefoneResidencial, TelefoneCelular, OperadoraCelular, Email, Facebook, TipoDeLogadouroResidencial, LogadouroResidencial, NumeroResidencial, ComplementoResidencial, BairroResidencial, CidadeResidencial, EstadoResidencial, CEPResidencial, DataDeCadastramento, DataDaSituacao, Situacao FROM Cliente WHERE Codigo =" & codigo.ToString

        cn.ConnectionString = My.Settings.banco201302ConnectionString
        cn.Open()
        cmd.CommandType =
        CommandType.Text
        cmd.CommandText = sql
        cmd.Connection = cn
        reader = cmd.ExecuteReader
        'carga das informações
        If reader.Read Then
            Habilita()
            ToolStrip1.Visible = False
            ToolStrip2.Visible = True
            Me.Txtcdg.Text = reader.Item("Codigo")
            Me.Txtnm.Text = reader.Item("Nome")
            Me.TxtCPF.Text = reader.Item("CPF")
            Me.TxtRG.Text = reader.Item("RG")
            Me.Dtpdtnscm.Value = reader.Item("DataDeNascimento")
            Me.TxtlocalNscm.Text = reader.Item("CidadeDeNascimento")
            Me.CmbUF.Text = reader.Item("UFDeNascimento")
            Me.Cmbsx.Text = reader.Item("Sexo")
            Me.Txttel.Text = reader.Item("TelefoneResidencial")
            Me.Txtcel.Text = reader.Item("TelefoneCelular")
            Me.TxtopCel.Text = reader.Item("OperadoraCelular")
            Me.Txtemail.Text = reader.Item("Email")
            Me.Txtface.Text = reader.Item("Facebook")
            Me.TxtlgdrTp.Text = reader.Item("TipoDeLogadouroResidencial")
            Me.TxtlgdrNm.Text = reader.Item("LogadouroResidencial")
            Me.TxtlgdrNum.Text = reader.Item("NumeroResidencial")
            Me.TxtlgdrCmpl.Text = reader.Item("ComplementoResidencial")
            Me.TxtlgdrBairro.Text = reader.Item("BairroResidencial")
            Me.TxtlgdrCid.Text = reader.Item("CidadeResidencial")
            Me.CmblgdrUF.Text = reader.Item("EstadoResidencial")
            Me.TxtlgdrCEP.Text = reader.Item("CEPResidencial")
            Me.Dtpcad.Value = reader.Item("DataDeCadastramento")
            Me.Dtpsit.Value = reader.Item("DataDaSituacao")
            Me.Txtsit.Text = reader.Item("Situacao")
        Else
            MessageBox.Show("Registro não localizado")
        End If
    End Sub

    Private Sub Tstbtngravar_Click(sender As Object, e As EventArgs) Handles Tstbtngravar.Click
        Dim codigo, nome, dataDeNascimento, cidadeDeNascimento, UFDeNascimento, rg, cpf, sexo, telefoneResidencial, telefoneCelular, operadoraCelular, email, facebook, tipoDeLogadouroResidencial, logadouroResidencial, numeroResidencial, complementoResidencial, bairroResidencial, cidadeResidencial, estadoResidencial, cepResidencial, situacao, dataDeCadastramento, dataDaSituacao As String
        Dim sql As String
        Dim cn As New OleDbConnection
        Dim cmd As New OleDbCommand
        If acao = "INCLUIR" Then
            'carga das informações
            codigo = Txtcdg.Text
            nome = Txtnm.Text
            dataDeNascimento = Dtpdtnscm.Value
            cidadeDeNascimento = TxtlocalNscm.Text
            UFDeNascimento = CmbUF.Text
            rg = TxtRG.Text
            cpf = TxtCPF.Text
            sexo = Cmbsx.Text
            telefoneResidencial = Txttel.Text
            telefoneCelular = Txtcel.Text
            operadoraCelular = TxtopCel.Text
            email = Txtemail.Text
            facebook = Txtface.Text
            tipoDeLogadouroResidencial = TxtlgdrTp.Text
            logadouroResidencial = TxtlgdrNm.Text
            numeroResidencial = TxtlgdrNum.Text
            complementoResidencial = TxtlgdrCmpl.Text
            bairroResidencial = TxtlgdrBairro.Text
            cidadeResidencial = TxtlgdrCid.Text
            cidadeDeNascimento = TxtlgdrCid.Text
            estadoResidencial = CmblgdrUF.Text
            cepResidencial = TxtlgdrCEP.Text
            situacao = "A"
            dataDeCadastramento = Now.ToString
            dataDaSituacao = Now.ToString

            sql = "INSERT INTO Cliente" &
                "( Nome" &
                ", Codigo" &
                ", DataDeNascimento" &
                ", CidadeDeNascimento" &
                ", UFDeNascimento" &
                ", RG" &
                ", CPF" &
                ", Sexo" &
                ", TelefoneResidencial" &
                ", TelefoneCelular" &
                ", OperadoraCelular" &
                ", Email" &
                ", Facebook" &
                ", TipoDeLogadouroResidencial" &
                ", LogadouroResidencial" &
                ", NumeroResidencial" &
                ", ComplementoResidencial" &
                ", BairroResidencial" &
                ", CidadeResidencial" &
                ", EstadoResidencial" &
                ", CEPResidencial" &
                ", DataDeCadastramento" &
                ", DataDaSituacao" &
                ", Situacao)" &
                " VALUES ('" &
                nome & "','" &
                codigo & "','" &
                dataDeNascimento & "','" &
                cidadeDeNascimento & "','" &
                UFDeNascimento & "', '" &
                rg & "','" &
                cpf & "', '" &
                sexo & "', '" &
                telefoneResidencial & "', '" &
                telefoneCelular & "', '" &
                operadoraCelular & "', '" &
                email & "', '" &
                facebook & "', '" &
                tipoDeLogadouroResidencial & "', '" &
                logadouroResidencial & "', '" &
                numeroResidencial & "', '" &
                complementoResidencial & "', '" &
                bairroResidencial & "', '" &
                cidadeResidencial & "', '" &
                estadoResidencial & "', '" &
                cepResidencial & "', '" &
                dataDeCadastramento & "', '" &
                dataDaSituacao & "', '" &
                situacao & "')"

            'abrir conexão
            cn.ConnectionString = My.Settings.banco201302ConnectionString
            cn.Open()
            'montar comando
            cmd.CommandType = CommandType.Text
            cmd.Connection = cn
            cmd.CommandText = sql
            'executar comando
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            cmd.Dispose()
            MessageBox.Show("Cliente incluido com sucesso...", "Atenção!")
            ToolStrip2.Visible = False
            ToolStrip1.Visible = True
            Limpa_total()
            Desabilita()
        ElseIf acao = "EDITAR" Then
            codigo = Txtcdg.Text
            nome = Txtnm.Text
            dataDeNascimento = Dtpdtnscm.Value
            cidadeDeNascimento = TxtlocalNscm.Text
            UFDeNascimento = CmbUF.Text
            rg = TxtRG.Text
            cpf = TxtCPF.Text
            sexo = Cmbsx.Text
            telefoneResidencial = Txttel.Text
            telefoneCelular = Txtcel.Text
            operadoraCelular = TxtopCel.Text
            email = Txtemail.Text
            facebook = Txtface.Text
            tipoDeLogadouroResidencial = TxtlgdrTp.Text
            logadouroResidencial = TxtlgdrNm.Text
            numeroResidencial = TxtlgdrNum.Text
            complementoResidencial = TxtlgdrCmpl.Text
            bairroResidencial = TxtlgdrBairro.Text
            cidadeResidencial = TxtlgdrCid.Text
            cidadeDeNascimento = TxtlgdrCid.Text
            estadoResidencial = CmblgdrUF.Text
            cepResidencial = TxtlgdrCEP.Text
            situacao = "E"
            dataDeCadastramento = Now.ToString
            dataDaSituacao = Now.ToString

            sql = "UPDATE Cliente SET " & _
            "Nome='" & nome & "'," & _
            "DataDeNascimento='" & dataDeNascimento & "'," & _
            "CidadeDeNascimento='" & cidadeDeNascimento & "'," & _
            "UFDeNascimento='" & UFDeNascimento & "'," & _
            "RG='" & rg & "'," & _
            "CPF='" & cpf & "'," & _
            "Sexo='" & sexo & "'," & _
            "TelefoneResidencial='" & telefoneResidencial & "'," & _
            "TelefoneCelular='" & telefoneCelular & "'," & _
            "OperadoraCelular='" & operadoraCelular & "'," & _
            "Email='" & email & "'," & _
            "Facebook='" & facebook & "'," & _
            "TipoDeLogadouroResidencial='" & tipoDeLogadouroResidencial & "'," & _
            "LogadouroResidencial='" & logadouroResidencial & "'," & _
            "NumeroResidencial='" & numeroResidencial & "'," & _
            "ComplementoResidencial='" & complementoResidencial & "'," & _
            "BairroResidencial='" & bairroResidencial & "'," & _
            "CidadeResidencial='" & cidadeResidencial & "'," & _
            "EstadoResidencial='" & estadoResidencial & "'," & _
            "CEPResidencial='" & cepResidencial & "'," & _
            "DataDeCadastramento='" & dataDeCadastramento & "'," & _
            "DataDaSituacao='" & dataDaSituacao & "', " & _
            "Situacao='" & situacao & "' " & _
            "WHERE Codigo=" & Convert.ToInt64(Txtcdg.Text)
            cn.ConnectionString = My.Settings.banco201302ConnectionString
            cn.Open()
            'montar comando
            cmd.CommandType = CommandType.Text
            cmd.Connection = cn
            cmd.CommandText = sql
            'executar comando
            cmd.ExecuteNonQuery()
            'fechar comando e conexão
            cmd.Dispose()
            cmd.Dispose()
            MessageBox.Show("Cliente alterado com sucesso...", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ToolStrip2.Visible = False
            ToolStrip1.Visible = True
            Desabilita()
            Limpa_total()

        ElseIf acao = "EXCLUIR" Then
            situacao = "D"
            dataDaSituacao = Now.ToString
            ' montar sql
            sql = "UPDATE Cliente SET " & _
            "DataDaSituacao='" & dataDaSituacao & "', " & _
            "Situacao='" & situacao & "' " & _
            "WHERE Codigo= " & Convert.ToInt64(Txtcdg.Text)
            'abrir conexão
            cn.ConnectionString = My.Settings.banco201302ConnectionString
            cn.Open()
            'montar comando
            cmd.CommandType = CommandType.Text
            cmd.Connection = cn
            cmd.CommandText = sql
            'executar comando
            cmd.ExecuteNonQuery()
            'fechar comando e conexão
            cmd.Dispose()
            cmd.Dispose()
            MessageBox.Show("Cliente excluído com sucesso...", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Tstbtngravar.Text = "Gravar"
            Limpa_total()
            ToolStrip2.Visible = False
            ToolStrip1.Visible = True
            Limpa_total()
            Desabilita()
        End If
    End Sub
End Class