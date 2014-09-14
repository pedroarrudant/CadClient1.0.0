Public Class Form1

    Private Sub CadastrosToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CadastrosToolStripMenuItem.Click
        Dim frm As New Form2
        frm.MdiParent = Me
        frm.Show()
        CadastrosToolStripMenuItem.Checked = True
        CadastrosToolStripMenuItem.Enabled = False
    End Sub
End Class
