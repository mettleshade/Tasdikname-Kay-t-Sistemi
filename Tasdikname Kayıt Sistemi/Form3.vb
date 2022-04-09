Public Class Form3

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FolderBrowserDialog1.ShowDialog()
        TextBox1.Text = FolderBrowserDialog1.SelectedPath & "\"
        My.Settings.yedekyer = FolderBrowserDialog1.SelectedPath & "\"
        My.Settings.Save()
        Me.Close()
    End Sub

End Class