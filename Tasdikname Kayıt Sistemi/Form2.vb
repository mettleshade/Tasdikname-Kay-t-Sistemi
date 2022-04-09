Imports System.IO

Public Class Form2

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If okulno.Text = "" Then
            MsgBox("Lütfen Sınıf/Şube Giriniz.", MsgBoxStyle.Critical, "Dikkat !")
        Else
            ListBox1.Items.Add(okulno.Text)
            Using sw As StreamWriter = File.AppendText(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
                sw.WriteLine(okulno.Text)
            End Using
            okulno.Text = ""
        End If
    End Sub

    Private Sub Form2_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        Form1.Timer1.Start()
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        okulno.Text = ListBox1.SelectedItem
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        ListBox1.Items.Remove(okulno.Text)

        Using sw As StreamWriter = File.CreateText(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
            Dim i As Integer
            For i = 1 To ListBox1.Items.Count
                sw.WriteLine(ListBox1.Items(i - 1))
            Next
        End Using
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        File.Delete(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
        My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath.ToString() & "\sinifsube", String.Empty, False)
        ListBox1.Items.Clear()
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ListBox1.Items.Clear()
        Dim r1 As IO.StreamReader
        r1 = New IO.StreamReader(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
        While (r1.Peek() > -1)
            ListBox1.Items.Add(r1.ReadLine)
        End While
        r1.Close()
    End Sub

    Private Sub okulno_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles okulno.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            If okulno.Text = "" Then
                MsgBox("Lütfen Sınıf/Şube Giriniz.", MsgBoxStyle.Critical, "Dikkat !")
            Else
                ListBox1.Items.Add(okulno.Text)
                Using sw As StreamWriter = File.AppendText(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
                    sw.WriteLine(okulno.Text)
                End Using
                okulno.Text = ""
            End If
        End If
    End Sub

End Class