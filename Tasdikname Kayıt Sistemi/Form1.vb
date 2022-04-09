Imports System.Data
Imports System.Data.OleDb
Imports System.IO

Public Class Form1
    Dim baglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Databasee.mdb")
    Dim kackayitbaglanti As New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Databasee.mdb")

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
        
        If System.IO.File.Exists(My.Application.Info.DirectoryPath.ToString() & "\sinifsube") = True Then
        Else
            My.Computer.FileSystem.WriteAllText(My.Application.Info.DirectoryPath.ToString() & "\sinifsube", String.Empty, False)
        End If

        Form3.TextBox1.Text = My.Settings.yedekyer
        TextBox1.Text = My.Settings.kayit

        kackayit()

        If TextBox2.Text = 0 Then
            no.Text = 1
        Else
            otomatikno()
        End If

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        kaydet()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        ara()
    End Sub

    Private Sub Form1_FormClosed(ByVal sender As System.Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles MyBase.FormClosed
        If MsgBox("Bilgileriniz Yedeklensinmi ?", MsgBoxStyle.YesNo, "Dikkat !") = vbYes Then
            Dim zaman As New Date
            Dim zaman2 As New Date
            Dim uzanti As String
            zaman = DateTime.Today
            uzanti = (".mdb")
            Select Case File.Exists(My.Settings.yedekyer & (zaman) & (uzanti))
                Case True
                    If MsgBox("Bugün Zaten Yedek Alınmış. Üstüne Yazılsınmı ?", MsgBoxStyle.YesNoCancel, "Dikkat !") = vbYes Then
                        File.Delete(My.Settings.yedekyer & zaman & uzanti)
                        FileCopy("Databasee.mdb", My.Settings.yedekyer & Path.GetFileName(zaman) & uzanti)
                        MsgBox("Bilgileriniz " & My.Settings.yedekyer & "  Dosyasının İçinde Kayıt Edilmiştir.", MsgBoxStyle.Information, "Dikkat !")
                    End If
                Case False
                    My.Computer.FileSystem.CreateDirectory(My.Settings.yedekyer)
                    FileCopy("Databasee.mdb", My.Settings.yedekyer & Path.GetFileName(zaman) & uzanti)
                    MsgBox("Bilgileriniz " & My.Settings.yedekyer & " Dosyasının İçinde Kayıt Edilmiştir.", MsgBoxStyle.Information, "Dikkat !")
                Case Else
                    MsgBox("İşlemi Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Information, Me.Text)
            End Select
        Else

        End If
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '    Try
        Dim uzanti As String
        Dim dosyadi As String
        Dim zaman As New Date
        dosyadi = ("Databasee")
        uzanti = (".mdb")
        Dim dosya As New FileInfo("\Databasee.mdb")
        Dim Ac As New OpenFileDialog
        Ac.FileName = vbNullString
        Ac.Filter = "Database(*.mdb*) |*.mdb*"
        Ac.InitialDirectory = (My.Settings.yedekyer)
        If Ac.ShowDialog = Windows.Forms.DialogResult.OK Then
            Select Case File.Exists(Application.StartupPath & "\" & (dosyadi) & (uzanti))
                Case True
                    Kill(Application.StartupPath & "\Databasee.mdb")
                    File.Copy(Ac.FileName, Application.StartupPath & "\" & (dosyadi) & (uzanti))
                    Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                    kackayit()
                    If TextBox2.Text = 0 Then
                        no.Text = 1
                    Else
                        otomatikno()
                    End If
                Case False
                    File.Copy(Ac.FileName, Application.StartupPath & "\" & (dosyadi) & (uzanti))
                    Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                    kackayit()
                    If TextBox2.Text = 0 Then
                        no.Text = 1
                    Else
                        otomatikno()
                    End If
                Case Else
                    MsgBox("İşlemi tekrar deneyiniz.", MsgBoxStyle.Information, Me.Text)
            End Select

        End If
        ' Catch ex As Exception
        '    MsgBox("İşlemi Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, Me.Text)
        'End Try

    End Sub

    Private Sub SınıfŞubeDüzenleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SınıfŞubeDüzenleToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ComboBox1.Items.Clear()
        Dim r1 As IO.StreamReader
        r1 = New IO.StreamReader(My.Application.Info.DirectoryPath.ToString() & "\sinifsube")
        While (r1.Peek() > -1)
            ComboBox1.Items.Add(r1.ReadLine)
        End While
        r1.Close()
        Timer1.Stop()
    End Sub

    Private Sub okulno_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles okulno.KeyPress
        If Not (Char.IsNumber(e.KeyChar) = True) And e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub adi_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles adi.KeyPress
        If Not (Char.IsLetter(e.KeyChar) = True) And e.KeyChar <> ChrW(Keys.Back) And e.KeyChar <> ChrW(Keys.Space) Then
            e.Handled = True
        End If
    End Sub

    Private Sub HakkındaToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HakkındaToolStripMenuItem.Click
        MsgBox("<< Tasdikname Kayıt Sistemi >> " & vbCrLf & vbCrLf & " Versiyon : 2.0.0 <FİNAL>" & vbCrLf & vbCrLf & " Mehmet Paçal" & vbCrLf & vbCrLf & " Mettless Dizayn" & vbCrLf & vbCrLf & " www.facebook.com/Metlesshade" & vbCrLf & vbCrLf & " www.mettlessdizayn.com", MsgBoxStyle.Information, "Hakkında")
    End Sub

    Private Sub DataGridView1_CellMouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        Label14.Text = DataGridView1.Item(5, i).Value
        Dim ii As Integer
        ii = DataGridView1.CurrentRow.Index
        ComboBox2.Text = DataGridView1.Item(11, ii).Value
        Timer2.Start()

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        'Try

        If ComboBox2.Text = "HAYIR" Then
            Dim komut As New OleDbCommand("select * from kayitlar", baglanti)
            Dim dr As OleDb.OleDbDataReader
            baglanti.Open()
            dr = komut.ExecuteReader
            Do While dr.Read

                If dr("TNo") = Label14.Text Then
                    adi.Text = dr("AdiSoyadi")
                    okulno.Text = dr("OkulNo")
                    ComboBox1.Text = dr("SinifSube")
                    no.Text = dr("TNo")
                    dusunce.Text = dr("D")
                    tarih.Text = dr("Tarih")
                    imza.Text = dr("Imza")
                    imzatarih.Text = dr("ImzaTarih")
                    Button3.Visible = True
                    Button4.Visible = True
                    Button1.Enabled = False
                    Button5.Visible = True
                    Button8.Visible = True
                    önizleme.Visible = True
                End If
            Loop
            baglanti.Close()
            Timer2.Stop()
        Else
            Dim komut As New OleDbCommand("select * from kayitlar", baglanti)
            Dim dr As OleDb.OleDbDataReader
            baglanti.Open()
            dr = komut.ExecuteReader
            Do While dr.Read

                If dr("TNo") = Label14.Text Then
                    adi.Text = dr("AdiSoyadi")
                    okulno.Text = dr("OkulNo")
                    ComboBox1.Text = dr("SinifSube")
                    no.Text = dr("TNo")
                    dusunce.Text = dr("D")
                    tarih.Text = dr("Tarih")
                    imza.Text = dr("Imza")
                    imzatarih.Text = dr("ImzaTarih")
                    kayiptarih.Text = dr("KayipTarih")
                    kayipno.Text = dr("KayipNo")
                    ComboBox2.Text = dr("Kayipmi")
                    Button3.Visible = True
                    Button4.Visible = True
                    Button1.Enabled = False
                    Button5.Visible = True
                    Button8.Visible = True
                    önizleme.Visible = True
                End If
            Loop
            baglanti.Close()
            Timer2.Stop()
        End If
        

        ' Catch ex As Exception
        'MsgBox("İşlemi Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, Me.Text)
        'End Try



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If MsgBox(adi.Text & " İsimli Kayıt Silinsinmi ?", MsgBoxStyle.YesNoCancel, "Dikkat !") = vbYes Then
                Dim sql As String
                sql = "delete from kayitlar where AdiSoyadi='" & adi.Text & "'"
                Dim komut As New OleDbCommand
                komut.Connection = baglanti
                komut.CommandText = sql
                baglanti.Open()
                komut.ExecuteNonQuery()
                Label13.Text = "Kayıt Başarıyla Silindi."
                TextBox1.Text = TextBox1.Text - 1
                My.Settings.kayit = TextBox1.Text
                My.Settings.Save()
                baglanti.Close()
                Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                kackayit()
            End If
            
        Catch ex As Exception
            MsgBox("İşlem Başarısız Lütfen Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, "Dikkat !")
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Try
            If ComboBox2.Text = "HAYIR" Then
                Dim sql As String
                sql = "update kayitlar set AdiSoyadi='" & adi.Text & "', OkulNo ='" & okulno.Text & "', SinifSube ='" & ComboBox1.Text & "', Tarih ='" & tarih.Text & "', TNo='" & no.Text & "', D ='" & dusunce.Text & "', ImzaTarih ='" & imzatarih.Text & "', Imza ='" & imza.Text & "', Kayipmi='" & "HAYIR" & "', KayipNo='" & "" & "', KayipTarih='" & "" & "' where TNo='" & Label14.Text & "'"
                Dim komut As New OleDbCommand
                komut.Connection = baglanti
                komut.CommandText = sql
                baglanti.Open()
                komut.ExecuteNonQuery()
                Label13.Text = "Bilgiler başarıyla güncellendi"
                baglanti.Close()
                Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                Timer2.Start()

            Else
                Dim sql As String
                sql = "update kayitlar set AdiSoyadi='" & adi.Text & "', OkulNo ='" & okulno.Text & "', SinifSube ='" & ComboBox1.Text & "', Tarih ='" & tarih.Text & "', TNo='" & no.Text & "', D ='" & dusunce.Text & "', ImzaTarih ='" & imzatarih.Text & "', Imza ='" & imza.Text & "', Kayipmi='" & ComboBox2.Text & "', KayipNo='" & kayipno.Text & "', KayipTarih='" & kayiptarih.Text & "' where TNo='" & Label14.Text & "'"
                Dim komut As New OleDbCommand
                komut.Connection = baglanti
                komut.CommandText = sql
                baglanti.Open()
                komut.ExecuteNonQuery()
                Label13.Text = "Bilgiler başarıyla güncellendi"
                baglanti.Close()
                Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                Timer2.Start()
            End If


        Catch ex As Exception
            MsgBox("İşlem Başarısız Lütfen Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, "Dikkat !")

        End Try


    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        temizle()
    End Sub

    Private Sub temizle()
        adi.Text = ""
        okulno.Text = ""
        ComboBox1.Text = ""

        kackayit()

        If ComboBox4.Text = "Otomatik" Then
            otomatikno()
        Else
            no.Text = no.Text + 1
        End If
        kayiptarih.Text = Today
        ComboBox2.Text = "HAYIR"
        dusunce.Text = ""
        kayipno.Text = ""
        kayiptarih.Text = Today
        Button3.Visible = False
        Button4.Visible = False
        Button5.Visible = False
        Button1.Enabled = True
        Button8.Visible = False
        önizleme.Visible = False
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
    End Sub

    Private Sub tarih_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tarih.ValueChanged
        imzatarih.Text = tarih.Text
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Try
            If MsgBox("Bilgileriniz Yedeklensinmi ?", MsgBoxStyle.YesNoCancel, "Dikkat !") = vbYes Then
                Dim zaman As New Date
                Dim zaman2 As New Date
                Dim uzanti As String
                zaman = DateTime.Today
                uzanti = (".mdb")
                Select Case File.Exists(My.Settings.yedekyer & (zaman) & (uzanti))
                    Case True
                        If MsgBox("Bugün Zaten Yedek Alınmış. Üstüne Yazılsınmı ?", MsgBoxStyle.YesNoCancel, "Dikkat !") = vbYes Then
                            File.Delete(My.Settings.yedekyer & zaman & uzanti)
                            FileCopy("Databasee.mdb", My.Settings.yedekyer & Path.GetFileName(zaman) & uzanti)
                            MsgBox("Bilgileriniz " & My.Settings.yedekyer & "  Dosyasının İçinde Kayıt Edilmiştir.", MsgBoxStyle.Information, "Dikkat !")
                        End If
                    Case False
                        My.Computer.FileSystem.CreateDirectory(My.Settings.yedekyer)
                        FileCopy("Databasee.mdb", My.Settings.yedekyer & Path.GetFileName(zaman) & uzanti)
                        MsgBox("Bilgileriniz " & My.Settings.yedekyer & "  Dosyasının İçinde Kayıt Edilmiştir.", MsgBoxStyle.Information, "Dikkat !")
                    Case Else
                        MsgBox("İşlemi tekrar deneyiniz.", MsgBoxStyle.Information, Me.Text)
                End Select
            End If
        Catch ex As Exception
            MsgBox("İşlem Başarısız Lütfen Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, "Dikkat !")

        End Try
        
    End Sub

    Private Sub YedeklemeAyarlarıToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YedeklemeAyarlarıToolStripMenuItem.Click
        Form3.Show()
    End Sub

    Private Sub otomatikno()
        '' TASDİKNAME OTOMATİK NO 

        baglanti.Open()
        Dim komut As New OleDbCommand("SELECT Max(kayitlar.TNo) AS TNo FROM kayitlar;", baglanti)
        Dim oku As OleDb.OleDbDataReader
        oku = komut.ExecuteReader()
        While oku.Read()
            no.Text = oku("TNo")
        End While
        baglanti.Close()

        no.Text = no.Text + 1

    End Sub
   
    Private Sub kackayit()
        kackayitbaglanti.Open()
        Dim sayi As New OleDbCommand("", kackayitbaglanti)
        sayi.CommandText = "Select count(*) from kayitlar"
        Label19.Text = "Toplam : " & sayi.ExecuteScalar & " Kayıt Var."
        TextBox2.Text = sayi.ExecuteScalar
        kackayitbaglanti.Close()
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        For Each dr As DataGridViewRow In DataGridView1.Rows
            If dr.Cells(9).Value.ToString = "" Then
                dr.DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim myprintresult As DialogResult
        myprintresult = PrintDialog1.ShowDialog()
        If myprintresult = DialogResult.OK Then
            PrintDocument1.Print()
        End If
    End Sub

    Private Sub kaydet()
        Try
            If adi.Text = "" And okulno.Text = "" And no.Text = "" Then
                MsgBox("Lütfen Boş Kısımları Doldurunuz. (Adı,Okul No,Tasdikname No", MsgBoxStyle.Critical, "Dikkat !")

            Else
                If ComboBox2.Text = "EVET" Then
                    Dim Sql, sql2 As String
                    Sql = "insert into [kayitlar](AdiSoyadi,OkulNo,SinifSube,Tarih,TNo,Imza,D,ImzaTarih,KayipTarih,KayipNo,Kayipmi) values ('" & adi.Text & "','" & okulno.Text & "','" & ComboBox1.Text & "','" & tarih.Text & "','" & no.Text & "','" & imza.Text & "','" & dusunce.Text & "','" & imzatarih.Text & "','" & kayiptarih.Text & "','" & kayipno.Text & "','" & ComboBox2.Text & "')"

                    sql2 = "select * from kayitlar"
                    Dim komut As New OleDbCommand(sql2, baglanti)
                    Dim dr As OleDb.OleDbDataReader
                    baglanti.Open()
                    dr = komut.ExecuteReader
                    Do While dr.Read
                        If dr("AdiSoyadi") = adi.Text Then
                            dr.Close()
                            baglanti.Close()
                            Label13.Text = "Tasdikname Daha Önceden Kayıt Edilmiştir."
                            MsgBox("Tasdikname Daha Önceden Kayıt Edilmiştir.", MsgBoxStyle.Critical, "Dikkat !")
                            adi.Text = ""
                            okulno.Text = ""
                            ComboBox1.Text = ""
                            no.Text = ""
                            dusunce.Text = ""
                            If ComboBox4.Text = "Otomatik" Then
                                otomatikno()
                            Else
                                no.Text = no.Text + 1
                            End If

                            Exit Sub

                        End If
                    Loop

                    dr.Close()
                    komut.CommandText = Sql
                    komut.ExecuteNonQuery()
                    baglanti.Close()
                    Label13.Text = "Tasdikname Kayıt Edilmiştir."
                    TextBox1.Text = TextBox1.Text + 1
                    My.Settings.kayit = TextBox1.Text
                    My.Settings.Save()
                    adi.Text = ""
                    okulno.Text = ""
                    ComboBox1.Text = ""
                    dusunce.Text = ""
                    If ComboBox4.Text = "Otomatik" Then
                        otomatikno()
                    Else
                        no.Text = no.Text + 1
                    End If
                    Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                    kackayit()
                Else
                    Dim Sql, sql2 As String
                    Sql = "insert into [kayitlar](AdiSoyadi,OkulNo,SinifSube,Tarih,TNo,Imza,D,ImzaTarih,Kayipmi) values ('" & adi.Text & "','" & okulno.Text & "','" & ComboBox1.Text & "','" & tarih.Text & "','" & no.Text & "','" & imza.Text & "','" & dusunce.Text & "','" & imzatarih.Text & "','" & ComboBox2.Text & "')"

                    sql2 = "select * from kayitlar"
                    Dim komut As New OleDbCommand(sql2, baglanti)
                    Dim dr As OleDb.OleDbDataReader
                    baglanti.Open()
                    dr = komut.ExecuteReader
                    Do While dr.Read
                        If dr("AdiSoyadi") = adi.Text Then
                            dr.Close()
                            baglanti.Close()
                            Label13.Text = "Tasdikname Daha Önceden Kayıt Edilmiştir."
                            MsgBox("Tasdikname Daha Önceden Kayıt Edilmiştir.", MsgBoxStyle.Critical, "Dikkat !")
                            adi.Text = ""
                            okulno.Text = ""
                            ComboBox1.Text = ""
                            no.Text = ""
                            dusunce.Text = ""
                            Exit Sub

                        End If
                    Loop

                    dr.Close()
                    komut.CommandText = Sql
                    komut.ExecuteNonQuery()
                    baglanti.Close()
                    Label13.Text = "Tasdikname Kayıt Edilmiştir."
                    TextBox1.Text = TextBox1.Text + 1
                    My.Settings.kayit = TextBox1.Text
                    My.Settings.Save()
                    adi.Text = ""
                    okulno.Text = ""
                    ComboBox1.Text = ""
                    dusunce.Text = ""
                    If ComboBox4.Text = "Otomatik" Then
                        otomatikno()
                    Else
                        no.Text = no.Text + 1
                    End If
                    Me.KayitlarTableAdapter.Fill(Me.DatabaseeDataSet.kayitlar)
                    kackayit()
                End If


            End If
        Catch ex As Exception
            MsgBox("İşlem Başarısız Lütfen Daha Sonra Tekrar Deneyiniz.", MsgBoxStyle.Critical, "Dikkat !")
        End Try
    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ComboBox1.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            kaydet()
        End If

    End Sub

    Private Sub ara()
        If ComboBox3.Text = "ADI" Then
            Dim dv As DataView = New DataView()
            dv.Table = DatabaseeDataSet.kayitlar
            dv.RowFilter = "[AdiSoyadi] Like '" & TextBox5.Text & "%'"
            DataGridView1.DataSource = dv
        End If
        If ComboBox3.Text = "SOYADI" Then
            Dim dv As DataView = New DataView()
            dv.Table = DatabaseeDataSet.kayitlar
            dv.RowFilter = "[AdiSoyadi] Like '%" & TextBox5.Text & "%'"
            DataGridView1.DataSource = dv
        End If
        If ComboBox3.Text = "T.NO" Then
            Dim dv1 As DataView = New DataView()
            dv1.Table = DatabaseeDataSet.kayitlar
            dv1.RowFilter = "[TNo] Like '" & TextBox5.Text & "%'"
            DataGridView1.DataSource = dv1
        End If
    End Sub

    Private Sub adi_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles adi.TextChanged
        TextBox5.Text = adi.Text
    End Sub

    Private Sub tarih_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles tarih.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            kaydet()
        End If
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim font1 As Font = New Font("Segoe UI", 18, FontStyle.Bold)
        Dim font2 As Font = New Font("Segoe UI", 13, FontStyle.Regular)
        e.Graphics.DrawString("Tasdikname Örnek Belge", font1, Brushes.Black, 250, 50)
        e.Graphics.DrawLine(Pens.Black, 90, 100, 750, 100)
        e.Graphics.DrawLine(Pens.Black, 90, 40, 750, 40)
        e.Graphics.DrawString("Adı Soyadı : " & adi.Text, font2, Brushes.Black, 90, 140)
        e.Graphics.DrawString("Okul No : " & okulno.Text, font2, Brushes.Black, 90, 170)
        e.Graphics.DrawString("Sınıf / Şube : " & ComboBox1.Text, font2, Brushes.Black, 90, 200)
        e.Graphics.DrawString("Tasdikname Tarihi : " & tarih.Text, font2, Brushes.Black, 90, 230)
        e.Graphics.DrawString("Tasdikname No : " & no.Text, font2, Brushes.Black, 90, 260)
        e.Graphics.DrawString("İmza : " & imza.Text, font2, Brushes.Black, 90, 290)
        e.Graphics.DrawString("İmza Tarihi : " & imzatarih.Text, font2, Brushes.Black, 90, 320)
        e.Graphics.DrawString("Düşünce : " & dusunce.Text, font2, Brushes.Black, 90, 350)
        e.Graphics.DrawLine(Pens.Black, 90, 390, 750, 390)
        e.Graphics.DrawString("Kayıp : " & ComboBox2.Text, font2, Brushes.Black, 90, 410)
        e.Graphics.DrawString("Kayıp Tarihi : " & kayiptarih.Text, font2, Brushes.Black, 90, 440)
        e.Graphics.DrawString("Kayıp No : " & kayipno.Text, font2, Brushes.Black, 90, 470)
    End Sub

    Private Sub önizleme_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles önizleme.Click
        Dim preview As New PrintPreviewDialog
        preview.Document = Me.PrintDocument1
        preview.WindowState = FormWindowState.Maximized
        preview.PrintPreviewControl.StartPage = 0
        preview.PrintPreviewControl.Zoom = 1.0
        preview.PrintPreviewControl.Columns = 2
        preview.ShowDialog()
    End Sub
End Class
