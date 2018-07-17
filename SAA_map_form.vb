Public Class SAA_map_form
    Private Sub SAAW_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For a = 0 To mapgroup.length - 1
            ComboBox1.Items.Add(mapgroup.getattribute(mapgroup.getid(a), 0))
        Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex <> -1 Then
            PictureBox1.Load(Application.StartupPath + "\data\map\" & Trim(ComboBox1.Text) & ".jpg")
            ComboBox2.Items.Clear()
            Button1.Text = "选择战斗点"
            Button2.Text = "选择战斗点"

            For a = 0 To mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.length - 1
                Dim id As String = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getid(a)
                Dim showstring As String = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 0)
                showstring = showstring & "(" & mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 2) & "|" & mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 3) & ")"
                If mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 4) > 0 Then
                    showstring = showstring & "[Max]"
                End If
                If mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 5) > 0 Then
                    showstring = showstring & "[Boss]"
                End If
                ComboBox2.Items.Add(showstring)
            Next
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim id As String = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getid(ComboBox2.SelectedIndex)
        Button1.Text = "优势：" & mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 2)
        Button2.Text = "确保：" & mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 3)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text <> "选择战斗点" Then
            Dim id As String = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getid(ComboBox2.SelectedIndex)
            If leading_intarget = 1 Then
                SAA_plane_form.TextBox1.Text = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 2)
            ElseIf leading_intarget = 2 Then
                SAA_plane_form.TextBox2.Text = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 2)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text <> "选择战斗点" Then
            Dim id As String = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getid(ComboBox2.SelectedIndex)
            If leading_intarget = 1 Then
                SAA_plane_form.TextBox1.Text = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 3)
            ElseIf leading_intarget = 2 Then
                SAA_plane_form.TextBox2.Text = mapgroup.getdata(Trim(ComboBox1.Text)).getdatagroup.getattribute(id, 3)
            End If
            Me.Close()
        End If
    End Sub

    Private Sub SAAW_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        saa_plane.Show()
    End Sub
End Class