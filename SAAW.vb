Structure mapdata
    Public name As String
    Public battle() As String
    Public max() As Integer
    Public boss() As Integer
    Public advantage() As Integer
    Public ensured() As Integer
End Structure



Public Class SAAW
    Dim map() As mapdata
    Dim mapcount As Integer

    Private Sub SAAW_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim saamd As New Xml.XmlDocument
        saamd.Load(Application.StartupPath + "\data\SAAmapdata.xml")

        Dim battlecount As Integer = 0

        For Each mdnode As Xml.XmlNode In saamd.DocumentElement.ChildNodes
            ReDim Preserve map(mapcount)
            For a = 0 To 25
                ReDim map(mapcount).battle(25)
                ReDim map(mapcount).max(25)
                ReDim map(mapcount).boss(25)
                ReDim map(mapcount).advantage(25)
                ReDim map(mapcount).ensured(25)
            Next

            map(mapcount).name = mdnode.Attributes("name").Value

            For Each bnode As Xml.XmlNode In mdnode.ChildNodes
                map(mapcount).battle(battlecount) = bnode.Attributes("name").Value
                map(mapcount).max(battlecount) = bnode.Attributes("max").Value
                map(mapcount).boss(battlecount) = bnode.Attributes("boss").Value

                For Each asnode As Xml.XmlNode In bnode.ChildNodes
                    If asnode.Attributes("state").Value = "advantage" Then
                        map(mapcount).advantage(battlecount) = asnode.InnerText
                    ElseIf asnode.Attributes("state").Value = "ensured" Then
                        map(mapcount).ensured(battlecount) = asnode.InnerText
                    End If
                Next

                battlecount = battlecount + 1
            Next

            mapcount = mapcount + 1
            battlecount = 0
        Next

        For a = 0 To mapcount - 1
            ComboBox1.Items.Add(map(a).name)
        Next

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex <> -1 Then
            PictureBox1.Load(Application.StartupPath + "\data\map\" & map(ComboBox1.SelectedIndex).name & ".jpg")
            ComboBox2.Items.Clear()
            Button1.Text = "选择战斗点"
            Button2.Text = "选择战斗点"

            Dim battlecount As Integer = 0
            Dim battlename As String
            Do While map(ComboBox1.SelectedIndex).battle(battlecount) <> ""
                battlename = ""
                battlename = map(ComboBox1.SelectedIndex).battle(battlecount)

                battlename = battlename & "(" & map(ComboBox1.SelectedIndex).advantage(battlecount) & "│"
                battlename = battlename & map(ComboBox1.SelectedIndex).ensured(battlecount) & ")"

                If map(ComboBox1.SelectedIndex).max(battlecount) = 1 Then
                    battlename = battlename & "[Max]"
                End If
                If map(ComboBox1.SelectedIndex).boss(battlecount) = 1 Then
                    battlename = battlename & "[Boss]"
                End If


                ComboBox2.Items.Add(battlename)
                battlecount = battlecount + 1
            Loop
        End If

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Button1.Text = "优势：" & map(ComboBox1.SelectedIndex).advantage(ComboBox2.SelectedIndex)
        Button2.Text = "确保：" & map(ComboBox1.SelectedIndex).ensured(ComboBox2.SelectedIndex)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text <> "选择战斗点" Then
            If leading_intarget = 1 Then
                SAAM.TextBox1.Text = map(ComboBox1.SelectedIndex).advantage(ComboBox2.SelectedIndex)
            ElseIf leading_intarget = 2 Then
                SAAM.TextBox2.Text = map(ComboBox1.SelectedIndex).advantage(ComboBox2.SelectedIndex)
            End If
            Me.Close()
            End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text <> "选择战斗点" Then
            If leading_intarget = 1 Then
                SAAM.TextBox1.Text = map(ComboBox1.SelectedIndex).ensured(ComboBox2.SelectedIndex)
            ElseIf leading_intarget = 2 Then
                SAAM.TextBox2.Text = map(ComboBox1.SelectedIndex).ensured(ComboBox2.SelectedIndex)
            End If
            Me.Close()
            End If
    End Sub

    Private Sub SAAW_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        SAAM.saam.Show()
    End Sub
End Class