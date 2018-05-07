Public Class SAAS
    Private Sub SAAS_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For a = 1 To 15
            ComboBox1.Items.Add(a)
        Next
        ComboBox1.SelectedIndex = 0
        For a = 0 To 10
            ComboBox2.Items.Add(a)
        Next
        ComboBox2.SelectedIndex = 0
        For a = 0 To 1
            ComboBox3.Items.Add("列表" & a + 1)
        Next
        ComboBox3.SelectedIndex = 0




        For a = 1 To bpcolle.Count
            bplane = bpcolle(a)
            'ListBox1.Items.Add("(+" & bplane.antiaircraft & ")" & bplane.name)
            ListBox1.Items.Add(bplane.name)
        Next


        Call list2refresh()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call addequip(ListBox1.SelectedIndex + 1, ComboBox2.Items(ComboBox2.SelectedIndex), ComboBox1.Text)
        Call list2refresh()
    End Sub

    Private Sub list2refresh()
        ListBox2.Items.Clear()
        Dim updrootnode As Xml.XmlNode = saaupd.DocumentElement
        Try
            For Each upnode As Xml.XmlNode In saaupd.DocumentElement.ChildNodes
                If ComboBox3.SelectedIndex = 0 Then
                    If upnode.Attributes("amountlist1").Value > 0 Then
                        ListBox2.Items.Add("(" & Format(Val(upnode.Attributes("amountlist1").Value), "00") & ")" & upnode.Attributes("name").Value)
                    End If
                ElseIf ComboBox3.SelectedIndex = 1 Then
                    If upnode.Attributes("amountlist2").Value > 0 Then
                        ListBox2.Items.Add("(" & Format(Val(upnode.Attributes("amountlist2").Value), "00") & ")" & upnode.Attributes("name").Value)
                    End If
                End If
            Next
        Catch
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim selectvalue As Integer = ListBox2.SelectedIndex
        If selectvalue <> -1 Then
            uplane.name = Mid(ListBox2.Items(ListBox2.SelectedIndex).ToString, 5, Len(ListBox2.Items(ListBox2.SelectedIndex).ToString) - 4)
            Dim updrootnode As Xml.XmlNode = saaupd.DocumentElement
            For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
                If upnode.Attributes("name").Value = uplane.name Then

                    If ComboBox3.SelectedIndex = 0 Then
                        upnode.SetAttribute("amountlist1", upnode.Attributes("amountlist1").Value + 1)
                        If upnode.Attributes("amountlist1").Value > 99 Then
                            upnode.SetAttribute("amountlist1", 99)
                        End If
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                        Call list2refresh()
                        ListBox2.SelectedIndex = selectvalue
                    ElseIf ComboBox3.SelectedIndex = 1 Then
                        upnode.SetAttribute("amountlist2", upnode.Attributes("amountlist2").Value + 1)
                        If upnode.Attributes("amountlist2").Value > 99 Then
                            upnode.SetAttribute("amountlist2", 99)
                        End If
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                        Call list2refresh()
                        ListBox2.SelectedIndex = selectvalue
                    End If
                End If
            Next
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim selectvalue As Integer = ListBox2.SelectedIndex
        If selectvalue <> -1 Then
            uplane.name = Mid(ListBox2.Items(ListBox2.SelectedIndex).ToString, 5, Len(ListBox2.Items(ListBox2.SelectedIndex).ToString) - 4)
            Dim updrootnode As Xml.XmlNode = saaupd.DocumentElement
            For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
                If upnode.Attributes("name").Value = uplane.name Then
                    If ComboBox3.SelectedIndex = 0 Then
                        upnode.SetAttribute("amountlist1", upnode.Attributes("amountlist1").Value - 1)
                        If upnode.Attributes("amountlist1").Value = 0 Then
                            selectvalue = selectvalue - 1
                        End If
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                        Call list2refresh()
                        ListBox2.SelectedIndex = selectvalue
                    ElseIf ComboBox3.SelectedIndex = 1 Then
                        upnode.SetAttribute("amountlist2", upnode.Attributes("amountlist2").Value - 1)
                        If upnode.Attributes("amountlist2").Value = 0 Then
                            selectvalue = selectvalue - 1
                        End If
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                        Call list2refresh()
                        ListBox2.SelectedIndex = selectvalue
                    End If

                End If

            Next
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Call list2refresh()
    End Sub

    Private Sub SAAS_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If SAAM.saam.ComboBox3.SelectedIndex = 0 Then
            SAAM.saam.ComboBox3.SelectedIndex = 1
            SAAM.saam.ComboBox3.SelectedIndex = 0
        ElseIf SAAM.saam.ComboBox3.SelectedIndex = 1 Then
            SAAM.saam.ComboBox3.SelectedIndex = 0
            SAAM.saam.ComboBox3.SelectedIndex = 1
        End If
        SAAM.saam.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim result As DialogResult
        result = MessageBox.Show("导入操作将清除此列表中所有已登录的装备，是否继续", "警告", MessageBoxButtons.YesNo)

        If result = DialogResult.Yes Then
            Dim linestring() As String
            OpenFileDialog1.InitialDirectory = Application.StartupPath
            OpenFileDialog1.Filter = "CSV文件|*.csv"
            OpenFileDialog1.FilterIndex = 1
            OpenFileDialog1.RestoreDirectory = False
            If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
                Dim count As Integer = 0
                Try
                    Dim sr As System.IO.StreamReader = New System.IO.StreamReader(OpenFileDialog1.FileName)
                    Dim line As String
                    line = sr.ReadLine
                    Do While line IsNot Nothing
                        ReDim Preserve linestring(count)
                        linestring(count) = line
                        count = count + 1
                        line = sr.ReadLine
                    Loop
                    sr.Close()
                Catch
                    MessageBox.Show("读取文件错误")
                    Exit Sub
                End Try

                For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
                    If ComboBox3.SelectedIndex = 0 Then
                        upnode.SetAttribute("amountlist1", 0)
                    ElseIf ComboBox3.SelectedIndex = 1 Then
                        upnode.SetAttribute("amountlist2", 0)
                    End If
                Next

                Dim commacount As Integer
                Dim idstart As Integer
                Dim idlength As Integer
                Dim improvestart As Integer
                Dim improvelength As Integer
                Dim lockstart As Integer
                Dim equipid As Integer
                Dim equipimprove As Integer
                Dim equiplock As Integer

                If count <> 0 Then
                    For a = 1 To count - 1
                        commacount = 0
                        For b = 1 To Len(linestring(a))
                            If Mid(linestring(a), b, 1) = "," Then
                                If commacount = 0 Then
                                    idstart = b + 1
                                ElseIf commacount = 1 Then
                                    idlength = b - idstart
                                ElseIf commacount = 2 Then
                                    improvestart = b + 1
                                ElseIf commacount = 3 Then
                                    improvelength = b - improvestart
                                ElseIf commacount = 4 Then
                                    lockstart = b + 1
                                End If
                                commacount = commacount + 1
                            End If
                        Next

                        equipid = Mid(linestring(a), idstart, idlength)
                        equipimprove = Mid(linestring(a), improvestart, improvelength)
                        equiplock = Mid(linestring(a), lockstart, 1)

                        If bpcolle.Contains(equipid) Then
                            For b = 1 To bpcolle.Count
                                If bpcolle(b).id = equipid Then
                                    If Not （CheckBox1.Checked And equiplock = 0） Then
                                        Call addequip(b, equipimprove, 1)
                                    End If
                                End If
                            Next
                        End If
                    Next
                    Call list2refresh()
                End If
            End If
        End If
    End Sub

    Private Sub addequip(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
        bplane = bpcolle(x)
        uplane.name = bplane.name & "(★" & y & ")"
        uplane.improve = y
        uplane.classification = bplane.classification
        If uplane.classification = 1 Or uplane.classification = 11 Then
            uplane.antiaircraft = bplane.antiaircraft + uplane.improve * 0.2
        ElseIf uplane.classification = 2 Then
            uplane.antiaircraft = bplane.antiaircraft + uplane.improve * 0.25
        Else
            uplane.antiaircraft = bplane.antiaircraft
        End If
        If ComboBox3.SelectedIndex = 0 Then
            uplane.amountlist1 = z
            uplane.amountlist2 = 0
            If uplane.amountlist1 > 99 Then
                uplane.amountlist1 = 99
            End If
        ElseIf ComboBox3.SelectedIndex = 1 Then
            uplane.amountlist2 = z
            uplane.amountlist1 = 0
            If uplane.amountlist2 > 99 Then
                uplane.amountlist2 = 99
            End If
        End If
        uplane.id = bplane.id
        uplane.fire = bplane.fire
        uplane.torpedo = bplane.torpedo
        uplane.antisubmarine = bplane.antisubmarine
        uplane.bombing = bplane.bombing
        uplane.hit = bplane.hit
        uplane.armor = bplane.armor
        uplane.avoid = bplane.avoid
        uplane.spotting = bplane.spotting
        uplane.airrange = bplane.airrange
        uplane.antibombing = bplane.antibombing
        uplane.intercept = bplane.intercept
        uplane.nightfighting = bplane.nightfighting
        uplane.uniquecode = 0



        Dim insertsucceed As Boolean = False




        Dim updrootnode As Xml.XmlNode = saaupd.DocumentElement

        Dim upnodelist As Xml.XmlNodeList = updrootnode.ChildNodes
        If upnodelist.Count = 0 Then
            Dim upnode As Xml.XmlElement = saaupd.CreateElement("plane")
            upnode.SetAttribute("name", uplane.name)
            upnode.SetAttribute("id", uplane.id)
            upnode.SetAttribute("classification", uplane.classification)
            upnode.SetAttribute("fire", uplane.fire)
            upnode.SetAttribute("torpedo", uplane.torpedo)
            upnode.SetAttribute("antiaircraft", uplane.antiaircraft)
            upnode.SetAttribute("antisubmarine", uplane.antisubmarine)
            upnode.SetAttribute("bombing", uplane.bombing)
            upnode.SetAttribute("hit", uplane.hit)
            upnode.SetAttribute("armor", uplane.armor)
            upnode.SetAttribute("avoid", uplane.avoid)
            upnode.SetAttribute("spotting", uplane.spotting)
            upnode.SetAttribute("airrange", uplane.airrange)
            upnode.SetAttribute("antibombing", uplane.antibombing)
            upnode.SetAttribute("intercept", uplane.intercept)
            upnode.SetAttribute("nightfighting", uplane.nightfighting)
            upnode.SetAttribute("improve", uplane.improve)
            upnode.SetAttribute("amountlist1", uplane.amountlist1)
            upnode.SetAttribute("amountlist2", uplane.amountlist2)
            upnode.SetAttribute("uniquecode", uplane.uniquecode)

            updrootnode.AppendChild(upnode)

            insertsucceed = True

            saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")

        Else
            For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
                If upnode.Attributes("name").Value = uplane.name Then
                    If uplane.amountlist1 <> 0 Then
                        upnode.SetAttribute("amountlist1", upnode.Attributes("amountlist1").Value + uplane.amountlist1)
                        If upnode.Attributes("amountlist1").Value > 99 Then
                            upnode.SetAttribute("amountlist1", 99)
                        End If
                        insertsucceed = True
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                    ElseIf uplane.amountlist2 <> 0 Then
                        upnode.SetAttribute("amountlist2", upnode.Attributes("amountlist2").Value + uplane.amountlist2)
                        If upnode.Attributes("amountlist2").Value > 99 Then
                            upnode.SetAttribute("amountlist2", 99)
                        End If
                        insertsucceed = True
                        saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
                    End If
                End If


            Next


            If insertsucceed = False Then

                Dim plane As Xml.XmlElement = saaupd.CreateElement("plane")
                plane.SetAttribute("name", uplane.name)
                plane.SetAttribute("id", uplane.id)
                plane.SetAttribute("classification", uplane.classification)
                plane.SetAttribute("fire", uplane.fire)
                plane.SetAttribute("torpedo", uplane.torpedo)
                plane.SetAttribute("antiaircraft", uplane.antiaircraft)
                plane.SetAttribute("antisubmarine", uplane.antisubmarine)
                plane.SetAttribute("bombing", uplane.bombing)
                plane.SetAttribute("hit", uplane.hit)
                plane.SetAttribute("armor", uplane.armor)
                plane.SetAttribute("avoid", uplane.avoid)
                plane.SetAttribute("spotting", uplane.spotting)
                plane.SetAttribute("airrange", uplane.airrange)
                plane.SetAttribute("antibombing", uplane.antibombing)
                plane.SetAttribute("intercept", uplane.intercept)
                plane.SetAttribute("nightfighting", uplane.nightfighting)
                plane.SetAttribute("improve", uplane.improve)
                plane.SetAttribute("amountlist1", uplane.amountlist1)
                plane.SetAttribute("amountlist2", uplane.amountlist2)
                plane.SetAttribute("uniquecode", uplane.uniquecode)

                updrootnode.AppendChild(plane)

                insertsucceed = True

                saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
            End If
        End If
    End Sub
End Class