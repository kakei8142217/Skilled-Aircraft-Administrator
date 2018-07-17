Public Class SAA_setting_form
    Dim idlist As New Collection

    Private Sub SAA_setting_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For a = 1 To 15
            ComboBox1.Items.Add(a)
        Next
        ComboBox1.SelectedIndex = 0
        For a = 0 To 10
            ComboBox2.Items.Add(a)
        Next
        ComboBox2.SelectedIndex = 0

        If setting_mode = 1 Then
            For a = 0 To 1
                ComboBox3.Items.Add("列表" & a + 1)
            Next
            ComboBox3.SelectedIndex = 0

            Dim cvequipmentcount As Integer = 0
            Do While cvequipmentgroup.getid(cvequipmentcount) <> 0
                ListBox1.Items.Add(cvequipmentgroup.getattribute(cvequipmentgroup.getid(cvequipmentcount), 1))
                cvequipmentcount = cvequipmentcount + 1
            Loop

        ElseIf setting_mode = 2 Then
            For a = 0 To 1
                ComboBox3.Items.Add("列表" & a + 1)
            Next
            ComboBox3.SelectedIndex = 0

            Dim ddequipmentcount As Integer = 0
            Do While ddequipmentgroup.getid(ddequipmentcount) <> 0
                ListBox1.Items.Add(ddequipmentgroup.getattribute(ddequipmentgroup.getid(ddequipmentcount), 1))
                ddequipmentcount = ddequipmentcount + 1
            Loop

        End If

        Call list2refresh()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If setting_mode = 1 Then
            Dim id As Integer = cvequipmentgroup.getattribute(cvequipmentgroup.getid(ListBox1.SelectedIndex), 0) * 100 + ComboBox2.Text
            Call editequipment(id, ComboBox1.Text， ComboBox3.SelectedIndex)
        ElseIf setting_mode = 2 Then
            Dim id As Integer = ddequipmentgroup.getattribute(ddequipmentgroup.getid(ListBox1.SelectedIndex), 0) * 100 + ComboBox2.Text
            Call editequipment(id, ComboBox1.Text， ComboBox3.SelectedIndex)
        End If
        Call list2refresh()
    End Sub

    Private Sub list2refresh()
        Dim selectedindex As Integer = ListBox2.SelectedIndex
        ListBox2.Items.Clear()
        idlist.Clear()
        If setting_mode = 1 Then
            If usercvequipmentgroup IsNot Nothing Then
                For a = 0 To usercvequipmentgroup.length - 1
                    If usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), ComboBox3.SelectedIndex + 3) <> 0 Then
                        Dim itemstring As String = "[" & Format(Val(usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), ComboBox3.SelectedIndex + 3)), "00") & "]"
                        itemstring = itemstring & usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), 1)
                        ListBox2.Items.Add(itemstring)
                        idlist.Add(usercvequipmentgroup.getid(a))
                    End If
                Next
            End If
        ElseIf setting_mode = 2 Then
            If userddequipmentgroup IsNot Nothing Then
                For a = 0 To userddequipmentgroup.length - 1
                    If userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), ComboBox3.SelectedIndex + 3) <> 0 Then
                        Dim itemstring As String = "[" & Format(Val(userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), ComboBox3.SelectedIndex + 3)), "00") & "]"
                        itemstring = itemstring & userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), 1)
                        ListBox2.Items.Add(itemstring)
                        idlist.Add(userddequipmentgroup.getid(a))
                    End If
                Next
            End If
        End If
        If ListBox2.Items.Count <> 0 Then
            If selectedindex >= ListBox2.Items.Count Then selectedindex = selectedindex - 1
            ListBox2.SelectedIndex = selectedindex
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim id As Integer = idlist(ListBox2.SelectedIndex + 1)
        Call editequipment(id, 1， ComboBox3.SelectedIndex)
        Call list2refresh()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim id As Integer = idlist(ListBox2.SelectedIndex + 1)
        Call editequipment(id, -1， ComboBox3.SelectedIndex)
        Call list2refresh()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Call list2refresh()
    End Sub

    Private Sub SAA_setting_form_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        If setting_mode = 1 Then
            plane.loadplaneunit(saa_plane.ComboBox3.SelectedIndex)
            saa_plane.Show()
        ElseIf setting_mode = 2 Then
            saa_ddrank.Show()
            smallgun.loadsmallgunnuit(saa_plane.ComboBox3.SelectedIndex)
        End If
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

                If setting_mode = 1 Then
                    For Each node As Xml.XmlElement In saa_u_pl_emd.DocumentElement.ChildNodes
                        If ComboBox3.SelectedIndex = 0 Then
                            node.SetAttribute("amountlist1", 0)
                        ElseIf ComboBox3.SelectedIndex = 1 Then
                            node.SetAttribute("amountlist2", 0)
                        End If
                    Next
                    If usercvequipmentgroup IsNot Nothing Then
                        usercvequipmentgroup.clear()
                    End If
                ElseIf setting_mode = 2 Then
                    For Each node As Xml.XmlElement In saa_u_sg_emd.DocumentElement.ChildNodes
                        If ComboBox3.SelectedIndex = 0 Then
                            node.SetAttribute("amountlist1", 0)
                        ElseIf ComboBox3.SelectedIndex = 1 Then
                            node.SetAttribute("amountlist2", 0)
                        End If
                    Next
                    If userddequipmentgroup IsNot Nothing Then
                        userddequipmentgroup.clear()
                    End If
                End If

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

                        If setting_mode = 1 And cvequipmentgroup.exist(equipid) Then
                            If Not （CheckBox1.Checked And equiplock = 0） Then
                                Call editequipment(equipid * 100 + equipimprove, 1, ComboBox3.SelectedIndex)
                            End If
                        ElseIf setting_mode = 2 And ddequipmentgroup.exist(equipid) Then
                            If Not （CheckBox1.Checked And equiplock = 0） Then
                                Call editequipment(equipid * 100 + equipimprove, 1, ComboBox3.SelectedIndex)
                            End If
                        End If
                    Next
                    Call list2refresh()
                End If
            End If
        End If
    End Sub

    Private Sub editequipment(ByVal id As String, ByVal amount As Integer, ByVal list As Integer)
        If setting_mode = 1 Then
            Dim usercvequipmentdata As New basedata_class(6)
            If usercvequipmentgroup Is Nothing Then
                usercvequipmentgroup = New basedatagroup_class(6)
            End If
            If usercvequipmentgroup.exist(Trim(Str(id))) Then
                usercvequipmentdata = usercvequipmentgroup.getdata(Trim(Str(id)))
            Else
                usercvequipmentdata.setattribute(0, id)
                Dim baseid As Integer = Int(id / 100)
                Dim improve As Integer = id - Int(id / 100) * 100
                Dim name As String = cvequipmentgroup.getattribute(baseid, 1) & "(★" & improve & ")"
                usercvequipmentdata.setattribute(1, name)
                usercvequipmentdata.setattribute(2, improve)
                usercvequipmentdata.setattribute(3, 0)
                usercvequipmentdata.setattribute(4, 0)
                Dim antiaircraft As Double
                If cvequipmentgroup.getattribute(baseid, 2) = 1 Or cvequipmentgroup.getattribute(baseid, 2) = 1 Then
                    antiaircraft = baseequipmentgroup.getattribute(baseid, 4) + improve * 0.2
                ElseIf cvequipmentgroup.getattribute(baseid, 2) = 2 Then
                    antiaircraft = baseequipmentgroup.getattribute(baseid, 4) + improve * 0.25
                Else
                    antiaircraft = baseequipmentgroup.getattribute(baseid, 4)
                End If
                usercvequipmentdata.setattribute(5, antiaircraft)
            End If

            usercvequipmentdata.setattribute(list + 3, usercvequipmentdata.getattribute(list + 3) + amount)
            usercvequipmentgroup.setdata(usercvequipmentdata)

            Dim insert As Boolean = False
            If saa_u_pl_emd.DocumentElement.ChildNodes.Count <> 0 Then
                For Each node As Xml.XmlElement In saa_u_pl_emd.DocumentElement.ChildNodes
                    If node.Attributes("id").Value = usercvequipmentdata.getattribute(0) Then
                        node.SetAttribute("amountlist1", usercvequipmentdata.getattribute(3))
                        node.SetAttribute("amountlist2", usercvequipmentdata.getattribute(4))
                        insert = True
                        Exit For
                    End If
                Next
            End If
            If insert = False Then
                Dim node As Xml.XmlElement = saa_u_pl_emd.CreateElement("plane")
                Dim attname() As String = {"id", "name", "improve", "amountlist1", "amountlist2", "antiaircraft"}
                For a = 0 To attname.Length - 1
                    node.SetAttribute(attname(a), usercvequipmentdata.getattribute(a))
                Next
                saa_u_pl_emd.DocumentElement.AppendChild(node)
            End If
            saa_u_pl_emd.Save(Application.StartupPath + "\data\user\SAA_u_pl_emd.xml")
        ElseIf setting_mode = 2 Then
            Dim userddequipmentdata As New basedata_class(7)
            If userddequipmentgroup Is Nothing Then
                userddequipmentgroup = New basedatagroup_class(7)
            End If
            If userddequipmentgroup.exist(Trim(Str(id))) Then
                userddequipmentdata = userddequipmentgroup.getdata(Trim(Str(id)))
            Else
                userddequipmentdata.setattribute(0, id)
                Dim baseid As Integer = Int(id / 100)
                Dim improve As Integer = id - Int(id / 100) * 100
                Dim name As String = ddequipmentgroup.getattribute(baseid, 1) & "(★" & improve & ")"
                userddequipmentdata.setattribute(1, name)
                userddequipmentdata.setattribute(2, improve)
                userddequipmentdata.setattribute(3, 0)
                userddequipmentdata.setattribute(4, 0)
                Dim fire As Double
                If ddequipmentgroup.getattribute(baseid, 2) = 1 Then
                    fire = baseequipmentgroup.getattribute(baseid, 2) + Math.Sqrt(improve)
                Else
                    fire = baseequipmentgroup.getattribute(baseid, 2)
                End If
                Dim torpedo As Double
                If ddequipmentgroup.getattribute(baseid, 2) = 2 Then
                    torpedo = baseequipmentgroup.getattribute(baseid, 3) + Math.Sqrt(improve)
                Else
                    torpedo = baseequipmentgroup.getattribute(baseid, 3)
                End If
                userddequipmentdata.setattribute(5, fire)
                userddequipmentdata.setattribute(6, torpedo)
            End If

            userddequipmentdata.setattribute(list + 3, userddequipmentdata.getattribute(list + 3) + amount)
            userddequipmentgroup.setdata(userddequipmentdata)

            Dim insert As Boolean = False
            If saa_u_sg_emd.DocumentElement.ChildNodes.Count <> 0 Then
                For Each node As Xml.XmlElement In saa_u_sg_emd.DocumentElement.ChildNodes
                    If node.Attributes("id").Value = userddequipmentdata.getattribute(0) Then
                        node.SetAttribute("amountlist1", userddequipmentdata.getattribute(3))
                        node.SetAttribute("amountlist2", userddequipmentdata.getattribute(4))
                        insert = True
                        Exit For
                    End If
                Next
            End If
            If insert = False Then
                Dim node As Xml.XmlElement = saa_u_sg_emd.CreateElement("smallgun")
                Dim attname() As String = {"id", "name", "improve", "amountlist1", "amountlist2", "fire", "torpedo"}
                For a = 0 To attname.Length - 1
                    node.SetAttribute(attname(a), userddequipmentdata.getattribute(a))
                Next
                saa_u_sg_emd.DocumentElement.AppendChild(node)
            End If
            saa_u_sg_emd.Save(Application.StartupPath + "\data\user\SAA_u_sg_emd.xml")
        End If

    End Sub

End Class