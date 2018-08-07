Public Class SAA_ddrank_form
    Dim targetshiptypeid As String
    Dim newshiptypeid As String

    Private Sub SAA_ddrank_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        saa_ddrank = Me

        Dim ddshiptypecount As Integer = 1
        Do While userddshiptype.getbaseid(ddshiptypecount) <> 0
            ComboBox1.Items.Add(userddshiptype.getbaseattribute(userddshiptype.getbaseid(ddshiptypecount), 1))
            ddshiptypecount = ddshiptypecount + 1
        Loop
        ComboBox1.Items.Add（"<点击加载未改造舰娘>"）

        For a = 1 To 165
            ComboBox2.Items.Add(a)
            If a < 100 Then
                ComboBox3.Items.Add(a)
            End If
        Next
        ComboBox2.SelectedIndex = 98

        ComboBox4.Items.Add(1)
        For a = 1 To 16
            ComboBox4.Items.Add(a * 10)
        Next
        ComboBox4.SelectedIndex = 8

        ComboBox8.Items.Add("按等级")
        ComboBox8.Items.Add("按火雷")
        ComboBox8.Items.Add("按舰型")
        ComboBox8.SelectedIndex = 0

        For a = 0 To 1
            ComboBox9.Items.Add("列表" & a + 1)
        Next
        ComboBox9.SelectedIndex = 0

        Call checklist1refresh()

        For a = 0 To 98
            ComboBox10.Items.Add(a + 1)
        Next
        ComboBox10.SelectedIndex = 39



    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        setting_mode = 2
        Dim saa_setting As New SAA_setting_form
        saa_setting.Show()
        saa_setting.Top = Me.Top
        saa_setting.Left = Me.Left
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedIndex >= 0 Or CheckedListBox1.SelectedIndex >= 0 Then
            If ComboBox1.SelectedIndex >= 0 Then
                newshiptypeid = Trim(Str(Val(userddshiptype.getbaseid(ComboBox1.SelectedIndex + 1) * 10000000) + (ComboBox2.Text) * 10000 + (ComboBox3.Text) * 10 - Val(CheckBox1.Checked)))
            ElseIf CheckedListBox1.SelectedIndex >= 0 Then
                newshiptypeid = Trim(Str(Val(userddshiptype.getattribute(targetshiptypeid, 2) * 10000000) + (ComboBox2.Text) * 10000 + (ComboBox3.Text) * 10 - Val(CheckBox1.Checked)))
            End If
            If targetshiptypeid <> "0" Then
                targetshiptypeid = Mid(targetshiptypeid, 1, Len(targetshiptypeid) - 1)
            End If
            userddshiptype.edituserddshiptype(targetshiptypeid, newshiptypeid)
            userddshiptype.loadddshiptypeunit()
            Call checklist1refresh()
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim result As DialogResult
        result = MessageBox.Show("导入操作将清除此列表中所有已登录的舰娘，是否继续", "警告", MessageBoxButtons.YesNo)

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

                saa_u_dd_std.DocumentElement.RemoveAll()
                If userddshiptypegroup IsNot Nothing Then
                    userddshiptypegroup.clear()
                End If

                Dim commacount As Integer
                Dim commalist(6) As Integer
                Dim idstart As Integer
                Dim idlength As Integer

                Dim levelstart As Integer
                Dim levellength As Integer

                Dim enhancedstart As Integer
                Dim enhancedlength As Integer

                Dim luckstart As Integer
                Dim lucklength As Integer

                Dim shiptypeid As Double
                Dim shiptypelevel As Integer
                Dim shiptypeluck As Integer
                Dim shiptypeenhanced As Integer

                If count <> 0 Then
                    For a = 1 To Len(linestring(1))
                        If Mid(linestring(1), a, 1) = "," Then
                            commacount = commacount + 1
                        End If
                    Next
                    If commacount = 66 Then
                        commalist = {2, 3, 4, 25, 26, 55, 56}
                    ElseIf commacount = 63 Then
                        commalist = {2, 3, 4, 23, 24, 52, 53}
                    End If
                    For a = 1 To count - 1
                        commacount = 0
                        For b = 1 To Len(linestring(a))
                            If Mid(linestring(a), b, 1) = "," Then
                                If commacount = commalist(0) Then
                                    idstart = b + 1
                                ElseIf commacount = commalist(1) Then
                                    idlength = b - idstart
                                    levelstart = b + 1
                                ElseIf commacount = commalist(2) Then
                                    levellength = b - levelstart
                                ElseIf commacount = commalist(3) Then
                                    enhancedstart = b + 1
                                ElseIf commacount = commalist(4) Then
                                    enhancedlength = b - enhancedstart
                                ElseIf commacount = commalist(5) Then
                                    luckstart = b + 1
                                ElseIf commacount = commalist(6) Then
                                    lucklength = b - luckstart
                                End If
                                commacount = commacount + 1
                            End If
                        Next

                        shiptypeid = Mid(linestring(a), idstart, idlength)
                        shiptypelevel = Mid(linestring(a), levelstart, levellength)
                        shiptypeenhanced = Mid(linestring(a), enhancedstart, enhancedlength)
                        shiptypeluck = Mid(linestring(a), luckstart, lucklength)
                        Dim enhanced As Integer = 1
                        If shiptypeenhanced = 0 Then
                            enhanced = 0
                        End If


                        If ddshiptypegroup.exist(shiptypeid) And shiptypelevel >= Val(ComboBox4.Text) Then
                            Dim newid As String = Trim(Str(shiptypeid * 10000000 + shiptypelevel * 10000 + shiptypeluck * 10 + enhanced))
                            userddshiptype.edituserddshiptype("0", newid)
                        End If
                    Next
                    userddshiptype.loadddshiptypeunit()
                    Call checklist1refresh()
                End If
            End If
        End If


    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex = ComboBox1.Items.Count - 1 Then
            ComboBox1.Items.Clear()
            userddshiptype.modification = 0
            Dim ddshiptypecount As Integer = 1
            Do While userddshiptype.getbaseid(ddshiptypecount) <> 0
                ComboBox1.Items.Add(userddshiptype.getbaseattribute(userddshiptype.getbaseid(ddshiptypecount), 1))
                ddshiptypecount = ddshiptypecount + 1
            Loop
        ElseIf ComboBox1.SelectedIndex <> -1 Then
            ComboBox2.SelectedIndex = 98
            ComboBox3.SelectedIndex = userddshiptype.getbaseattribute(userddshiptype.getbaseid(ComboBox1.SelectedIndex + 1), 2) - 1
            CheckBox1.Checked = 0
            targetshiptypeid = "0"
            CheckedListBox1.SelectedIndex = -1
            Button2.Text = "登录"
        End If
    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        If CheckedListBox1.SelectedIndex <> -1 Then
            targetshiptypeid = userddshiptype.getid(CheckedListBox1.SelectedIndex)
            ComboBox2.SelectedIndex = userddshiptype.getattribute(targetshiptypeid, 6) - 1
            ComboBox3.SelectedIndex = userddshiptype.getattribute(targetshiptypeid, 7) - 1
            CheckBox1.Checked = userddshiptype.getattribute(targetshiptypeid, 5) * -1
            ComboBox1.SelectedIndex = -1
            Button2.Text = "修改"
        End If
    End Sub

    Private Sub checklist1refresh()
        CheckedListBox1.Items.Clear()
        Dim count As Integer = 0
        Do While sgUIcontrol.checklist1refresh(count) <> "0"
            CheckedListBox1.Items.Add(sgUIcontrol.checklist1refresh(count))
            count = count + 1
        Loop
    End Sub

    Private Sub list1refresh()
        ListBox1.Items.Clear()
        Dim count As Integer = 1
        Do While sgUIcontrol.showlist1(count) <> ""
            ListBox1.Items.Add(sgUIcontrol.showlist1(count))
            count = count + 1
        Loop
    End Sub

    Private Sub SAA_ddrank_form_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        ddshipgroup.removeallship()
        saa_plane.Show()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim count As Integer = 0
        For a = 0 To CheckedListBox1.Items.Count - 1
            If CheckedListBox1.GetItemChecked(a) = True Then
                ddshipgroup.addship(userddshiptype.getid(a))
            End If
        Next
        Call list1refresh()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Button5.Enabled = False
        smallgun.clearcollection(1)
        For a = 0 To ddshipgroup.length - 1
            ddshipgroup.ship(a).removeequipment()
        Next

        Dim gunlist() As String = smallgun.getbasesmallgunlist(1)
        Dim torpedo() As String = smallgun.getbasesmallgunlist(2)
        'Dim radar() As String = smallgun.getbasesmallgunlist(3)

        Dim basesmallgungroup(2) As basedatagroup_class
        Dim basesmallgun As basedata_class
        For a = 0 To 2
            basesmallgungroup(a) = New basedatagroup_class(2)
        Next

        If torpedo IsNot Nothing Then
            For a = 0 To torpedo.Length - 1
                basesmallgun = New basedata_class(2)
                basesmallgun.setattribute(0, a)
                basesmallgun.setattribute(1, torpedo(a))
                basesmallgungroup(1).setdata(basesmallgun)
                basesmallgungroup(2).setdata(basesmallgun)
            Next
        End If
        Dim torpedocount As Integer = 0
        If torpedo IsNot Nothing Then
            torpedocount = torpedo.Length
        End If
        If gunlist IsNot Nothing Then
            For a = 0 To gunlist.Length - 1
                basesmallgun = New basedata_class(2)
                basesmallgun.setattribute(0, a)
                basesmallgun.setattribute(1, gunlist(a))
                basesmallgungroup(0).setdata(basesmallgun)
                basesmallgun = New basedata_class(2)
                basesmallgun.setattribute(0, a + torpedocount - 1 + 1)
                basesmallgun.setattribute(1, gunlist(a))
                basesmallgungroup(2).setdata(basesmallgun)
            Next
        End If


        Dim needgunlist As New Collection
        For b = 0 To ddshipgroup.length - 1
            Dim gunneed As Boolean = False
            For c = 0 To ddextrabuffgroup.length - 1
                Dim buffid As String = ddextrabuffgroup.getid(c)
                If ddextrabuffgroup.getattribute(buffid, 1) = 2 And ddequipmentgroup.getattribute(ddextrabuffgroup.getattribute(buffid, 2), 2) = 1 Then
                    For d = 0 To ddextrabuffgroup.getdata(buffid).getdatagroup.length - 1
                        Dim class2buffid As String = ddextrabuffgroup.getdata(buffid).getdatagroup.getid(d)
                        If userddshiptype.unit(ddshipgroup.ship(b).id).type >= ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 2) And userddshiptype.unit(ddshipgroup.ship(b).id).type <= ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 3) Then
                            gunneed = True
                            If needgunlist.Contains(b) = False Then
                                needgunlist.Add(gunneed, b)
                            End If
                        End If
                    Next
                End If
            Next
        Next

        Dim cacheship As ddship_class
        Dim getsmallgunid As String
        Dim buffidlist() As Integer

        Dim mark() As Integer
        Dim marklimit() As Integer
        Dim groupmark() As Integer

        For a = 0 To ddshipgroup.length - 1
            For b = a To ddshipgroup.length - 1
                Dim nfmode As Integer = ddshipgroup.ship(b).nfmode
                If nfmode = 0 Then
                    If userddshiptype.unit(ddshipgroup.ship(b).id).luck >= Val(ComboBox10.Text) Then
                        nfmode = 2
                    Else
                        nfmode = 1
                    End If
                End If



                cacheship = New ddship_class(ddshipgroup.ship(b).id)
                cacheship.buffmultiple = ddshipgroup.ship(b).buffmultiple
                If nfmode = 1 Then
                    If userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode > 2 Then
                        getsmallgunid = smallgun.getsmallgun(3)
                        If getsmallgunid <> "0" Then
                            cacheship.carry(2).equipmentid = getsmallgunid
                        End If
                        mark = {0, 0}
                        marklimit = {basesmallgungroup(0).length - 1, basesmallgungroup(0).length - 1}
                        groupmark = {0, 0}
                    Else
                        ReDim mark(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                        ReDim marklimit(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                        ReDim groupmark(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                        For c = 0 To mark.Length - 1
                            mark(c) = 0
                            marklimit(c) = basesmallgungroup(0).length - 1
                            groupmark(c) = 0
                        Next
                    End If
                ElseIf nfmode = 2 Then
                    ReDim mark(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                    ReDim marklimit(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                    ReDim groupmark(userddshiptype.unit(ddshipgroup.ship(b).id).gridcount - ddshipgroup.ship(b).restrictmode - 1)
                    For c = 0 To mark.Length - 1
                        mark(c) = 0
                        marklimit(c) = basesmallgungroup(1).length - 1
                        groupmark(c) = 1
                    Next
                    If mark.Length > 2 And needgunlist.Contains(Trim(Str(b))) Then
                        marklimit(mark.Length - 1) = basesmallgungroup(2).length - 1
                        groupmark(mark.Length - 1) = 2
                    End If
                End If



                Dim finish As Boolean = False
                Do While (finish = False)
                    For c = 0 To mark.Length - 1
                        getsmallgunid = smallgun.getsmallgun(Trim(Str(basesmallgungroup(groupmark(c)).getattribute(mark(c), 1))))
                        If getsmallgunid <> "0" Then
                            cacheship.carry(c).equipmentid = getsmallgunid
                            smallgun.addcollection(getsmallgunid)
                        End If
                    Next
                    buffidlist = smallgun.extrabuff(cacheship)
                    For c = 0 To buffidlist.Length - 1
                        cacheship.carry(c).buff = buffidlist(c)
                    Next
                    If cacheship.damage > ddshipgroup.ship(b).damage Then
                        For c = 0 To cacheship.length - 1
                            ddshipgroup.ship(b).carry(c).equipmentid = cacheship.carry(c).equipmentid
                            ddshipgroup.ship(b).carry(c).buff = cacheship.carry(c).buff
                        Next
                    End If
                    smallgun.clearcollection()
                    For c = 0 To mark.Length - 1
                        cacheship.carry(c).equipmentid = "0"
                        cacheship.carry(c).buff = 0
                    Next


                    mark(mark.Length - 1) = mark(mark.Length - 1) + 1
                    For c = 1 To mark.Length
                        If mark(mark.Length - c) > marklimit(mark.Length - c) Then
                            If mark.Length - c <> 0 Then
                                mark(mark.Length - c - 1) = mark(mark.Length - c - 1) + 1
                                For d = mark.Length - c To mark.Length - 1
                                    mark(d) = mark(-1 + d)
                                Next
                            End If
                        End If
                    Next
                    If mark(0) > marklimit(0) Then
                        finish = True
                    End If
                Loop



                If ddshipgroup.ship(b).carry(2).equipmentid <> "0" Then
                    If smallgun.unit(ddshipgroup.ship(b).carry(2).equipmentid).classification = 3 And ddshipgroup.ship(b).carry(2).buff = 0 Then
                        ddshipgroup.ship(b).carry(2).equipmentid = "0"
                        ddshipgroup.ship(b).carry(2).buff = 0
                    End If
                End If

                If userddshiptype.unit(ddshipgroup.ship(b).id).enhanced = 1 And ddshipgroup.ship(b).enhancedlock = 1 Then
                    getsmallgunid = smallgun.getsmallgun(6)
                    If getsmallgunid <> "0" Then
                        ddshipgroup.ship(b).carry(ddshipgroup.ship(b).length - 1).equipmentid = getsmallgunid
                    End If
                End If
            Next
            ddshipgroup.sort()
            For b = 0 To ddshipgroup.ship(a).length - 1
                If ddshipgroup.ship(a).carry(b).equipmentid <> 0 Then
                    smallgun.addcollection(ddshipgroup.ship(a).carry(b).equipmentid, 1)
                End If
            Next
            For b = a + 1 To ddshipgroup.length - 1
                ddshipgroup.ship(b).removeequipment()
            Next
        Next
        Call list1refresh()
        Button5.Enabled = True
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ComboBox5.Items.Clear()
        ComboBox6.Items.Clear()
        ComboBox7.Items.Clear()
        TextBox2.Text = ""
        Dim index() As String = sgUIcontrol.list1clickid(ListBox1.SelectedIndex)
        If index(0) <> "-1" Then
            If index(1) = "-1" Then
                ComboBox5.Items.Add("自动判断")
                ComboBox5.Items.Add("二连")
                ComboBox5.Items.Add("CI")
                ComboBox5.SelectedIndex = ddshipgroup.ship(index(0)).nfmode

                ComboBox6.Items.Add("无限制")
                If userddshiptype.unit(ddshipgroup.ship(index(0)).id).gridcount = 3 Then
                    ComboBox6.Items.Add("锁定一格")
                ElseIf userddshiptype.unit(ddshipgroup.ship(index(0)).id).gridcount = 4 Then
                    ComboBox6.Items.Add("锁定一格")
                    ComboBox6.Items.Add("锁定两格")
                End If
                ComboBox6.SelectedIndex = ddshipgroup.ship(index(0)).restrictmode

                If userddshiptype.unit(ddshipgroup.ship(index(0)).id).enhanced = 0 Then
                    ComboBox7.Items.Add("未打孔")
                ElseIf userddshiptype.unit(ddshipgroup.ship(index(0)).id).enhanced = 1 Then
                    ComboBox7.Items.Add("锁定打孔")
                    ComboBox7.Items.Add("启用打孔")
                End If
                ComboBox7.SelectedIndex = ddshipgroup.ship(index(0)).enhancedlock

                TextBox2.Text = ddshipgroup.ship(index(0)).buffmultiple
            End If
        End If
    End Sub


    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If ListBox1.Items.Count > 0 Then
            If ListBox1.SelectedIndex >= 0 Then
                Dim selectedindex As Integer = ListBox1.SelectedIndex
                sgUIcontrol.changeshipnfset(sgUIcontrol.list1clickid(ListBox1.SelectedIndex)(0), ComboBox5.SelectedIndex, ComboBox6.SelectedIndex, ComboBox7.SelectedIndex, Val(TextBox2.Text))
                Call list1refresh()
                ListBox1.SelectedIndex = selectedindex
            End If
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If CheckedListBox1.Items.Count <> 0 Then
            For a = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(a, True)
            Next
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        If CheckedListBox1.Items.Count <> 0 Then
            For a = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(a, Not CheckedListBox1.GetItemChecked(a))
            Next
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If CheckedListBox1.Items.Count <> 0 Then
            For a = 0 To CheckedListBox1.Items.Count - 1
                CheckedListBox1.SetItemChecked(a, False)
            Next
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Dim index() As String = sgUIcontrol.list1clickid(ListBox1.SelectedIndex)
        If index(0) <> "-1" Then
            ddshipgroup.removeship(index(0))
            Call list1refresh()
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        ddshipgroup.removeallship()
        Call list1refresh()
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        sgUIcontrol.changesortmode(ComboBox8.SelectedIndex)
        Call checklist1refresh()
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        smallgun.loadsmallgunnuit(ComboBox9.SelectedIndex)
    End Sub
End Class