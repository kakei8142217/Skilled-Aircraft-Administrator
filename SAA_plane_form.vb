Public Class SAA_plane_form

    Private Sub SAA_plane_form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        filecontrol = New filecontrol_class
        cvshiptype = New cvshiptype_class
        restrictcontrol = New restrictcontrol_class
        userddshiptype = New userddshiptype_class
        smallgun = New smallgun_class
        sgUIcontrol = New sgUIcontrol_class
        ddshipgroup = New ddshipgroup_class


        saa_plane = Me
        plUIcontrol.changeenhancedmode()


        Dim cvshiptypecount As Integer = 1
        Do While cvshiptype.getid(cvshiptypecount) <> 0
            ComboBox1.Items.Add(cvshiptype.getattribute(cvshiptype.getid(cvshiptypecount), 1))
            cvshiptypecount = cvshiptypecount + 1
        Loop
        ComboBox1.Items.Add（"<点击加载未改造舰娘>"）

        For a = 0 To 19
            cvship(a) = New cvship_class(a)
        Next

        For a = 0 To 1
            ComboBox3.Items.Add("使用列表" & a + 1)
        Next
        ComboBox3.SelectedIndex = 0

        Call list1refresh()
        Call list2refresh()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        setting_mode = 1
        Dim saa_setting As New SAA_setting_form
        saa_setting.Show()
        saa_setting.Top = Me.Top
        saa_setting.Left = Me.Left
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If ComboBox1.SelectedIndex >= 0 Then
            Dim code As Integer = 64
            For a = 0 To 9
                If cvship(a).active = False Then
                    For b = 0 To 19
                        If cvship(b).active = True Then
                            If Val(Mid(Trim(cvship(b).uniquecode), 1, 3)) = cvshiptype.getid(ComboBox1.SelectedIndex + 1) Then
                                If Asc(Mid(Trim(cvship(b).uniquecode), Len(Trim(cvship(b).uniquecode)), 1)) > code Then
                                    code = Asc(Mid(Trim(cvship(b).uniquecode), Len(Trim(cvship(b).uniquecode)), 1))
                                End If
                            End If
                        End If
                    Next
                    cvship(a).uniquecode = Format(cvshiptype.getid(ComboBox1.SelectedIndex + 1), "000") & Chr(code + 1)
                    For b = 0 To 4
                        If cvshiptype.getattribute(cvship(a).uniquecode, b + 8) > 0 Then
                            vessel.clear()
                            vessel.amount = cvshiptype.getattribute(cvship(a).uniquecode, b + 8)
                            vessel.shipid = a
                            vessel.carryid = b
                            vessel.uniquecode = cvship(a).uniquecode & b
                            cvship(a).setcarry(vessel)
                        End If
                    Next
                    cvship(a).restrict = cvshiptype.getattribute(cvship(a).uniquecode, 22) * 10 + 0

                    plUIcontrol.setship(a)
                    Call list1refresh()
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If ComboBox1.SelectedIndex >= 0 Then
            Dim code As Integer = 64
            For a = 10 To 19
                If cvship(a).active = False Then
                    For b = 0 To 19
                        If cvship(b).active = True Then
                            If Val(Mid(Trim(cvship(b).uniquecode), 1, 3)) = cvshiptype.getid(ComboBox1.SelectedIndex + 1) Then
                                If Asc(Mid(Trim(cvship(b).uniquecode), Len(Trim(cvship(b).uniquecode)), 1)) > code Then
                                    code = Asc(Mid(Trim(cvship(b).uniquecode), Len(Trim(cvship(b).uniquecode)), 1))
                                End If
                            End If
                        End If
                    Next
                    cvship(a).uniquecode = Format(cvshiptype.getid(ComboBox1.SelectedIndex + 1), "000") & Chr(code + 1)
                    For b = 0 To 4
                        If cvshiptype.getattribute(cvship(a).uniquecode, b + 8) > 0 Then
                            vessel.clear()
                            vessel.amount = cvshiptype.getattribute(cvship(a).uniquecode, b + 8)
                            vessel.shipid = a
                            vessel.carryid = b
                            vessel.uniquecode = cvship(a).uniquecode & b
                            cvship(a).setcarry(vessel)
                        End If
                    Next
                    cvship(a).restrict = cvshiptype.getattribute(cvship(a).uniquecode, 22) * 10 + 0

                    plUIcontrol.setship(a)
                    Call list1refresh()
                    Exit For
                End If
            Next
        End If
    End Sub

    Private Sub list1refresh()
        ListBox1.Items.Clear()
        Dim num As Integer = 1
        Do While plUIcontrol.showstringlist1(num) <> ""
            ListBox1.Items.Add(plUIcontrol.showstringlist1(num))
            num = num + 1
        Loop
    End Sub

    Private Sub list2refresh()
        ListBox2.Items.Clear()
        Dim num As Integer = 1
        Do While plUIcontrol.showstringlist2(num, ListBox1.SelectedIndex) <> ""
            ListBox2.Items.Add(plUIcontrol.showstringlist2(num, ListBox1.SelectedIndex))
            num = num + 1
        Loop
    End Sub

    Private Sub ListBox1_Click(sender As Object, e As EventArgs) Handles ListBox1.Click
        ComboBox2.Items.Clear()
        Dim num As Integer = 1
        Do While plUIcontrol.showstringcombo2(num, ListBox1.SelectedIndex) <> ""
            ComboBox2.Items.Add(plUIcontrol.showstringcombo2(num, ListBox1.SelectedIndex))
            num = num + 1
        Loop
        Call list2refresh()
        If plUIcontrol.getshipid(ListBox1.SelectedIndex) <> -1 Then
            ComboBox2.SelectedIndex = cvship(plUIcontrol.getshipid(ListBox1.SelectedIndex)).restrict - Int(cvship(plUIcontrol.getshipid(ListBox1.SelectedIndex)).restrict / 10) * 10
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If ListBox1.SelectedIndex >= 0 Then
            Dim selectedindexvalue As Integer = ListBox1.SelectedIndex
            If ComboBox2.SelectedIndex >= 0 Then
                plUIcontrol.setshipattribute(ListBox1.SelectedIndex, ComboBox2.SelectedIndex)
                Call list1refresh()
                ListBox1.SelectedIndex = selectedindexvalue
                Call list2refresh()
            End If
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        plUIcontrol.removeship(ListBox1.SelectedIndex)
        Call list1refresh()
        Call list2refresh()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        plUIcontrol.clearship()
        Call list1refresh()
        Call list2refresh()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Button6.Enabled = False

        Dim ffexpectaavalue As Integer
        Dim cfexpectaavalue As Integer

        Dim onlysfnf As Boolean = False

        Dim complete As Boolean = False




        ordinaryfleetaavalue = Val(TextBox1.Text)
        combinedfleetaavalue = Val(TextBox2.Text)
        If combinedfleetaavalue < ordinaryfleetaavalue Then
            combinedfleetaavalue = ordinaryfleetaavalue
        End If


        If CheckBox6.Checked = False Then
            For a = 10 To 19
                If cvship(a).active = True Then
                    restrictcontrol.start = 10
                    onlysfnf = True
                End If
            Next
        End If
        If onlysfnf = False Then
            restrictcontrol.start = 0
        End If

        ffexpectaavalue = ordinaryfleetaavalue
        cfexpectaavalue = combinedfleetaavalue

        Do While complete = False
            For a = 0 To 19
                If cvship(a).active = True Then
                    cvship(a).resetcarryobtained()
                End If
            Next
            plane.clearcollection(1)
            restrictcontrol.resetequipcarry()
            plUIcontrol.clearerror()
            ffcaavalue = 0
            sfcaavalue = 0
            ffwaavalue = 0
            sfwaavalue = 0

            If CheckBox10.Checked = False Then
                Call matchingwplane(ffexpectaavalue, cfexpectaavalue)
            End If

            Call matchingcplane(ffexpectaavalue - ffwaavalue, cfexpectaavalue - ffwaavalue - sfwaavalue)

            If CheckBox10.Checked = True Then
                Call matchingwplane(ffexpectaavalue - ffcaavalue, cfexpectaavalue - ffcaavalue - sfcaavalue)
            End If


            If ffcaavalue + ffwaavalue = 0 Then
                complete = True
                plUIcontrol.adderror(100)
            End If
            If ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue >= combinedfleetaavalue Then
                complete = True
            End If
            If ffcaavalue + ffwaavalue < ordinaryfleetaavalue Then
                complete = True
                plUIcontrol.adderror(101)
            ElseIf ffcaavalue + ffwaavalue < ffexpectaavalue Then
                complete = True
                plUIcontrol.adderror(102)
            End If

            If ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue < combinedfleetaavalue Then
                ffexpectaavalue = ffexpectaavalue + combinedfleetaavalue - ffcaavalue - ffwaavalue - sfcaavalue - sfwaavalue
            End If
        Loop

        Call wfightermajorization()

        For a = 0 To 19
            If cvship(a).active = True Then
                If cvship(a).shellingfire = 0 And restrictcontrol.getrestrictattribute(a, 4) <= 1 And cvship(a).carrycount >= 3 Then
                    plUIcontrol.adderror(101 * 100 + a)
                End If
                If restrictcontrol.getrestrictattribute(a, 4) = 1 And restrictcontrol.CIstate(cvship(a).getcarry(0)) <> 1.25 Then
                    plUIcontrol.adderror(102 * 100 + a)
                End If
                If restrictcontrol.getrestrictattribute(a, 4) = 5 And restrictcontrol.CIstate(cvship(a).getcarry(0)) <= 1 Then
                    plUIcontrol.adderror(102 * 100 + a)
                End If
                If a >= (Val(onlysfnf) * 10) * -1 Then
                    If restrictcontrol.getrestrictattribute(a, 10) - restrictcontrol.getrestrictattribute(a, 7) < 0 Then
                        plUIcontrol.adderror(103 * 100 + a)
                    End If
                    If restrictcontrol.getrestrictattribute(a, 10) = 1 Then
                        If restrictcontrol.getrestrictattribute(a, 11) + restrictcontrol.getrestrictattribute(a, 12) = 0 Then
                            plUIcontrol.adderror(103 * 100 + a)
                        End If
                    End If
                End If
                If restrictcontrol.getrestrictattribute(a, 9) = 1 And restrictcontrol.getrestrictattribute(a, 18) < 1 Then
                    plUIcontrol.adderror(111 * 100 + a)
                End If
            End If
        Next

        Call list1refresh()
        Call list2refresh()

        Button6.Enabled = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        leading_intarget = 1

        Dim saa_map As New SAA_map_form
        saa_map.Show()
        saa_map.Top = Me.Top
        saa_map.Left = Me.Left
        Me.Hide()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        plUIcontrol.exportconfigurecode()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        plane.loadplaneunit(ComboBox3.SelectedIndex)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        saa_ddrank = New SAA_ddrank_form
        saa_ddrank.Show()
        saa_ddrank.Show()
        saa_ddrank.Top = Me.Top
        saa_ddrank.Left = Me.Left
        Me.Hide()
    End Sub

    Private Sub matchingcplane(ByVal ffexpectaavalue As Integer, ByVal cfexpectaavalue As Integer)
        Dim expectaavalue As Integer

        Dim cacheclsogroup As New carrygroup_class
        Dim cachecfightergroup As New carrygroup_class
        Dim cachecboomergroup As New carrygroup_class
        Dim cachectorpedogroup As New carrygroup_class
        Dim cachecplanegroup As New carrygroup_class

        Dim getplaneid As Integer
        Dim existed As New Collection

        carrygroup.clear()
        auxiliarygruop.clear()

        clsogroup.clear()
        cfightergroup.clear()
        cboomergroup.clear()
        ctorpedogroup.clear()
        cplanegroup.clear()



        For a = 0 To 19     '第一次装填容器
            If cvship(a).active = True Then
                If restrictcontrol.getrestrictattribute(a, 7) > 0 Then     '装载夜战要员容器
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 7)
                        clsogroup.setcarry(cvship(a).getcarry(-1))
                    Next
                End If

                If restrictcontrol.getrestrictattribute(a, 6) > 0 Then     '装载舰攻容器
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 6)
                        ctorpedogroup.setcarry(cvship(a).getcarry(10))
                    Next
                End If

                If restrictcontrol.getrestrictattribute(a, 5) > 0 Then     '装载舰爆容器
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 5)
                        cboomergroup.setcarry(cvship(a).getcarry(10))
                    Next
                End If

                If restrictcontrol.getrestrictattribute(a, 4) > 0 Then     '装载舰战容器                
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 4)
                        cfightergroup.setcarry(cvship(a).getcarry(10), 1)
                    Next
                End If

                If restrictcontrol.getrestrictattribute(a, 3) > 0 Then     '装载舰载机容器
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 3)
                        cplanegroup.setcarry(cvship(a).getcarry(10))
                    Next
                End If
            End If
        Next

        cfightergroup.sort()
        cboomergroup.sort()
        ctorpedogroup.sort()
        cplanegroup.sort()



        If clsogroup.carrycount > 0 Then  '装载夜战要员
            getplaneid = plane.getplane(8, 0, 0)
            For a = 0 To clsogroup.carrycount - 1
                vessel = clsogroup.getcarry(a)
                vessel.planeid = getplaneid
                If getplaneid <> 0 Then
                    If restrictcontrol.estimate(vessel) Then
                        clsogroup.setcarry(vessel)
                        plane.addcollection(getplaneid, 1)
                        getplaneid = plane.getplane(8, 0, 0)
                    End If
                End If
            Next
            '回收多余搭载格
            For a = 0 To clsogroup.carrycount - 1
                If clsogroup.getcarry(a).planeid = 0 Then
                    carrygroup.setcarry(clsogroup.getcarry(a))
                End If
            Next
            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    clsogroup.removecarry(carrygroup.getcarry(a))
                    cplanegroup.setcarry(carrygroup.getcarry(a))
                Next
            End If
            carrygroup.clear()
        End If

        If ctorpedogroup.carrycount <> 0 Then   '第一次装载类型1夜攻
            getplaneid = plane.getplane(0, 3, 0)
            For a = 0 To ctorpedogroup.carrycount - 1
                vessel = ctorpedogroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            ctorpedogroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 3, 0)
                        End If
                    End If
                End If
            Next
        End If

        If cfightergroup.carrycount <> 0 Then    '第一次装载夜战
            getplaneid = plane.getplane(0, 0, 0)
            For a = 0 To cfightergroup.carrycount - 1
                vessel = cfightergroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cfightergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 0, 0)
                        End If
                    End If
                End If
            Next
        End If

        If restrictcontrol.allnfpossible = False Then   '第一次补充装载夜战
            If cplanegroup.carrycount <> 0 Then
                getplaneid = plane.getplane(0, 0, 0)
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        vessel.planeid = getplaneid
                        If vessel.planeid <> 0 Then
                            If restrictcontrol.estimate(vessel) Then
                                cplanegroup.setcarry(vessel)
                                plane.addcollection(getplaneid, 1)
                                getplaneid = plane.getplane(0, 0, 0)
                            End If
                        End If
                    End If
                Next
            End If
        End If

        If ctorpedogroup.carrycount <> 0 Then   '第二次装载类型1夜攻
            Do
                getplaneid = plane.getplane(0, 3, 0)
                plane.addcollection(getplaneid)
            Loop Until plane.getattribute(getplaneid, 16) = 1 Or getplaneid = 0
            For a = 0 To ctorpedogroup.carrycount - 1
                vessel = ctorpedogroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            ctorpedogroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            Do
                                getplaneid = plane.getplane(0, 3, 0)
                                plane.addcollection(getplaneid)
                            Loop Until plane.getattribute(getplaneid, 16) = 1 Or getplaneid = 0
                        End If
                    End If
                End If
            Next
            plane.clearcollection()
        End If

        If cplanegroup.carrycount <> 0 Then   '第二次补充装载类型1夜攻
            Do
                getplaneid = plane.getplane(0, 3, 0)
                plane.addcollection(getplaneid)
            Loop Until plane.getattribute(getplaneid, 16) = 1 Or getplaneid = 0
            For a = 0 To cplanegroup.carrycount - 1
                vessel = cplanegroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            Do
                                getplaneid = plane.getplane(0, 3, 0)
                                plane.addcollection(getplaneid)
                            Loop Until plane.getattribute(getplaneid, 16) = 1 Or getplaneid = 0
                        End If
                    End If
                End If
            Next
            plane.clearcollection()
        End If

        If cfightergroup.carrycount <> 0 Then    '第二次装载夜战
            getplaneid = plane.getplane(0, 0, 0)
            For a = 0 To cfightergroup.carrycount - 1
                vessel = cfightergroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cfightergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 0, 0)
                        End If
                    End If
                End If
            Next
        End If

        If cplanegroup.carrycount <> 0 Then   '第二次补充装载夜战
            getplaneid = plane.getplane(0, 0, 0)
            For a = 0 To cplanegroup.carrycount - 1
                vessel = cplanegroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 0, 0)
                        End If
                    End If
                End If
            Next
        End If

        If cboomergroup.carrycount <> 0 Then      '第一次装载夜爆战
            getplaneid = plane.getplane(0, 1, 0)
            For a = 0 To cboomergroup.carrycount - 1
                vessel = cboomergroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cboomergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 1, 0)
                        End If
                    End If
                End If
            Next
        End If

        '补充夜爆战在装载类型2的夜攻之后

        If ctorpedogroup.carrycount <> 0 Then      '第一次装载类型2夜攻
            getplaneid = plane.getplane(0, 3, 0)
            For a = 0 To ctorpedogroup.carrycount - 1
                vessel = ctorpedogroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            ctorpedogroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 3, 0)
                        End If
                    End If
                End If
            Next
        End If

        If cplanegroup.carrycount <> 0 Then      '第一次补充装载类型2夜攻
            getplaneid = plane.getplane(0, 3, 0)
            For a = 0 To cplanegroup.carrycount - 1
                vessel = cplanegroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 3, 0)
                        End If
                    End If
                End If
            Next
        End If

        If cplanegroup.carrycount <> 0 Then      '第一次补充装载夜爆战
            getplaneid = plane.getplane(0, 1, 0)
            For a = 0 To cplanegroup.carrycount - 1
                vessel = cplanegroup.getcarry(a)
                If vessel.planeid = 0 Then
                    vessel.planeid = getplaneid
                    If vessel.planeid <> 0 Then
                        If restrictcontrol.estimate(vessel) Then
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            getplaneid = plane.getplane(0, 1, 0)
                        End If
                    End If
                End If
            Next
        End If

        If clsogroup.carrycount <> 0 Then     '回收LSO中无法达成夜战条件的搭载格
            For a = 0 To clsogroup.carrycount - 1
                vessel = clsogroup.getcarry(a)
                If restrictcontrol.nfstate(vessel) = 0 Then
                    vessel.planeid = 0
                    carrygroup.setcarry(vessel, 11)
                End If
            Next
            If carrygroup.carrycount > 0 Then
                For a = 0 To carrygroup.carrycount
                    clsogroup.removecarry(carrygroup.getcarry(a))
                    cplanegroup.setcarry(carrygroup.getcarry(a))
                Next
            End If
            carrygroup.clear()
        End If

        If cfightergroup.carrycount <> 0 Then     '回收舰战容器中未搭载夜战机的搭载格
            For a = 0 To cfightergroup.carrycount - 1
                vessel = cfightergroup.getcarry(a)
                If vessel.planeid = 0 Then
                    carrygroup.setcarry(vessel, 11)
                End If
            Next
            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    cfightergroup.removecarry(carrygroup.getcarry(a), 1)
                    cplanegroup.setcarry(carrygroup.getcarry(a))
                Next
            End If
            carrygroup.clear()
        End If

        cplanegroup.sort()

        Dim setup As Boolean = False

        If CheckBox1.Checked = True Then           '装载彩云
            If cplanegroup.carrycount <> 0 Then
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(cplanegroup.carrycount - 1 - a)
                    If vessel.planeid = 0 And vessel.shipid < 10 Then
                        getplaneid = plane.getplane(6)
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            setup = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If setup = False Then
                plUIcontrol.adderror(6)
            End If
        End If

        If CheckBox8.Checked = True Then       '装载夜战要员
            Dim combined As Boolean = False
            setup = False
            For b = 10 To 19
                If cvship(b).active = True Then
                    combined = True
                End If
            Next
            If cplanegroup.carrycount <> 0 Then
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(cplanegroup.carrycount - 1 - a)
                    If vessel.planeid = 0 And vessel.shipid < 10 Then
                        If combined Then
                            getplaneid = plane.getplane(21)
                        Else
                            getplaneid = plane.getplane(21, 0, 1)
                        End If
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                            setup = True
                            Exit For
                        End If
                    End If
                Next
            End If
            If setup = False Then
                plUIcontrol.adderror(21)
            End If
        End If



        If cplanegroup.carrycount <> 0 Then     '第二次装填舰战容器
            For a = 0 To cplanegroup.carrycount - 1
                vessel = cplanegroup.getcarry(cplanegroup.carrycount - 1 - a)
                If vessel.planeid = 0 Then
                    If restrictcontrol.getrestrictattribute(vessel.shipid, 14) - restrictcontrol.getrestrictattribute(vessel.shipid, 4) < 0 Then
                        carrygroup.setcarry(vessel, 1)
                    End If
                    If restrictcontrol.getrestrictattribute(vessel.shipid, 4) + restrictcontrol.getrestrictattribute(vessel.shipid, 7) = cvship(vessel.shipid).carrycount Then
                        carrygroup.setcarry(vessel, 11)
                    End If
                End If
            Next
            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    cplanegroup.removecarry(carrygroup.getcarry(a))
                    cfightergroup.setcarry(carrygroup.getcarry(a))
                Next
            End If
            carrygroup.clear()
        End If



        Dim exchange As Boolean = True

        If CheckBox5.Checked = True Then             '若勾选CV夜战搭载优化，将类型1的夜战换入较大的搭载格中
            If cfightergroup.carrycount <> 0 Then      '舰战容器中类型1夜战与类型2夜攻的搭载格的优化
                Do While exchange = True
                    exchange = False
                    For a = 0 To cfightergroup.carrycount - 1
                        vessel = cfightergroup.getcarry(a)
                        If vessel.planeid <> 0 Then
                            If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                                If ctorpedogroup.carrycount <> 0 Then
                                    For b = 0 To ctorpedogroup.carrycount - 1
                                        If ctorpedogroup.getcarry(b).planeid <> 0 Then
                                            If plane.getattribute(ctorpedogroup.getcarry(b).planeid, 16) = 2 Then
                                                If ctorpedogroup.getcarry(b).shipid = vessel.shipid And ctorpedogroup.getcarry(b).amount > vessel.amount Then
                                                    getplaneid = vessel.planeid
                                                    vessel.planeid = ctorpedogroup.getcarry(b).planeid
                                                    cfightergroup.setcarry(vessel)
                                                    vessel = ctorpedogroup.getcarry(b)
                                                    vessel.planeid = getplaneid
                                                    ctorpedogroup.setcarry(vessel)

                                                    vessel = cfightergroup.getcarry(a)
                                                    carrygroup.setcarry(vessel, 11)
                                                    vessel = ctorpedogroup.getcarry(b)
                                                    carrygroup.setcarry(vessel, 11)

                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    Next

                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            vessel = carrygroup.getcarry(a)
                            If plane.getattribute(vessel.planeid, 3) = 1 Then
                                ctorpedogroup.removecarry(vessel)
                                cfightergroup.setcarry(vessel)
                            ElseIf plane.getattribute(vessel.planeid, 3) = 4 Then
                                cfightergroup.removecarry(vessel)
                                ctorpedogroup.setcarry(vessel)
                            End If
                        Next
                        carrygroup.clear()
                        cfightergroup.sort()
                        ctorpedogroup.sort()
                        exchange = True
                    End If
                Loop
            End If

            If cplanegroup.carrycount <> 0 Then      '任意舰载机容器中类型1夜战与类型2夜攻的搭载格的优化
                exchange = True
                Do While exchange = True
                    exchange = False
                    For a = 0 To cplanegroup.carrycount - 1
                        vessel = cplanegroup.getcarry(a)
                        If vessel.planeid <> 0 Then
                            If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                                If ctorpedogroup.carrycount <> 0 Then
                                    For b = 0 To ctorpedogroup.carrycount - 1
                                        If ctorpedogroup.getcarry(b).planeid <> 0 Then
                                            If plane.getattribute(ctorpedogroup.getcarry(b).planeid, 16) = 2 Then
                                                If ctorpedogroup.getcarry(b).shipid = vessel.shipid And ctorpedogroup.getcarry(b).amount > vessel.amount Then
                                                    getplaneid = vessel.planeid
                                                    vessel.planeid = ctorpedogroup.getcarry(b).planeid
                                                    cplanegroup.setcarry(vessel)
                                                    vessel = ctorpedogroup.getcarry(b)
                                                    vessel.planeid = getplaneid
                                                    ctorpedogroup.setcarry(vessel)

                                                    vessel = cplanegroup.getcarry(a)
                                                    carrygroup.setcarry(vessel, 11)
                                                    vessel = ctorpedogroup.getcarry(b)
                                                    carrygroup.setcarry(vessel, 11)

                                                    Exit For
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        End If
                    Next

                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            vessel = carrygroup.getcarry(a)
                            If plane.getattribute(vessel.planeid, 3) = 1 Then
                                ctorpedogroup.removecarry(vessel)
                                cplanegroup.setcarry(vessel)
                            ElseIf plane.getattribute(vessel.planeid, 3) = 4 Then
                                cplanegroup.removecarry(vessel)
                                ctorpedogroup.setcarry(vessel)
                            End If
                        Next
                        carrygroup.clear()
                        cplanegroup.sort()
                        ctorpedogroup.sort()
                        exchange = True
                    End If
                Loop
            End If


            If cfightergroup.carrycount <> 0 Then      '舰战容器中类型1夜战与类型2夜爆战的搭载格的优化
                exchange = True
                Do While exchange = True
                    exchange = False
                    For a = 0 To cfightergroup.carrycount - 1
                        vessel = cfightergroup.getcarry(a)
                        If vessel.planeid <> 0 Then
                            If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                                If cboomergroup.carrycount <> 0 Then
                                    For b = 0 To cboomergroup.carrycount - 1
                                        'If cboomergroup.getcarry(b).planeid <> 0 Then
                                        '    If plane.getattribute(cboomergroup.getcarry(b).planeid, 16) = 2 Then
                                        If cboomergroup.getcarry(b).shipid = vessel.shipid And cboomergroup.getcarry(b).amount > vessel.amount Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cboomergroup.getcarry(b).planeid
                                            cfightergroup.setcarry(vessel)
                                            vessel = cboomergroup.getcarry(b)
                                            vessel.planeid = getplaneid
                                            cboomergroup.setcarry(vessel)

                                            vessel = cfightergroup.getcarry(a)
                                            carrygroup.setcarry(vessel, 11)
                                            vessel = cboomergroup.getcarry(b)
                                            carrygroup.setcarry(vessel, 11)

                                            Exit For
                                        End If
                                        '    End If
                                        'End If
                                    Next
                                End If
                            End If
                        End If
                    Next

                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            vessel = carrygroup.getcarry(a)
                            If plane.getattribute(vessel.planeid, 3) = 1 Then
                                cboomergroup.removecarry(vessel)
                                cfightergroup.setcarry(vessel)
                            ElseIf plane.getattribute(vessel.planeid, 3) = 2 Or vessel.planeid = 0 Then
                                cfightergroup.removecarry(vessel)
                                cboomergroup.setcarry(vessel)
                            End If
                        Next
                        carrygroup.clear()
                        cfightergroup.sort()
                        cboomergroup.sort()
                        exchange = True
                    End If
                Loop
            End If



            If cplanegroup.carrycount <> 0 Then      '任意舰载机容器中类型1夜战与类型2夜爆战的搭载格的优化
                exchange = True
                Do While exchange = True
                    exchange = False
                    For a = 0 To cplanegroup.carrycount - 1
                        vessel = cplanegroup.getcarry(a)
                        If vessel.planeid <> 0 Then
                            If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                                If cboomergroup.carrycount <> 0 Then
                                    For b = 0 To cboomergroup.carrycount - 1
                                        'If cboomergroup.getcarry(b).planeid <> 0 Then
                                        '    If plane.getattribute(cboomergroup.getcarry(b).planeid, 16) = 2 Then
                                        If cboomergroup.getcarry(b).shipid = vessel.shipid And cboomergroup.getcarry(b).amount > vessel.amount Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cboomergroup.getcarry(b).planeid
                                            cplanegroup.setcarry(vessel)
                                            vessel = cboomergroup.getcarry(b)
                                            vessel.planeid = getplaneid
                                            cboomergroup.setcarry(vessel)

                                            vessel = cplanegroup.getcarry(a)
                                            carrygroup.setcarry(vessel, 11)
                                            vessel = cboomergroup.getcarry(b)
                                            carrygroup.setcarry(vessel, 11)

                                            Exit For
                                        End If
                                        '    End If
                                        'End If
                                    Next
                                End If
                            End If
                        End If
                    Next

                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            vessel = carrygroup.getcarry(a)
                            If plane.getattribute(vessel.planeid, 3) = 1 Then
                                cboomergroup.removecarry(vessel)
                                cplanegroup.setcarry(vessel)
                            ElseIf plane.getattribute(vessel.planeid, 3) = 2 Or vessel.planeid = 0 Then
                                cplanegroup.removecarry(vessel)
                                cboomergroup.setcarry(vessel)
                            End If
                        Next
                        carrygroup.clear()
                        cplanegroup.sort()
                        cboomergroup.sort()
                        exchange = True
                    End If
                Loop
            End If
        End If

        If cfightergroup.carrycount + cplanegroup.carrycount > 0 Then   '夜战对空值与搭载格大小的优化
            For a = 0 To cfightergroup.carrycount + cplanegroup.carrycount - 1
                If a <= cfightergroup.carrycount - 1 Then
                    vessel = cfightergroup.getcarry(a)
                    If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                        For b = 0 To cfightergroup.carrycount + cplanegroup.carrycount - 1
                            If b <= cfightergroup.carrycount - 1 Then
                                If plane.getattribute(cfightergroup.getcarry(b).planeid, 3) = 1 And plane.getattribute(cfightergroup.getcarry(b).planeid, 16) = 1 Then
                                    If cfightergroup.getcarry(b).amount > vessel.amount Then
                                        If plane.getattribute(cfightergroup.getcarry(b).planeid, 6) < plane.getattribute(vessel.planeid, 6) Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cfightergroup.getcarry(b).planeid
                                            cfightergroup.setcarry(vessel)
                                            vessel = cfightergroup.getcarry(b)
                                            vessel.planeid = getplaneid
                                            cfightergroup.setcarry(vessel)
                                        End If
                                    End If
                                End If
                            Else
                                If plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 3) = 1 And plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 16) = 1 Then
                                    If cplanegroup.getcarry(b - cfightergroup.carrycount).amount > vessel.amount Then
                                        If plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 6) < plane.getattribute(vessel.planeid, 6) Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cplanegroup.getcarry(b - cfightergroup.carrycount).planeid
                                            cfightergroup.setcarry(vessel)
                                            vessel = cplanegroup.getcarry(b - cfightergroup.carrycount)
                                            vessel.planeid = getplaneid
                                            cplanegroup.setcarry(vessel)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                Else
                    vessel = cplanegroup.getcarry(a - cfightergroup.carrycount)
                    If plane.getattribute(vessel.planeid, 3) = 1 And plane.getattribute(vessel.planeid, 16) = 1 Then
                        For b = 0 To cfightergroup.carrycount + cplanegroup.carrycount - 1
                            If b <= cfightergroup.carrycount - 1 Then
                                If plane.getattribute(cfightergroup.getcarry(b).planeid, 3) = 1 And plane.getattribute(cfightergroup.getcarry(b).planeid, 16) = 1 Then
                                    If cfightergroup.getcarry(b).amount > vessel.amount Then
                                        If plane.getattribute(cfightergroup.getcarry(b).planeid, 6) < plane.getattribute(vessel.planeid, 6) Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cfightergroup.getcarry(b).planeid
                                            cplanegroup.setcarry(vessel)
                                            vessel = cfightergroup.getcarry(b)
                                            vessel.planeid = getplaneid
                                            cfightergroup.setcarry(vessel)
                                        End If
                                    End If
                                End If
                            Else
                                If plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 3) = 1 And plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 16) = 1 Then
                                    If cplanegroup.getcarry(b - cfightergroup.carrycount).amount > vessel.amount Then
                                        If plane.getattribute(cplanegroup.getcarry(b - cfightergroup.carrycount).planeid, 6) < plane.getattribute(vessel.planeid, 6) Then
                                            getplaneid = vessel.planeid
                                            vessel.planeid = cplanegroup.getcarry(b - cfightergroup.carrycount).planeid
                                            cplanegroup.setcarry(vessel)
                                            vessel = cplanegroup.getcarry(b - cfightergroup.carrycount)
                                            vessel.planeid = getplaneid
                                            cplanegroup.setcarry(vessel)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
            Next
        End If

        If CheckBox7.Checked = True Then    '若勾选不装备类型2夜间机，则将所有类型2的夜战机清除
            If ctorpedogroup.carrycount <> 0 Then
                For a = 0 To ctorpedogroup.carrycount - 1
                    vessel = ctorpedogroup.getcarry(a)
                    If plane.getattribute(vessel.planeid, 16) = 2 Then
                        carrygroup.setcarry(vessel, 11)
                        vessel.planeid = 0
                        ctorpedogroup.setcarry(vessel)
                    End If
                Next
            End If
            If cboomergroup.carrycount <> 0 Then
                For a = 0 To cboomergroup.carrycount - 1
                    vessel = cboomergroup.getcarry(a)
                    If plane.getattribute(vessel.planeid, 16) = 2 Then
                        carrygroup.setcarry(vessel, 11)
                        vessel.planeid = 0
                        cboomergroup.setcarry(vessel)
                    End If
                Next
            End If
            If cplanegroup.carrycount <> 0 Then
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(a)
                    If plane.getattribute(vessel.planeid, 16) = 2 Then
                        carrygroup.setcarry(vessel, 11)
                        vessel.planeid = 0
                        cplanegroup.setcarry(vessel)
                    End If
                Next
            End If
            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    vessel = carrygroup.getcarry(a)
                    plane.removecollection(vessel.planeid)
                Next
                carrygroup.clear()
            End If
        End If


        If clsogroup.carrycount <> 0 Then                          '将联合舰队的搭载格存入缓存
            For a = 0 To clsogroup.carrycount - 1
                cacheclsogroup.setcarry(clsogroup.getcarry(a), 11)
            Next
        End If
        If cfightergroup.carrycount <> 0 Then
            For a = 0 To cfightergroup.carrycount - 1
                cachecfightergroup.setcarry(cfightergroup.getcarry(a), 11)
            Next
        End If
        If cboomergroup.carrycount <> 0 Then
            For a = 0 To cboomergroup.carrycount - 1
                cachecboomergroup.setcarry(cboomergroup.getcarry(a), 11)
            Next
        End If
        If ctorpedogroup.carrycount <> 0 Then
            For a = 0 To ctorpedogroup.carrycount - 1
                cachectorpedogroup.setcarry(ctorpedogroup.getcarry(a), 11)
            Next
        End If
        If cplanegroup.carrycount <> 0 Then
            For a = 0 To cplanegroup.carrycount - 1
                cachecplanegroup.setcarry(cplanegroup.getcarry(a), 11)
            Next
        End If




        For l = 0 To 1
            '由缓存中取出对应舰队的搭载格数据
            clsogroup.clear()
            cfightergroup.clear()
            cboomergroup.clear()
            ctorpedogroup.clear()
            cplanegroup.clear()
            If cacheclsogroup.carrycount <> 0 Then
                For a = 0 To cacheclsogroup.carrycount - 1
                    vessel = cacheclsogroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        clsogroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachecfightergroup.carrycount <> 0 Then
                For a = 0 To cachecfightergroup.carrycount - 1
                    vessel = cachecfightergroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        cfightergroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachecboomergroup.carrycount <> 0 Then
                For a = 0 To cachecboomergroup.carrycount - 1
                    vessel = cachecboomergroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        cboomergroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachectorpedogroup.carrycount <> 0 Then
                For a = 0 To cachectorpedogroup.carrycount - 1
                    vessel = cachectorpedogroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        ctorpedogroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachecplanegroup.carrycount <> 0 Then
                For a = 0 To cachecplanegroup.carrycount - 1
                    vessel = cachecplanegroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        cplanegroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            cfightergroup.sort()
            cboomergroup.sort()
            ctorpedogroup.sort()
            cplanegroup.sort()

            If l = 0 Then          '计算第一舰队或第二舰队需求的制空值
                expectaavalue = ffexpectaavalue
            ElseIf l = 1 Then
                expectaavalue = cfexpectaavalue - ffcaavalue
            End If


            If ctorpedogroup.carrycount <> 0 Then            '向舰攻容器装载舰攻
                For a = 0 To ctorpedogroup.carrycount - 1
                    vessel = ctorpedogroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        getplaneid = plane.getplane(4)
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            ctorpedogroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                        Else
                            plUIcontrol.adderror(4)
                        End If
                    End If
                Next
            End If

            If cboomergroup.carrycount <> 0 Then                '向舰爆容器装载舰爆
                For a = 0 To cboomergroup.carrycount - 1
                    vessel = cboomergroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        getplaneid = plane.getplane(3)
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            cboomergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                        Else
                            plUIcontrol.adderror(3)
                        End If
                    End If
                Next
            End If

            If cplanegroup.carrycount <> 0 Then                '向任何舰载机容器装载舰攻/舰爆
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        getplaneid = plane.getplane(vessel.enabletype({4, 3}))
                        If getplaneid = 0 Then
                            getplaneid = plane.getplane(3)
                            plUIcontrol.adderror(4)
                        End If
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            cplanegroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                        Else
                            plUIcontrol.adderror(3)
                        End If
                    End If
                Next
            End If



            Dim stageexpectaavalue As Integer                  '第一次轮装舰战
            stageexpectaavalue = expectaavalue - cfightergroup.AAvalue - cboomergroup.AAvalue - ctorpedogroup.AAvalue - cplanegroup.AAvalue

            Dim round As Integer = 0
            Dim getcarrycount As Integer = 0



            getplaneid = plane.getplane(1)
            If getplaneid = 0 Then
                plUIcontrol.adderror(1)
            End If
            Do While carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue And getplaneid <> 0 And round <= cplanegroup.carrycount
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.clearcollection()
                getcarrycount = 0
                If cfightergroup.carrycount <> 0 Then
                    For a = 0 To cfightergroup.carrycount - 1
                        vessel = cfightergroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            carrygroup.setcarry(vessel, 11)
                        End If
                    Next
                End If
                If cplanegroup.carrycount <> 0 Then
                    For a = 0 To cplanegroup.carrycount - 1
                        vessel = cplanegroup.getcarry(cplanegroup.carrycount - 1 - a)
                        If Not (plane.getattribute(vessel.planeid, 16) <> 0 And restrictcontrol.nfstate(vessel) > 1) Then
                            If plane.getattribute(vessel.planeid, 3) <> 6 Then
                                If getcarrycount < round Then
                                    carrygroup.setcarry(vessel, 11)
                                    getcarrycount = getcarrycount + 1
                                End If
                            End If
                        End If
                    Next
                End If
                carrygroup.sort()
                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        getplaneid = plane.getplane(1)
                        If getplaneid <> 0 Then
                            vessel = carrygroup.getcarry(a)
                            auxiliarygruop.setcarry(vessel, 11)
                            vessel.planeid = getplaneid
                            carrygroup.setcarry(vessel, 11)
                            plane.addcollection(getplaneid)
                        Else
                            plUIcontrol.adderror(1)
                        End If
                    Next
                End If
                round = round + 1
                getplaneid = plane.getplane(1)
            Loop

            Dim insert As Boolean
            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    insert = False
                    vessel = carrygroup.getcarry(a)
                    If cfightergroup.carrycount <> 0 Then
                        For b = 0 To cfightergroup.carrycount - 1
                            If insert = True Then
                                Exit For
                            End If
                            If vessel.uniquecode = cfightergroup.getcarry(b).uniquecode Then
                                cfightergroup.setcarry(vessel)
                                insert = True
                            End If
                        Next
                    End If
                    If cplanegroup.carrycount <> 0 Then
                        For b = 0 To cplanegroup.carrycount - 1
                            If insert = True Then
                                Exit For
                            End If
                            If vessel.uniquecode = cplanegroup.getcarry(b).uniquecode Then
                                cplanegroup.setcarry(vessel)
                                Exit For
                            End If
                        Next
                        vessel = auxiliarygruop.getcarry(0)
                        plane.removecollection(vessel.planeid)
                    End If
                Next
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.movecollection()
            End If


            If cfightergroup.carrycount <> 0 Then           '向舰战容器补充舰战
                For a = 0 To cfightergroup.carrycount - 1
                    vessel = cfightergroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        getplaneid = plane.getplane(1, 0, 1)
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            cfightergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                        Else
                            plUIcontrol.adderror(1)
                        End If
                    End If
                Next
            End If

            If CheckBox2.Checked = False Then    '当允许CI使用爆战时，若制空不足，尝试向舰爆容器装载爆战
                stageexpectaavalue = expectaavalue - cfightergroup.AAvalue - cboomergroup.AAvalue - ctorpedogroup.AAvalue - cplanegroup.AAvalue
                If stageexpectaavalue > 0 Then
                    round = 0
                    getcarrycount = 0
                    getplaneid = plane.getplane(2)
                    Do While carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue And getplaneid <> 0 And round <= cboomergroup.carrycount
                        carrygroup.clear()
                        auxiliarygruop.clear()
                        plane.clearcollection()
                        getcarrycount = 0
                        If cboomergroup.carrycount <> 0 Then
                            For a = 0 To cboomergroup.carrycount - 1
                                vessel = cboomergroup.getcarry(a)
                                If Not (plane.getattribute(vessel.planeid, 16) <> 0 And restrictcontrol.nfstate(vessel) > 1) Then
                                    If getcarrycount < round Then
                                        carrygroup.setcarry(vessel, 11)
                                        getcarrycount = getcarrycount + 1
                                    End If
                                End If
                            Next
                        End If
                        carrygroup.sort()
                        If carrygroup.carrycount <> 0 Then
                            For a = 0 To carrygroup.carrycount - 1
                                getplaneid = plane.getplane(2)
                                If getplaneid <> 0 Then
                                    vessel = carrygroup.getcarry(a)
                                    auxiliarygruop.setcarry(vessel, 11)
                                    vessel.planeid = getplaneid
                                    carrygroup.setcarry(vessel, 11)
                                    plane.addcollection(getplaneid)
                                End If
                            Next
                        End If
                        round = round + 1
                        getplaneid = plane.getplane(2)
                    Loop
                End If
                plane.movecollection()

                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        cboomergroup.setcarry(carrygroup.getcarry(a))
                    Next
                    For a = 0 To auxiliarygruop.carrycount - 1
                        plane.removecollection(auxiliarygruop.getcarry(a).planeid)
                    Next
                    For a = 0 To cboomergroup.carrycount - 1
                        vessel = cboomergroup.getcarry(a)
                        If plane.getattribute(vessel.planeid, 3) <> 2 Then
                            plane.removecollection(vessel.planeid)
                            vessel.planeid = 0
                            cboomergroup.setcarry(vessel)
                        End If
                    Next
                    For a = 0 To cboomergroup.carrycount - 1
                        vessel = cboomergroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            getplaneid = plane.getplane(3)
                            If getplaneid <> 0 Then
                                vessel.planeid = getplaneid
                                cboomergroup.setcarry(vessel)
                                plane.addcollection(getplaneid, 1)
                            End If
                        End If
                    Next
                    carrygroup.clear()
                    auxiliarygruop.clear()
                End If
            End If


            stageexpectaavalue = expectaavalue - cfightergroup.AAvalue - cboomergroup.AAvalue - ctorpedogroup.AAvalue - cplanegroup.AAvalue      '若制空不足，尝试向舰攻容器装载喷气机/爆战
            If stageexpectaavalue > 0 Then
                round = 0
                getcarrycount = 0
                getplaneid = plane.getplane(5)
                If getplaneid = 0 Then
                    getplaneid = plane.getplane(2)
                End If
                existed.Clear()
                Do While carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue And getplaneid <> 0 And round <= ctorpedogroup.carrycount
                    carrygroup.clear()
                    auxiliarygruop.clear()
                    plane.clearcollection()
                    getcarrycount = 0
                    If ctorpedogroup.carrycount <> 0 Then
                        If CheckBox4.Checked = False Then
                            For a = 0 To ctorpedogroup.carrycount - 1
                                vessel = ctorpedogroup.getcarry(a)
                                If restrictcontrol.CIstate(vessel) < 1 Then
                                    If cvshiptype.getattribute(cvship(vessel.shipid).uniquecode, 18) = 1 Then
                                        If Not (plane.getattribute(vessel.planeid, 16) <> 0 And restrictcontrol.nfstate(vessel) > 1) Then
                                            getplaneid = plane.getplane(5)
                                            If getplaneid <> 0 Then
                                                If getcarrycount < round Then
                                                    carrygroup.setcarry(vessel, 11)
                                                    existed.Add(vessel.shipid, vessel.shipid)
                                                    plane.addcollection(getplaneid)
                                                    getcarrycount = getcarrycount + 1
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                            plane.clearcollection()
                        End If
                        For a = 0 To ctorpedogroup.carrycount - 1
                            vessel = ctorpedogroup.getcarry(a)
                            If restrictcontrol.CIstate(vessel) < 1 Then
                                If Not (plane.getattribute(vessel.planeid, 16) <> 0 And restrictcontrol.nfstate(vessel) > 1) Then
                                    If existed.Contains(vessel.shipid) = False Then
                                        If getcarrycount < round Then
                                            carrygroup.setcarry(vessel, 11)
                                            getcarrycount = getcarrycount + 1
                                        End If
                                    End If
                                End If
                            End If
                        Next
                        If carrygroup.carrycount <> 0 Then
                            For a = 0 To carrygroup.carrycount - 1
                                vessel = carrygroup.getcarry(a)
                                If cvshiptype.getattribute(cvship(vessel.shipid).uniquecode, 18) = 1 Then
                                    If CheckBox4.Checked = False Then
                                        getplaneid = plane.getplane(5)
                                        If getplaneid <> 0 Then
                                            auxiliarygruop.setcarry(vessel, 11)
                                            vessel.planeid = getplaneid
                                            carrygroup.setcarry(vessel, 11)
                                            plane.addcollection(getplaneid)
                                        End If
                                    Else
                                        getplaneid = plane.getplane(2)
                                        If getplaneid <> 0 Then
                                            auxiliarygruop.setcarry(vessel, 11)
                                            vessel.planeid = getplaneid
                                            carrygroup.setcarry(vessel, 11)
                                            plane.addcollection(getplaneid)
                                        End If
                                    End If
                                Else
                                    getplaneid = plane.getplane(2)
                                    If getplaneid <> 0 Then
                                        auxiliarygruop.setcarry(vessel, 11)
                                        vessel.planeid = getplaneid
                                        carrygroup.setcarry(vessel, 11)
                                        plane.addcollection(getplaneid)
                                    End If
                                End If
                            Next
                        End If
                    End If
                    round = round + 1
                    plane.clearcollection()
                    existed.Clear()
                Loop
                plane.movecollection()
                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        ctorpedogroup.setcarry(carrygroup.getcarry(a))
                    Next
                    For a = 0 To auxiliarygruop.carrycount - 1
                        plane.removecollection(auxiliarygruop.getcarry(a).planeid)
                    Next
                    For a = 0 To ctorpedogroup.carrycount - 1
                        vessel = ctorpedogroup.getcarry(a)
                        If plane.getattribute(vessel.planeid, 3) = 4 Or plane.getattribute(vessel.planeid, 3) = 3 Then
                            If plane.getattribute(vessel.planeid, 16) = 0 Then
                                plane.removecollection(vessel.planeid)
                                vessel.planeid = 0
                                ctorpedogroup.setcarry(vessel)
                            End If
                        End If
                    Next
                    For a = 0 To ctorpedogroup.carrycount - 1
                        vessel = ctorpedogroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            getplaneid = plane.getplane(4)
                            If getplaneid = 0 Then
                                getplaneid = plane.getplane(3)
                            End If
                            If getplaneid <> 0 Then
                                vessel.planeid = getplaneid
                                ctorpedogroup.setcarry(vessel)
                                plane.addcollection(getplaneid, 1)
                            End If
                        End If
                    Next
                End If
                carrygroup.clear()
                auxiliarygruop.clear()
            End If

            If CheckBox3.Checked = True Then     '尝试将任意舰载机容器中的小格舰战更换为攻击机
                stageexpectaavalue = expectaavalue - cfightergroup.AAvalue - cboomergroup.AAvalue - ctorpedogroup.AAvalue - cplanegroup.AAvalue
                If stageexpectaavalue < 0 Then
                    If cfightergroup.carrycount <> 0 Then
                        For a = 0 To cfightergroup.carrycount - 1
                            vessel = cfightergroup.getcarry(a)
                            If Not (restrictcontrol.nfstate(vessel) >= 1 And plane.getattribute(vessel.planeid, 16) <> 0) Then
                                If restrictcontrol.getrestrictattribute(vessel.shipid, 4) = 1 Then
                                    If cplanegroup.carrycount <> 0 Then
                                        For b = 0 To cplanegroup.carrycount - 1
                                            If cplanegroup.getcarry(cplanegroup.carrycount - 1 - b).shipid = vessel.shipid And plane.getattribute(cplanegroup.getcarry(cplanegroup.carrycount - 1 - b).planeid, 3) = 1 Then
                                                If Not (restrictcontrol.nfstate(cplanegroup.getcarry(cplanegroup.carrycount - 1 - b)) >= 1 And plane.getattribute(cplanegroup.getcarry(cplanegroup.carrycount - 1 - b).planeid, 16) <> 0) Then
                                                    carrygroup.setcarry(vessel, 11)
                                                    auxiliarygruop.setcarry(cplanegroup.getcarry(cplanegroup.carrycount - 1 - b), 11)
                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If
                        Next
                        If carrygroup.carrycount <> 0 Then
                            For a = 0 To carrygroup.carrycount - 1
                                vessel = carrygroup.getcarry(a)
                                cfightergroup.removecarry(vessel)
                                cplanegroup.setcarry(vessel)
                                vessel = auxiliarygruop.getcarry(a)
                                cplanegroup.removecarry(vessel)
                                cfightergroup.setcarry(vessel)
                            Next
                        End If
                    End If
                    carrygroup.clear()
                    auxiliarygruop.clear()

                    If cplanegroup.carrycount <> 0 Then
                        For a = 0 To cplanegroup.carrycount - 1
                            vessel = cplanegroup.getcarry(a)
                            If plane.getattribute(vessel.planeid, 3) = 1 Then
                                If Not (restrictcontrol.nfstate(vessel) >= 1 And plane.getattribute(vessel.planeid, 16) <> 0) Then
                                    getplaneid = plane.getplane(vessel.enabletype({4, 3}))
                                    If getplaneid = 0 Then
                                        getplaneid = plane.getplane(3)
                                    End If
                                    If getplaneid <> 0 Then
                                        auxiliarygruop.setcarry(vessel, 11)
                                        vessel.planeid = getplaneid
                                        carrygroup.setcarry(vessel, 11)
                                        If carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue Then
                                            carrygroup.removecarry(vessel, 11)
                                            auxiliarygruop.removecarry(vessel, 11)
                                        Else
                                            'stageexpectaavalue = expectaavalue - cfightergroup.AAvalue - cboomergroup.AAvalue - ctorpedogroup.AAvalue - cplanegroup.AAvalue - carrygroup.AAvalue + auxiliarygruop.AAvalue
                                            plane.addcollection(getplaneid)
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            cplanegroup.setcarry(carrygroup.getcarry(a))
                            vessel = auxiliarygruop.getcarry(a)
                            plane.removecollection(vessel.planeid)
                        Next
                    End If
                End If
            End If


            If CheckBox9.Checked = True Then         '若勾选航空CI装配两爆战，尝试将CI舰娘的一个不必要的舰攻替换为舰爆
                If cplanegroup.carrycount <> 0 Then
                    For a = 0 To cplanegroup.carrycount - 1
                        vessel = cplanegroup.getcarry(cplanegroup.carrycount - 1 - a)
                        If plane.getattribute(vessel.planeid, 3) = 4 Then
                            If Not (restrictcontrol.nfstate(vessel) >= 1 And plane.getattribute(vessel.planeid, 16) <> 0) Then
                                If restrictcontrol.getrestrictattribute(vessel.shipid, 5) >= 1 And restrictcontrol.getrestrictattribute(vessel.shipid, 15) = 1 Then
                                    getplaneid = plane.getplane(3)
                                    If getplaneid <> 0 Then
                                        carrygroup.setcarry(vessel)
                                        vessel.planeid = getplaneid
                                        cplanegroup.setcarry(vessel)
                                    End If
                                End If
                            End If
                        End If
                    Next
                    If carrygroup.carrycount <> 0 Then
                        For a = 0 To carrygroup.carrycount - 1
                            plane.removecollection(carrygroup.getcarry(a).planeid)
                        Next
                    End If
                    carrygroup.clear()
                End If
            End If




            'ListBox2.Items.Add("clso:" & clsogroup.carrycount & "/" & clsogroup.AAvalue)
            'If clsogroup.carrycount <> 0 Then
            '    For a = 0 To clsogroup.carrycount - 1
            '        ListBox2.Items.Add(clsogroup.getcarry(a).planeid & plane.getattribute(clsogroup.getcarry(a).planeid, 1) & clsogroup.getcarry(a).uniquecode)
            '    Next
            'End If
            'If ctorpedogroup.carrycount <> 0 Then
            '    ListBox2.Items.Add("ctorpedo:" & ctorpedogroup.carrycount & "/" & ctorpedogroup.AAvalue)
            '    For a = 0 To ctorpedogroup.carrycount - 1
            '        ListBox2.Items.Add(ctorpedogroup.getcarry(a).planeid & plane.getattribute(ctorpedogroup.getcarry(a).planeid, 1) & ctorpedogroup.getcarry(a).uniquecode)
            '    Next
            'End If
            'If cboomergroup.carrycount <> 0 Then
            '    ListBox2.Items.Add("cboomer:" & cboomergroup.carrycount & "/" & cboomergroup.AAvalue)
            '    For a = 0 To cboomergroup.carrycount - 1
            '        ListBox2.Items.Add(cboomergroup.getcarry(a).planeid & plane.getattribute(cboomergroup.getcarry(a).planeid, 1) & cboomergroup.getcarry(a).uniquecode)
            '    Next
            'End If
            'If cfightergroup.carrycount <> 0 Then
            '    ListBox2.Items.Add("cfighter:" & cfightergroup.carrycount & "/" & cfightergroup.AAvalue)
            '    For a = 0 To cfightergroup.carrycount - 1
            '        ListBox2.Items.Add(cfightergroup.getcarry(a).planeid & plane.getattribute(cfightergroup.getcarry(a).planeid, 1) & cfightergroup.getcarry(a).uniquecode)
            '    Next
            'End If
            'If cplanegroup.carrycount <> 0 Then
            '    ListBox2.Items.Add("cplane:" & cplanegroup.carrycount & "/" & cplanegroup.AAvalue)
            '    For a = 0 To cplanegroup.carrycount - 1
            '        ListBox2.Items.Add(cplanegroup.getcarry(a).planeid & plane.getattribute(cplanegroup.getcarry(a).planeid, 1) & cplanegroup.getcarry(a).uniquecode)
            '    Next
            'End If


            If cfightergroup.carrycount <> 0 Then
                For a = 0 To cfightergroup.carrycount - 1
                    vessel = cfightergroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If

            If cboomergroup.carrycount <> 0 Then
                For a = 0 To cboomergroup.carrycount - 1
                    vessel = cboomergroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If

            If ctorpedogroup.carrycount <> 0 Then
                For a = 0 To ctorpedogroup.carrycount - 1
                    vessel = ctorpedogroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If

            If clsogroup.carrycount <> 0 Then
                For a = 0 To clsogroup.carrycount - 1
                    vessel = clsogroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If

            If cplanegroup.carrycount <> 0 Then
                For a = 0 To cplanegroup.carrycount - 1
                    vessel = cplanegroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If


            If l = 0 Then
                ffcaavalue = cfightergroup.AAvalue + cboomergroup.AAvalue + ctorpedogroup.AAvalue + cplanegroup.AAvalue
            ElseIf l = 1 Then
                sfcaavalue = cfightergroup.AAvalue + cboomergroup.AAvalue + ctorpedogroup.AAvalue + cplanegroup.AAvalue
            End If

        Next



    End Sub

    Private Sub matchingwplane(ByVal ffexpectaavalue As Integer, ByVal cfexpectaavalue As Integer)
        Dim expectaavalue As Integer

        Dim cachewplanegroup As New carrygroup_class
        Dim cachewboomergroup As New carrygroup_class
        Dim cachewfightgroup As New carrygroup_class

        Dim getplaneid As Integer
        Dim existed As New Collection

        carrygroup.clear()
        auxiliarygruop.clear()

        wfightergroup.clear()
        wboomergroup.clear()
        wplanegroup.clear()

        For a = 0 To 19     '第一次装填容器
            If cvship(a).active = True Then
                If restrictcontrol.getrestrictattribute(a, 9) > 0 Then     '装载水爆容器
                    For b = 1 To restrictcontrol.getrestrictattribute(a, 9)
                        wboomergroup.setcarry(cvship(a).getcarry(11))
                    Next
                End If

                If restrictcontrol.getrestrictattribute(a, 8) > 0 Then    '装载水上机容器
                    If cvshiptype.getattribute(cvship(a).uniquecode, 19) = 0 Then
                        For b = 1 To restrictcontrol.getrestrictattribute(a, 8)
                            wplanegroup.setcarry(cvship(a).getcarry(11))
                        Next
                    End If
                End If
            End If
        Next

        wboomergroup.sort()
        wplanegroup.sort()

        For a = 0 To 1            '第一次装载水战容器
            auxiliarygruop.clear()
            For b = 0 To 19
                If cvship(b).active = True Then
                    If restrictcontrol.getrestrictattribute(b, 8) > 0 Then     '第二次装载水战容器
                        If cvshiptype.getattribute(cvship(b).uniquecode, 21) = 4 - a Then
                            If cvshiptype.getattribute(cvship(b).uniquecode, 21) >= 3 And restrictcontrol.getrestrictattribute(b, 8) <= 2 Then
                                For c = 1 To restrictcontrol.getrestrictattribute(b, 8)
                                    auxiliarygruop.setcarry(cvship(b).getcarry(11), 11)
                                Next
                            End If
                        End If
                    End If
                End If
            Next
            auxiliarygruop.sort()
            If auxiliarygruop.carrycount <> 0 Then
                For b = 0 To auxiliarygruop.carrycount - 1
                    wfightergroup.setcarry(auxiliarygruop.getcarry(b))
                Next
            End If
        Next
        auxiliarygruop.clear()

        For a = 0 To 3         '第二次装载水战容器
            auxiliarygruop.clear()
            For b = 0 To 19
                If cvship(b).active = True Then
                    If restrictcontrol.getrestrictattribute(b, 8) > 0 Then
                        If cvshiptype.getattribute(cvship(b).uniquecode, 21) = 4 - a Then
                            If Not (cvshiptype.getattribute(cvship(b).uniquecode, 21) >= 3 And restrictcontrol.getrestrictattribute(b, 8) <= 2) Then
                                For c = 1 To restrictcontrol.getrestrictattribute(b, 8)
                                    auxiliarygruop.setcarry(cvship(b).getcarry(11), 11)
                                Next
                            End If
                        End If
                    End If
                End If
            Next
            auxiliarygruop.sort()
            If auxiliarygruop.carrycount <> 0 Then
                For b = 0 To auxiliarygruop.carrycount - 1
                    wfightergroup.setcarry(auxiliarygruop.getcarry(b))
                Next
            End If
        Next
        auxiliarygruop.clear()

        If wfightergroup.carrycount <> 0 Then
            For a = 0 To wfightergroup.carrycount - 1
                cachewfightgroup.setcarry(wfightergroup.getcarry(a), 11)
            Next
        End If
        If wboomergroup.carrycount <> 0 Then
            For a = 0 To wboomergroup.carrycount - 1
                cachewboomergroup.setcarry(wboomergroup.getcarry(a), 11)
            Next
        End If
        If wplanegroup.carrycount <> 0 Then
            For a = 0 To wplanegroup.carrycount - 1
                cachewplanegroup.setcarry(wplanegroup.getcarry(a), 11)
            Next
        End If

        For l = 0 To 1
            wfightergroup.clear()
            wboomergroup.clear()
            wplanegroup.clear()

            If cachewfightgroup.carrycount <> 0 Then
                For a = 0 To cachewfightgroup.carrycount - 1
                    vessel = cachewfightgroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        wfightergroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachewboomergroup.carrycount <> 0 Then
                For a = 0 To cachewboomergroup.carrycount - 1
                    vessel = cachewboomergroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        wboomergroup.setcarry(vessel, 11)
                    End If
                Next
            End If
            If cachewplanegroup.carrycount <> 0 Then
                For a = 0 To cachewplanegroup.carrycount - 1
                    vessel = cachewplanegroup.getcarry(a)
                    If vessel.shipid >= 0 + l * 10 And vessel.shipid <= 9 + l * 10 Then
                        wplanegroup.setcarry(vessel, 11)
                    End If
                Next
            End If


            If wboomergroup.carrycount <> 0 Then       '第一次装载水爆
                For a = 0 To wboomergroup.carrycount - 1
                    getplaneid = plane.getplane(12)
                    If getplaneid <> 0 Then
                        vessel = wboomergroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            vessel.planeid = getplaneid
                            wboomergroup.setcarry(vessel)
                            plane.addcollection(getplaneid, 1)
                        End If
                    End If
                Next
            End If

            If l = 0 Then           '计算需求的制空值
                expectaavalue = ffexpectaavalue
            ElseIf l = 1 Then
                expectaavalue = cfexpectaavalue - ffwaavalue
            End If

            Dim stageexpectaavalue As Integer                  '轮装水战
            stageexpectaavalue = expectaavalue - wfightergroup.AAvalue - wboomergroup.AAvalue - wplanegroup.AAvalue

            Dim round As Integer = 0
            Dim getcarrycount As Integer = 0



            getplaneid = plane.getplane(11)
            Do While carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue And getplaneid <> 0 And round <= wfightergroup.carrycount
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.clearcollection()
                getcarrycount = 0
                If wfightergroup.carrycount <> 0 Then
                    For a = 0 To wfightergroup.carrycount - 1
                        vessel = wfightergroup.getcarry(wfightergroup.carrycount - 1 - a)
                        If vessel.planeid = 0 Then
                            If getcarrycount < round Then
                                carrygroup.setcarry(vessel, 11)
                                getcarrycount = getcarrycount + 1
                            End If
                        End If
                    Next
                End If
                carrygroup.sort()
                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        getplaneid = plane.getplane(11)
                        If getplaneid <> 0 Then
                            vessel = carrygroup.getcarry(a)
                            auxiliarygruop.setcarry(vessel, 11)
                            vessel.planeid = getplaneid
                            carrygroup.setcarry(vessel, 11)
                            plane.addcollection(getplaneid)
                        End If
                    Next
                End If
                round = round + 1
                getplaneid = plane.getplane(11)
            Loop

            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    wfightergroup.setcarry(carrygroup.getcarry(a))
                Next
                For a = 0 To auxiliarygruop.carrycount - 1
                    plane.removecollection(auxiliarygruop.getcarry(a).planeid)
                Next
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.movecollection()
            End If

            stageexpectaavalue = expectaavalue - wfightergroup.AAvalue - wboomergroup.AAvalue - wplanegroup.AAvalue

            If stageexpectaavalue > 0 Then     '如果制空不足，卸载已装载的水爆
                If wboomergroup.carrycount <> 0 Then
                    For a = 0 To wboomergroup.carrycount - 1
                        vessel = wboomergroup.getcarry(a)
                        If vessel.planeid <> 0 Then
                            auxiliarygruop.setcarry(vessel, 11)
                            vessel.planeid = 0
                            wboomergroup.setcarry(vessel, 11)
                        End If
                    Next
                    If auxiliarygruop.carrycount <> 0 Then
                        For a = 0 To auxiliarygruop.carrycount - 1
                            plane.removecollection(auxiliarygruop.getcarry(a).planeid)
                        Next
                        For a = 0 To auxiliarygruop.carrycount - 1
                            auxiliarygruop.removecarry(auxiliarygruop.getcarry(a))
                        Next
                    End If
                    auxiliarygruop.clear()
                End If
            End If

            If wfightergroup.carrycount <> 0 Then
                For a = 0 To wfightergroup.carrycount - 1
                    vessel = wfightergroup.getcarry(a)
                    If vessel.planeid = 0 Then
                        carrygroup.setcarry(vessel, 11)
                    End If
                Next
                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        vessel = carrygroup.getcarry(a)
                        wfightergroup.removecarry(vessel, 11)
                        wplanegroup.setcarry(vessel, 11)
                    Next
                    carrygroup.clear()
                End If
            End If

            wplanegroup.sort()

            stageexpectaavalue = expectaavalue - wfightergroup.AAvalue - wboomergroup.AAvalue - wplanegroup.AAvalue
            round = 0
            getcarrycount = 0

            getplaneid = plane.getplane(12)
            Do While carrygroup.AAvalue < stageexpectaavalue + auxiliarygruop.AAvalue And getplaneid <> 0 And round <= wplanegroup.carrycount   '第一次轮装水爆
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.clearcollection()
                getcarrycount = 0
                If wboomergroup.carrycount <> 0 Then
                    For a = 0 To wboomergroup.carrycount - 1
                        carrygroup.setcarry(wboomergroup.getcarry(a), 11)
                    Next
                End If
                If wplanegroup.carrycount <> 0 Then
                    For a = 0 To wplanegroup.carrycount - 1
                        vessel = wplanegroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            If getcarrycount < round Then
                                carrygroup.setcarry(vessel, 11)
                                getcarrycount = getcarrycount + 1
                            End If
                        End If
                    Next
                End If
                carrygroup.sort()
                If carrygroup.carrycount <> 0 Then
                    For a = 0 To carrygroup.carrycount - 1
                        getplaneid = plane.getplane(12)
                        If getplaneid <> 0 Then
                            vessel = carrygroup.getcarry(a)
                            auxiliarygruop.setcarry(vessel, 11)
                            vessel.planeid = getplaneid
                            carrygroup.setcarry(vessel, 11)
                            plane.addcollection(getplaneid)
                        End If
                    Next
                End If
                round = round + 1
                getplaneid = plane.getplane(12)
            Loop

            If carrygroup.carrycount <> 0 Then
                For a = 0 To carrygroup.carrycount - 1
                    wplanegroup.setcarry(carrygroup.getcarry(a))
                Next
                For a = 0 To auxiliarygruop.carrycount - 1
                    plane.removecollection(auxiliarygruop.getcarry(a).planeid)
                Next
                carrygroup.clear()
                auxiliarygruop.clear()
                plane.movecollection()
            End If


            If CheckBox11.Checked = True Then    '若勾选水爆开幕极限输出，向任意水上机容器中未使用的搭载格装填水爆
                If wplanegroup.carrycount <> 0 Then
                    For a = 0 To wplanegroup.carrycount - 1
                        vessel = wplanegroup.getcarry(a)
                        If vessel.planeid = 0 Then
                            getplaneid = plane.getplane(12, 0, 1)
                            If getplaneid <> 0 Then
                                vessel.planeid = getplaneid
                                wplanegroup.setcarry(vessel)
                                plane.addcollection(getplaneid, 1)
                            Else
                                plUIcontrol.adderror(12)
                            End If
                        End If
                    Next
                End If
            End If

            '将水战/水爆/任意水上机容器中的飞机反装填回舰娘容器
            If wfightergroup.carrycount <> 0 Then
                For a = 0 To wfightergroup.carrycount - 1
                    vessel = wfightergroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If
            If wboomergroup.carrycount <> 0 Then
                For a = 0 To wboomergroup.carrycount - 1
                    vessel = wboomergroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If
            If wplanegroup.carrycount <> 0 Then
                For a = 0 To wplanegroup.carrycount - 1
                    vessel = wplanegroup.getcarry(a)
                    cvship(vessel.shipid).setcarry(vessel)
                Next
            End If

            If l = 0 Then
                ffwaavalue = wfightergroup.AAvalue + wboomergroup.AAvalue + wplanegroup.AAvalue
            ElseIf l = 1 Then
                sfwaavalue = wfightergroup.AAvalue + wboomergroup.AAvalue + wplanegroup.AAvalue
            End If
        Next

    End Sub

    Private Sub wfightermajorization()
        Dim getplaneid As Integer
        carrygroup.clear()
        auxiliarygruop.clear()

        For a = 0 To 9
            If cvship(a).active = True Then
                For b = 0 To cvship(a).carrycount - 1
                    vessel = cvship(a).getcarry(b)
                    If plane.getattribute(vessel.planeid, 3) = 11 Then
                        carrygroup.setcarry(vessel, 11)
                    End If
                Next
            End If
        Next
        carrygroup.sort()
        If carrygroup.carrycount <> 0 Then
            For a = 0 To carrygroup.carrycount - 1
                vessel = carrygroup.getcarry(a)
                If ffcaavalue + ffwaavalue - vessel.AAvalue >= ordinaryfleetaavalue Then
                    If ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue - vessel.AAvalue >= combinedfleetaavalue Then
                        ffwaavalue = ffwaavalue - vessel.AAvalue
                        plane.removecollection(vessel.planeid)
                        If CheckBox11.Checked = False Then
                            vessel.planeid = 0
                        Else
                            getplaneid = plane.getplane(12, 0, 1)
                            If getplaneid <> 0 Then
                                vessel.planeid = getplaneid
                                plane.addcollection(getplaneid, 1)
                                ffwaavalue = ffwaavalue + vessel.AAvalue
                            End If
                        End If
                        carrygroup.setcarry(vessel)
                    End If
                End If
            Next
        End If


        For a = 10 To 19
            If cvship(a).active = True Then
                For b = 0 To cvship(a).carrycount - 1
                    vessel = cvship(a).getcarry(b)
                    If plane.getattribute(vessel.planeid, 3) = 11 Then
                        auxiliarygruop.setcarry(vessel, 11)
                    End If
                Next
            End If
        Next
        auxiliarygruop.sort()
        If auxiliarygruop.carrycount <> 0 Then
            For a = 0 To auxiliarygruop.carrycount - 1
                vessel = auxiliarygruop.getcarry(a)
                If ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue - vessel.AAvalue >= combinedfleetaavalue Then
                    sfwaavalue = sfwaavalue - vessel.AAvalue
                    plane.removecollection(vessel.planeid)
                    If CheckBox11.Checked = False Then
                        vessel.planeid = 0
                    Else
                        getplaneid = plane.getplane(12, 0, 1)
                        If getplaneid <> 0 Then
                            vessel.planeid = getplaneid
                            plane.addcollection(getplaneid, 1)
                            sfwaavalue = sfwaavalue + vessel.AAvalue
                        End If
                    End If
                    auxiliarygruop.setcarry(vessel)
                End If
            Next
        End If

        If carrygroup.carrycount <> 0 Then
            For a = 0 To carrygroup.carrycount - 1
                vessel = carrygroup.getcarry(a)
                cvship(vessel.shipid).setcarry(vessel)
            Next
        End If
        If auxiliarygruop.carrycount <> 0 Then
            For a = 0 To auxiliarygruop.carrycount - 1
                vessel = auxiliarygruop.getcarry(a)
                cvship(vessel.shipid).setcarry(vessel)
            Next
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Items(ComboBox1.SelectedIndex).ToString = "<点击加载未改造舰娘>" Then
            ComboBox1.Items.Clear()
            cvshiptype.modification = 0
            Dim cvshiptypecount As Integer = 1
            Do While cvshiptype.getid(cvshiptypecount) <> 0
                ComboBox1.Items.Add(cvshiptype.getattribute(cvshiptype.getid(cvshiptypecount), 1))
                cvshiptypecount = cvshiptypecount + 1
            Loop
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        leading_intarget = 2

        Dim saa_map As New SAA_map_form
        saa_map.Show()
        saa_map.Top = Me.Top
        saa_map.Left = Me.Left
        Me.Hide()
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        plUIcontrol.changeenhancedmode()
    End Sub
End Class
