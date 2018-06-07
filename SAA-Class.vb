Structure baseplanedata
    Public name As String
    Public id As Integer
    Public classification As Integer
    Public fire As Integer
    Public torpedo As Integer
    Public antiaircraft As Integer
    Public antisubmarine As Integer
    Public bombing As Integer
    Public hit As Integer
    Public armor As Integer
    Public avoid As Integer
    Public spotting As Integer
    Public airrange As Integer
    Public antibombing As Integer
    Public intercept As Integer
    Public nightfighting As Integer
End Structure

Structure userplanedata
    Public name As String
    Public id As Integer
    Public classification As Integer
    Public fire As Integer
    Public torpedo As Integer
    Public antiaircraft As Double
    Public antisubmarine As Integer
    Public bombing As Integer
    Public hit As Integer
    Public armor As Integer
    Public avoid As Integer
    Public spotting As Integer
    Public airrange As Integer
    Public antibombing As Integer
    Public intercept As Integer
    Public nightfighting As Integer

    Public improve As Integer
    Public amountlist1 As Integer
    Public amountlist2 As Integer
    Public uniquecode As Integer
End Structure

Structure baseshiptypedata
    Public name As String
    Public id As Integer
    Public fire As Integer
    Public armor As Integer
    Public antisubmarine As Integer
    Public avoid As Integer
    Public spotting As Integer
    Public carry1 As Integer
    Public carry2 As Integer
    Public carry3 As Integer
    Public carry4 As Integer
    Public carry5 As Integer
    Public modification As Integer
    Public ncfighter As Integer
    Public nctorpedo As Integer
    Public ncboomer As Integer
    Public ncscount As Integer
    Public ncjet As Integer
    Public nwfighter As Integer
    Public nwboomer As Integer
    Public wfpriority As Integer
    Public restrict As Integer
End Structure

Structure baserestrictdata
    Public name As String
    Public id As Integer
    Public anycplane As Integer
    Public needcfighter As Integer
    Public needcboomer As Integer
    Public needctorpedo As Integer
    Public needlso As Integer
    Public anywplane As Integer
    Public needwfboomer As Integer
End Structure

Structure shiprestrictdata
    Public name As String
    Public id As Integer
    Public anycplane As Integer
    Public needcfighter As Integer
    Public needcboomer As Integer
    Public needctorpedo As Integer
    Public needlso As Integer
    Public anywplane As Integer
    Public needwboomer As Integer

    Public shipuniquecode As String

    Public equiplso As Integer
    Public equipcfighter As Integer
    Public equipcboomer As Integer
    Public equipctorpedo As Integer
    Public equipcjet As Integer
    Public equipwboomer As Integer

    Public equipnfighter As Integer
    Public equipntorpedo As Integer
    Public equipnothers As Integer
    '暂时没用，为今后出类型1的夜爆保留
    'Public equipnboomer As Integer 

    Public getcfightercarry As Integer
End Structure

Module SAA_module
    Public saabpd As New Xml.XmlDocument
    Public bplane As New baseplanedata
    Public bpcolle As New Collection
    Public filecontrol As New filecontrol_class

    Public saaupd As New Xml.XmlDocument
    Public uplane As New userplanedata
    Public plane As New plane_class

    Public saastd As New Xml.XmlDocument
    Public bship As New baseshiptypedata
    Public baseship As New baseship_class
    Public ship(19) As ship_class

    Public vessel As New carry_class

    Public saartd As New Xml.XmlDocument
    Public restrictcontrol As New restrictcontrol_class

    Public UIcontrol As New UIcontrol_class

    Public ordinaryfleetaavalue As Integer
    Public combinedfleetaavalue As Integer
    Public ffcaavalue As Integer
    Public sfcaavalue As Integer
    Public ffwaavalue As Integer
    Public sfwaavalue As Integer

    Public leading_intarget As Integer

    Public carrygroup As New carrygroup_class
    Public auxiliarygruop As New carrygroup_class

    Public cplanegroup As New carrygroup_class
    Public cfightergroup As New carrygroup_class
    Public cboomergroup As New carrygroup_class
    Public ctorpedogroup As New carrygroup_class
    Public clsogroup As New carrygroup_class

    Public wplanegroup As New carrygroup_class
    Public wfightergroup As New carrygroup_class
    Public wboomergroup As New carrygroup_class


End Module

Public Class plane_class
    'n-海军飞机;a-陆军飞机
    'c-舰载;w-水上机;l-陆基
    'Dim ncfighter As New Collection
    'Dim ncfboomer As New Collection
    'Dim ncboomer As New Collection
    'Dim nctorpedo As New Collection
    'Dim ncjet As New Collection
    'Dim ncscount As New Collection
    'Dim nwfighter As New Collection
    'Dim nwboomer As New Collection
    'Dim lso As New Collection
    'Dim flyingboat As New Collection
    'Dim nltorpedo As New Collection
    'Dim nalfighter As New Collection

    Dim planelist As New Collection

    Dim uselist As New Collection
    Dim temuselist As New Collection

    Sub New()
        If System.IO.File.Exists(Application.StartupPath + "\data\SAAuserplanedata.xml") Then
            saaupd.Load(Application.StartupPath + "\data\SAAuserplanedata.xml")
        Else
            Dim xmlstring As New System.Text.StringBuilder
            xmlstring.Append("<?xml version=""1.0"" encoding=""gb2312""?>")
            xmlstring.Append("<userplanedata>")
            xmlstring.Append("</userplanedata>")
            saaupd.LoadXml(xmlstring.ToString)
            saaupd.DocumentElement.SetAttribute("version", "00000000")
            saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
        End If
    End Sub

    Public Sub loadplanedata(ByVal x As Integer)
        planelist.Clear()
        For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
            uplane.name = upnode.Attributes("name").Value
            uplane.id = upnode.Attributes("id").Value
            uplane.classification = upnode.Attributes("classification").Value
            uplane.fire = upnode.Attributes("fire").Value
            uplane.torpedo = upnode.Attributes("torpedo").Value
            uplane.antiaircraft = upnode.Attributes("antiaircraft").Value
            uplane.antisubmarine = upnode.Attributes("antisubmarine").Value
            uplane.bombing = upnode.Attributes("bombing").Value
            uplane.hit = upnode.Attributes("hit").Value
            uplane.armor = upnode.Attributes("armor").Value
            uplane.avoid = upnode.Attributes("avoid").Value
            uplane.spotting = upnode.Attributes("spotting").Value
            uplane.airrange = upnode.Attributes("airrange").Value
            uplane.antibombing = upnode.Attributes("antibombing").Value
            uplane.intercept = upnode.Attributes("intercept").Value
            uplane.nightfighting = upnode.Attributes("nightfighting").Value

            uplane.improve = upnode.Attributes("improve").Value
            uplane.amountlist1 = upnode.Attributes("amountlist1").Value
            uplane.amountlist2 = upnode.Attributes("amountlist2").Value

            Dim list As Integer
            If x = 0 Then
                list = uplane.amountlist1
            ElseIf x = 1 Then
                list = uplane.amountlist2
            End If

            If list > 0 Then
                For a = 1 To list
                    uplane.uniquecode = uplane.id * 10000 + uplane.improve * 100 + a
                    planelist.Add(uplane, uplane.uniquecode)
                Next
            End If
        Next
    End Sub

    Public Function getplane(ByVal x As Integer, Optional ByVal y As Double = 0, Optional ByVal z As Double = 0, Optional ByVal zz As Double = 0) As Integer
        getplane = 0
        uplane.uniquecode = "0"
        Dim fattribute As Double = y
        Dim sattribute As Double = 0

        If x = 0 Then fattribute = 0


        Dim unuselist As New Collection
        'Dim alternativelist As New Collection
        Dim fattributegroup As New Collection
        Dim sattributegroup As New Collection


        For a = 1 To planelist.Count
            If used(planelist(a).uniquecode) = False Then
                If x = 0 And planelist(a).nightfighting >= 1 Then '取得夜战机
                    If y = 0 And planelist(a).classification = 1 Then 'y0取得夜战   
                        unuselist.Add(planelist(a))
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If z = 0 Then 'z0优先对潜 
                            sattributegroup.Add(planelist(a).antisubmarine)
                        End If
                    ElseIf y = 1 And planelist(a).classification = 2 Then 'y1取得夜爆战
                        unuselist.Add(planelist(a))
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If z = 0 Then 'z0优先爆装  
                            sattributegroup.Add(planelist(a).bombing)
                        End If
                    ElseIf y = 2 And planelist(a).classification = 3 Then 'y1取得夜爆
                        unuselist.Add(planelist(a))
                        fattributegroup.Add(planelist(a).bombing)
                        If z = 0 Then 'z0优先对空  
                            sattributegroup.Add(planelist(a).antiaircraft)
                        End If
                    ElseIf y = 3 And planelist(a).classification = 4 Then 'y1取得夜攻
                        unuselist.Add(planelist(a))
                        fattributegroup.Add(planelist(a).torpedo)
                        If z = 0 Then 'z0优先对空  
                            sattributegroup.Add(planelist(a).antiaircraft)
                        End If
                    End If

                ElseIf x = 1 And planelist(a).classification = 1 Then  '取得舰战
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then 'zz0副优先闪避
                            sattributegroup.Add(planelist(a).avoid)
                        End If
                    ElseIf z = 1 Then 'z1主优先闪避
                        fattributegroup.Add(planelist(a).avoid)
                        If zz = 0 Then 'zz0副优先对空
                            sattributegroup.Add(planelist(a).antiaircraft)
                        End If
                    End If
                ElseIf x = 2 And planelist(a).classification = 2 Then  '取得爆战
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then 'zz0副优先爆装
                            sattributegroup.Add(planelist(a).bombing)
                        End If
                    End If
                ElseIf x = 3 And planelist(a).classification = 3 Then  '取得舰爆
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先爆装
                        fattributegroup.Add(planelist(a).bombing)
                        If zz = 0 Then  'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    End If
                ElseIf x = 4 And planelist(a).classification = 4 Then  '取得舰攻
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先雷装
                        fattributegroup.Add(planelist(a).torpedo)
                        If zz = 0 Then 'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    End If
                ElseIf x = 5 And planelist(a).classification = 5 Then  '取得喷气机
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then 'zz0主优先爆装
                            sattributegroup.Add(planelist(a).bombing)
                        End If
                    End If
                ElseIf x = 6 And planelist(a).classification = 6 Then  '取得彩云
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先索敌
                        fattributegroup.Add(planelist(a).spotting)
                        If zz = 0 Then 'zz0副优先火力
                            sattributegroup.Add(planelist(a).fire)
                        End If
                    End If
                ElseIf x = 7 And planelist(a).classification = 7 Then  '取得二式舰侦
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先索敌
                        fattributegroup.Add(planelist(a).spotting)
                        If zz = 0 Then 'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    End If
                ElseIf x = 8 And planelist(a).classification = 8 Then  '取得夜战要员
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先火力
                        fattributegroup.Add(planelist(a).fire)
                        If zz = 0 Then 'zz0副优先回避
                            sattributegroup.Add(planelist(a).avoid)
                        End If
                    End If
                ElseIf x = 11 And planelist(a).classification = 11 Then  '取得水战
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then 'zz0副优先回避
                            sattributegroup.Add(planelist(a).avoid)
                        End If
                    End If
                ElseIf x = 12 And planelist(a).classification = 12 Then  '取得水爆
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then  'zz0副优先爆装
                            sattributegroup.Add(planelist(a).bombing)
                        End If
                    ElseIf z = 1 Then 'z1主优先爆装
                        fattributegroup.Add(planelist(a).bombing)
                        If zz = 0 Then 'zz0副优先对潜
                            sattributegroup.Add(planelist(a).antisubmarine)
                        End If
                    End If
                ElseIf x = 13 And planelist(a).classification = 13 Then  '取得飞行艇
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先索敌
                        fattributegroup.Add(planelist(a).spotting)
                        If zz = 0 Then 'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    End If
                ElseIf x = 21 And planelist(a).classification = 21 Then  '取得司令部
                    unuselist.Add(planelist(a))
                    If z = 0 Then 'z0主优先对空
                        fattributegroup.Add(planelist(a).antiaircraft)
                        If zz = 0 Then 'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    ElseIf z = 1 Then 'z1主优先装甲
                        fattributegroup.Add(planelist(a).armor)
                        If zz = 0 Then 'zz0副优先命中
                            sattributegroup.Add(planelist(a).hit)
                        End If
                    End If
                End If
            End If

        Next



        If unuselist.Count <> 0 Then
            If fattribute = 0 Then
                For a = 1 To unuselist.Count
                    If fattributegroup(a) > fattribute Then
                        fattribute = fattributegroup(a)
                    End If
                Next
            End If
            For a = 1 To unuselist.Count
                If fattributegroup(a) = fattribute Then
                    If sattributegroup(a) >= sattribute Then
                        uplane = unuselist(a)
                        sattribute = sattributegroup(a)
                    End If
                End If
            Next
            getplane = uplane.uniquecode
        End If

    End Function

    Public Function getattribute(ByVal x As Integer, ByVal y As Integer)
        getattribute = 0
        If x <> 0 Then
            Dim uniquecode As String = Trim(x)
            If y = 1 Then
                getattribute = planelist.Item(uniquecode).name
            ElseIf y = 2 Then
                getattribute = planelist.Item(uniquecode).id
            ElseIf y = 3 Then
                getattribute = planelist.Item(uniquecode).classification
            ElseIf y = 4 Then
                getattribute = planelist.Item(uniquecode).fire
            ElseIf y = 5 Then
                getattribute = planelist.Item(uniquecode).torpedo
            ElseIf y = 6 Then
                getattribute = planelist.Item(uniquecode).antiaircraft
            ElseIf y = 7 Then
                getattribute = planelist.Item(uniquecode).antisubmarine
            ElseIf y = 8 Then
                getattribute = planelist.Item(uniquecode).bombing
            ElseIf y = 9 Then
                getattribute = planelist.Item(uniquecode).hit
            ElseIf y = 10 Then
                getattribute = planelist.Item(uniquecode).armor
            ElseIf y = 11 Then
                getattribute = planelist.Item(uniquecode).avoid
            ElseIf y = 12 Then
                getattribute = planelist.Item(uniquecode).spotting
            ElseIf y = 13 Then
                getattribute = planelist.Item(uniquecode).airrange
            ElseIf y = 14 Then
                getattribute = planelist.Item(uniquecode).antibombing
            ElseIf y = 15 Then
                getattribute = planelist.Item(uniquecode).intercept
            ElseIf y = 16 Then
                getattribute = planelist.Item(uniquecode).nightfighting
            ElseIf y = 17 Then
                getattribute = planelist.Item(uniquecode).improve
            End If
        End If
    End Function

    Public Function getnfequipnum(ByVal x As Integer) As Integer
        getnfequipnum = 0
        If x = 0 Then
            For a = 1 To planelist.Count
                If planelist(a).classification = 8 Then
                    getnfequipnum = getnfequipnum + 1
                End If
            Next
        ElseIf x > 0 Then
            For a = 1 To planelist.Count - 1
                If planelist(a).nightfighting = x Then
                    getnfequipnum = getnfequipnum + 1
                End If
            Next
        End If
    End Function

    Private Function used(ByVal x As Integer) As Boolean
        used = False
        If uselist.Count <> 0 Then
            For a = 1 To uselist.Count
                If x = uselist(a) Then
                    used = True
                End If
            Next
        End If
        If temuselist.Count <> 0 Then
            For a = 1 To temuselist.Count
                If x = temuselist(a) Then
                    used = True
                End If
            Next
        End If
    End Function

    Public Sub addcollection(ByVal x As Integer, Optional ByVal y As Integer = 0)
        If y = 0 Then
            temuselist.Add(x)
        Else
            uselist.Add(x)
        End If
    End Sub

    Public Sub removecollection(ByVal x As Integer)
        For a = 1 To temuselist.Count
            If temuselist(a) = x Then
                temuselist.Remove(a)
                Exit For
            End If
        Next
        For a = 1 To uselist.Count
            If uselist(a) = x Then
                uselist.Remove(a)
                Exit For
            End If
        Next
    End Sub
    Public Sub movecollection()
        If temuselist.Count <> 0 Then
            For a = 1 To temuselist.Count
                uselist.Add(temuselist.Item(a))
            Next
            temuselist.Clear()
        End If
    End Sub

    Public Sub clearcollection(Optional ByVal x As Integer = 0)
        If x = 0 Then
            temuselist.Clear()
        ElseIf x = 1 Then
            temuselist.Clear()
            uselist.Clear()
        End If
    End Sub
End Class

Public Class baseship_class

    Dim bshiplist As New Collection

    Public Sub New()
        saastd.Load(Application.StartupPath + "\data\SAAshiptypedata.xml")
        Call loadshipdate()
    End Sub

    Public Sub loadshipdate(Optional ByVal x As Integer = 1)
        bshiplist.Clear()

        For Each upnode As Xml.XmlElement In saastd.DocumentElement.ChildNodes
            bship.name = upnode.Attributes("name").Value
            bship.id = upnode.Attributes("id").Value
            bship.fire = upnode.Attributes("fire").Value
            bship.armor = upnode.Attributes("armor").Value
            bship.antisubmarine = upnode.Attributes("antisubmarine").Value
            bship.avoid = upnode.Attributes("avoid").Value
            bship.spotting = upnode.Attributes("spotting").Value
            bship.carry1 = upnode.Attributes("carry1").Value
            bship.carry2 = upnode.Attributes("carry2").Value
            bship.carry3 = upnode.Attributes("carry3").Value
            bship.carry4 = upnode.Attributes("carry4").Value
            bship.carry5 = upnode.Attributes("carry5").Value
            bship.modification = upnode.Attributes("modification").Value
            bship.ncfighter = upnode.Attributes("ncfighter").Value
            bship.nctorpedo = upnode.Attributes("nctorpedo").Value
            bship.ncboomer = upnode.Attributes("ncboomer").Value
            bship.ncscount = upnode.Attributes("ncscount").Value
            bship.ncjet = upnode.Attributes("ncjet").Value
            bship.nwfighter = upnode.Attributes("nwfighter").Value
            bship.nwboomer = upnode.Attributes("nwboomer").Value
            bship.wfpriority = upnode.Attributes("wfpriority").Value
            bship.restrict = upnode.Attributes("restrict").Value

            If bship.modification >= x Then
                bshiplist.Add(bship, bship.id)
            End If
        Next
    End Sub

    Public Function getattribute(ByVal x As String, ByVal y As Integer)
        getattribute = 0
        If Val(x) > 0 Then
            Dim uniquecode As String = Trim(Val(x))

            If y = 1 Then
                getattribute = bshiplist.Item(uniquecode).name
            ElseIf y = 2 Then
                getattribute = bshiplist.Item(uniquecode).id
            ElseIf y = 3 Then
                getattribute = bshiplist.Item(uniquecode).fire
            ElseIf y = 4 Then
                getattribute = bshiplist.Item(uniquecode).armor
            ElseIf y = 5 Then
                getattribute = bshiplist.Item(uniquecode).antisubmarine
            ElseIf y = 6 Then
                getattribute = bshiplist.Item(uniquecode).avoid
            ElseIf y = 7 Then
                getattribute = bshiplist.Item(uniquecode).spotting
            ElseIf y = 8 Then
                getattribute = bshiplist.Item(uniquecode).carry1
            ElseIf y = 9 Then
                getattribute = bshiplist.Item(uniquecode).carry2
            ElseIf y = 10 Then
                getattribute = bshiplist.Item(uniquecode).carry3
            ElseIf y = 11 Then
                getattribute = bshiplist.Item(uniquecode).carry4
            ElseIf y = 12 Then
                getattribute = bshiplist.Item(uniquecode).carry5
            ElseIf y = 13 Then
                getattribute = bshiplist.Item(uniquecode).modification
            ElseIf y = 14 Then
                getattribute = bshiplist.Item(uniquecode).ncfighter
            ElseIf y = 15 Then
                getattribute = bshiplist.Item(uniquecode).nctorpedo
            ElseIf y = 16 Then
                getattribute = bshiplist.Item(uniquecode).ncboomer
            ElseIf y = 17 Then
                getattribute = bshiplist.Item(uniquecode).ncscount
            ElseIf y = 18 Then
                getattribute = bshiplist.Item(uniquecode).ncjet
            ElseIf y = 19 Then
                getattribute = bshiplist.Item(uniquecode).nwfighter
            ElseIf y = 20 Then
                getattribute = bshiplist.Item(uniquecode).nwboomer
            ElseIf y = 21 Then
                getattribute = bshiplist.Item(uniquecode).wfpriority
            ElseIf y = 22 Then
                getattribute = bshiplist.Item(uniquecode).restrict
            End If
        End If

    End Function

    Public Function getid(ByVal x As Integer) As Integer
        Try
            getid = bshiplist(x).id
        Catch
            getid = 0
        End Try
    End Function
End Class

Public Class ship_class
    Dim idvalue As Integer
    Dim code As String
    Dim activevalue As Boolean = False
    Dim restrictvalue As Integer
    Dim carrycountvalue As Integer = 0
    Dim carry(4) As carry_class
    Dim carryobtained(4) As Integer

    Public Sub New(ByVal x As Integer)
        For a = 0 To 4
            carry(a) = New carry_class
        Next
        idvalue = x
    End Sub

    Public Property uniquecode As String
        Get
            uniquecode = code
        End Get
        Set(value As String)
            code = value
            If code = "" Then
                active = False
            Else
                active = True
            End If
        End Set
    End Property

    ReadOnly Property carrycount As Integer
        Get
            carrycount = carrycountvalue
        End Get
    End Property

    Public Sub setcarry(ByVal x As carry_class)
        Dim rewrite As Boolean = False
        For a = 0 To 4
            If x.uniquecode = carry(a).uniquecode Then
                Call copy(x, carry(a))
                rewrite = True
                Exit For
            End If
        Next
        If rewrite = False Then
            Call copy(x, carry(carrycountvalue))
            carrycountvalue = carrycountvalue + 1
        End If
    End Sub

    Public Function getcarry(ByVal x As Integer) As carry_class
        Dim transport As New carry_class
        Dim extreme(1) As Integer
        If x >= 0 And x <= 4 Then    '按序号取得搭载格
            Call copy(carry(x), transport)
        ElseIf x < 0 Then   '取得最小的未取得过的搭载格
            extreme(0) = 200
            For a = 0 To carrycountvalue - 1
                If carryobtained(a) = 0 And carry(a).amount < extreme(0) Then
                    extreme(0) = carry(a).amount
                    extreme(1) = a
                End If
            Next
            Call copy(carry(extreme(1)), transport)
            carryobtained(extreme(1)) = 1
        ElseIf x > 4 Then   '取得最大的未取得过的搭载格
            extreme(0) = 0
            For a = 0 To carrycountvalue - 1
                If carryobtained(carrycountvalue - 1 - a) = 0 And carry(carrycountvalue - 1 - a).amount > extreme(0) Then
                    extreme(0) = carry(carrycountvalue - 1 - a).amount
                    extreme(1) = carrycountvalue - 1 - a

                End If
            Next
            Call copy(carry(extreme(1)), transport)
            carryobtained(extreme(1)) = 1
        End If
        getcarry = transport
    End Function

    Private Sub copy(ByVal x As carry_class, ByVal y As carry_class)
        y.amount = x.amount
        y.shipid = x.shipid
        y.carryid = x.carryid
        y.planeid = x.planeid
        y.uniquecode = x.uniquecode
    End Sub

    Public Property restrict As Integer
        Get
            restrict = restrictvalue
        End Get
        Set(value As Integer)
            restrictvalue = value
            If restrictvalue <> 0 Then
                restrictcontrol.initialize(idvalue, code, restrictvalue)
            End If
        End Set
    End Property

    Public Property active As Boolean
        Get
            active = activevalue
        End Get
        Set(value As Boolean)
            activevalue = value
        End Set
    End Property

    ReadOnly Property AAvalue As Integer
        Get
            Dim value As Integer = 0

            For a = 0 To carrycountvalue - 1
                value = value + carry(a).AAvalue
            Next
            AAvalue = value
        End Get
    End Property

    ReadOnly Property minaviationfire As Double
        Get
            minaviationfire = 0
            For a = 0 To carrycountvalue - 1
                minaviationfire = minaviationfire + carry(a).minaviationfire
            Next
        End Get
    End Property

    ReadOnly Property maxaviationfire As Double
        Get
            maxaviationfire = 0
            For a = 0 To carrycountvalue - 1
                maxaviationfire = maxaviationfire + carry(a).maxaviationfire
            Next
        End Get
    End Property

    ReadOnly Property shellingfire As Integer
        Get
            shellingfire = -1
            If restrictvalue < 180 Then
                If restrictcontrol.getrestrictattribute(idvalue, 15) + restrictcontrol.getrestrictattribute(idvalue, 16) + restrictcontrol.getrestrictattribute(idvalue, 17) <= 0 Then
                    shellingfire = 0
                Else
                    Dim fire As Integer = baseship.getattribute(code, 3)
                    Dim booming As Integer
                    Dim torpedo As Integer
                    For a = 0 To carrycountvalue - 1
                        If carry(a).planeid <> 0 Then
                            fire = fire + plane.getattribute(carry(a).planeid, 4)
                            booming = booming + plane.getattribute(carry(a).planeid, 8)
                            torpedo = torpedo + plane.getattribute(carry(a).planeid, 5)
                        End If
                    Next
                    shellingfire = Int((fire + booming * 1.3 + torpedo) * 1.5 + 55)
                End If
            End If
        End Get
    End Property

    ReadOnly Property CIshellingfire As Integer
        Get
            CIshellingfire = -1
            If restrictvalue < 180 Then
                CIshellingfire = shellingfire * restrictcontrol.CIstate(carry(0))
            End If
        End Get
    End Property

    ReadOnly Property nightfightfire As Integer
        Get
            nightfightfire = -1
            If restrictvalue < 180 Then
                nightfightfire = 0
                If restrictcontrol.nfstate(carry(0)) <> 0 Then
                    Dim fire As Integer = baseship.getattribute(code, 3)
                    Dim torpedo As Integer
                    Dim nfcorrection As Double
                    For a = 0 To carrycountvalue - 1
                        If carry(a).planeid <> 0 Then
                            If plane.getattribute(carry(a).planeid, 16) <> 0 Then
                                fire = fire + plane.getattribute(carry(a).planeid, 4)
                                torpedo = torpedo + plane.getattribute(carry(a).planeid, 5)
                                nfcorrection = nfcorrection + carry(a).nfcorrection
                            End If
                        End If
                    Next
                    nightfightfire = Int(fire + torpedo + nfcorrection)
                End If
            End If
        End Get
    End Property

    ReadOnly Property CInightfightfire As Integer
        Get
            CInightfightfire = -1
            If restrictvalue < 180 Then
                If restrictcontrol.nfstate(carry(0)) <> 0 Then
                    CInightfightfire = Int(nightfightfire * restrictcontrol.nfstate(carry(0)))
                ElseIf restrictcontrol.nfstate(carry(0)) <= 1 Then
                    CInightfightfire = 0
                End If
            End If
        End Get
    End Property

    Public Sub resetcarryobtained()
        For a = 0 To 4
            If carryobtained(a) <> 2 Then
                carryobtained(a) = 0
                carry(a).planeid = 0
            End If
        Next
    End Sub

    Public Sub clear()
        uniquecode = ""
        For a = 0 To 4
            carry(a).clear()
            carryobtained(a) = 0
        Next
        restrict = 0
        active = False
        carrycountvalue = 0
    End Sub
End Class

Public Class carry_class
    Dim amountvalue As Integer = 0
    Dim shipidvalue As Integer = -1
    Dim carryidvalue As Integer = 0
    Dim planeidvalue As Integer = 0
    Dim uniquecodevalue As String = ""

    Public Property amount As Integer
        Get
            amount = amountvalue
        End Get
        Set(value As Integer)
            amountvalue = value
        End Set
    End Property

    Public Property shipid As Integer
        Get
            shipid = shipidvalue
        End Get
        Set(value As Integer)
            shipidvalue = value
        End Set
    End Property

    Public Property carryid As Integer
        Get
            carryid = carryidvalue
        End Get
        Set(value As Integer)
            carryidvalue = value
        End Set
    End Property

    Public Property planeid As Integer
        Get
            planeid = planeidvalue
        End Get
        Set(value As Integer)
            planeidvalue = value
        End Set
    End Property

    Public Property uniquecode As String
        Get
            uniquecode = uniquecodevalue
        End Get
        Set(value As String)
            uniquecodevalue = value
        End Set
    End Property

    ReadOnly Property AAvalue As Integer
        Get
            Dim shilledbuff As Integer
            If plane.getattribute(planeidvalue, 3) = 1 Or plane.getattribute(planeidvalue, 3) = 11 Then
                shilledbuff = 25
            ElseIf plane.getattribute(planeidvalue, 3) = 12 Or plane.getattribute(planeidvalue, 3) = 13 Then
                shilledbuff = 9
            ElseIf plane.getattribute(planeidvalue, 3) >= 2 And plane.getattribute(planeidvalue, 3) <= 5 Then
                shilledbuff = 3
            End If
            AAvalue = Int(plane.getattribute(planeidvalue, 6) * Math.Sqrt(amountvalue) + shilledbuff)
        End Get
    End Property

    ReadOnly Property minaviationfire As Double
        Get
            Dim basefire As Integer = 0
            Dim coefficient As Double = 0
            If planeidvalue <> 0 Then
                If plane.getattribute(planeidvalue, 3) = 4 Then
                    basefire = plane.getattribute(planeidvalue, 5)
                    coefficient = 0.8
                ElseIf plane.getattribute(planeidvalue, 3) = 2 Or plane.getattribute(planeidvalue, 3) = 3 Or plane.getattribute(planeidvalue, 3) = 12 Then
                    basefire = plane.getattribute(planeidvalue, 8)
                    coefficient = 1
                ElseIf plane.getattribute(planeidvalue, 3) = 5 Then
                    basefire = plane.getattribute(planeidvalue, 8)
                    coefficient = 1 / Math.Sqrt(2)
                End If
            End If
            minaviationfire = coefficient * (basefire * Math.Sqrt(amountvalue) + 25)
        End Get
    End Property

    ReadOnly Property maxaviationfire As Double
        Get
            Dim basefire As Integer
            Dim coefficient As Double
            If planeidvalue <> 0 Then
                If plane.getattribute(planeidvalue, 3) = 4 Then
                    basefire = plane.getattribute(planeidvalue, 5)
                    coefficient = 1.5
                ElseIf plane.getattribute(planeidvalue, 3) = 2 Or plane.getattribute(planeidvalue, 3) = 3 Or plane.getattribute(planeidvalue, 3) = 12 Then
                    basefire = plane.getattribute(planeidvalue, 8)
                    coefficient = 1
                ElseIf plane.getattribute(planeidvalue, 3) = 5 Then
                    basefire = plane.getattribute(planeidvalue, 8)
                    coefficient = 1 / Math.Sqrt(2)
                End If
            End If
            maxaviationfire = coefficient * (basefire * Math.Sqrt(amountvalue) + 25)
        End Get
    End Property

    ReadOnly Property nfcorrection As Double
        Get
            nfcorrection = 0
            If planeidvalue <> 0 Then
                Dim coefficientA As Double = 0
                Dim coefficientB As Double = 0
                If plane.getattribute(planeidvalue, 16) = 1 Then
                    coefficientA = 3
                    coefficientB = 0.45
                ElseIf plane.getattribute(planeidvalue, 16) = 2 Then
                    coefficientA = 0
                    coefficientB = 0.3
                End If
                nfcorrection = coefficientA * amountvalue + coefficientB * (plane.getattribute(planeidvalue, 4) + plane.getattribute(planeidvalue, 5) + plane.getattribute(planeidvalue, 8) + plane.getattribute(planeidvalue, 7)) * Math.Sqrt(amountvalue) + Math.Sqrt(plane.getattribute(planeidvalue, 17))
            End If
        End Get
    End Property

    Public Sub clear()
        amountvalue = 0
        shipidvalue = -1
        carryidvalue = 0
        planeidvalue = 0
        uniquecodevalue = ""
    End Sub
End Class

Public Class carrygroup_class
    Dim carry(100) As carry_class
    Dim carrycountvalue As Integer = 0
    Public Sub New()
        For a = 0 To 100
            carry(a) = New carry_class
        Next
    End Sub

    ReadOnly Property carrycount As Integer
        Get
            carrycount = carrycountvalue
        End Get
    End Property

    Public Sub setcarry(ByVal x As carry_class, Optional ByVal y As Integer = 0)
        Dim rewrite As Boolean = False
        For a = 0 To 99
            If x.uniquecode = carry(a).uniquecode Then
                If y = 0 And carry(a).planeid <> 0 Then restrictcontrol.removecarry(carry(a))
                Call copy(x, carry(a))
                rewrite = True
                Exit For
            End If
        Next
        If rewrite = False Then
            Call copy(x, carry(carrycountvalue))
            carrycountvalue = carrycountvalue + 1
        End If
        If y = 0 Then
            restrictcontrol.setcarry(x)
        ElseIf y > 0 And y <= 10 Then
            restrictcontrol.setcarry(x, y)
        End If
    End Sub

    Public Function getcarry(ByVal x As Integer) As carry_class
        Dim transport As New carry_class
        Call copy(carry(x), transport)
        getcarry = transport
    End Function

    Public Sub removecarry(ByVal x As carry_class, Optional ByVal y As Integer = 0)
        For a = 0 To 99
            If carry(a).uniquecode = x.uniquecode Then
                If a < carrycountvalue - 1 Then
                    For b = a + 1 To carrycountvalue - 1
                        Call copy(carry(b), carry(b - 1))
                    Next
                    carry(carrycountvalue - 1).clear()
                    carrycountvalue = carrycountvalue - 1
                    Exit For
                ElseIf a = carrycountvalue - 1 Then
                    carry(a).clear()
                    carrycountvalue = carrycountvalue - 1
                End If
            End If
        Next
        If y = 0 Then
            restrictcontrol.removecarry(x)
        ElseIf y > 0 And y <= 10 Then
            restrictcontrol.removecarry(x, y)
        End If
    End Sub

    ReadOnly Property AAvalue As Integer
        Get
            Dim value As Integer = 0
            For a = 0 To carrycountvalue - 1
                value = value + carry(a).AAvalue
            Next
            AAvalue = value
        End Get
    End Property

    Private Sub copy(ByVal x As carry_class, ByVal y As carry_class)
        y.amount = x.amount
        y.shipid = x.shipid
        y.carryid = x.carryid
        y.planeid = x.planeid
        y.uniquecode = x.uniquecode
    End Sub

    Public Sub sort()
        For a = 0 To 99
            For b = a To 99
                If carry(a).amount < carry(b).amount Then
                    Call copy(carry(a), carry(100))
                    Call copy(carry(b), carry(a))
                    Call copy(carry(100), carry(b))
                End If
            Next
        Next
    End Sub

    Public Sub clear()
        For a = 0 To 100
            carry(a).clear()
        Next
        carrycountvalue = 0
    End Sub
End Class

Public Class restrictcontrol_class
    Dim brestrict As baserestrictdata
    Dim brestrictlist As New Collection

    Dim shiprestrict(19) As shiprestrictdata

    Dim startvalue As Integer = 0

    Public Sub New()
        saartd.Load(Application.StartupPath + "\data\SAArestrictdata.xml")
        Call loadrestrictdata()
    End Sub

    Public Sub loadrestrictdata()
        For Each upnode As Xml.XmlElement In saartd.DocumentElement.ChildNodes
            brestrict.name = upnode.Attributes("name").Value
            brestrict.id = upnode.Attributes("id").Value
            brestrict.anycplane = upnode.Attributes("anycplane").Value
            brestrict.needcfighter = upnode.Attributes("needcfighter").Value
            brestrict.needcboomer = upnode.Attributes("needcboomer").Value
            brestrict.needctorpedo = upnode.Attributes("needctorpedo").Value
            brestrict.needlso = upnode.Attributes("needlso").Value
            brestrict.anywplane = upnode.Attributes("anywplane").Value
            brestrict.needwfboomer = upnode.Attributes("needwfboomer").Value

            brestrictlist.Add(brestrict, brestrict.id)
        Next
    End Sub

    Public Sub initialize(ByVal x As Integer, ByVal y As String, ByVal z As Integer)
        brestrict = brestrictlist.Item(Trim(z))
        shiprestrict(x).name = brestrict.name
        shiprestrict(x).id = brestrict.id
        shiprestrict(x).anycplane = brestrict.anycplane
        shiprestrict(x).needcfighter = brestrict.needcfighter
        shiprestrict(x).needcboomer = brestrict.needcboomer
        shiprestrict(x).needctorpedo = brestrict.needctorpedo
        shiprestrict(x).needlso = brestrict.needlso
        shiprestrict(x).anywplane = brestrict.anywplane
        shiprestrict(x).needwboomer = brestrict.needwfboomer

        shiprestrict(x).shipuniquecode = y

        If shiprestrict(x).id >= 133 And shiprestrict(x).id <= 136 Then
            shiprestrict(x).equiplso = 1
        Else
            shiprestrict(x).equiplso = 0
        End If

        shiprestrict(x).equipcfighter = 0
        shiprestrict(x).equipcboomer = 0
        shiprestrict(x).equipctorpedo = 0
        shiprestrict(x).equipcjet = 0
        shiprestrict(x).equipnfighter = 0
        shiprestrict(x).equipntorpedo = 0
        shiprestrict(x).equipnothers = 0

        shiprestrict(x).getcfightercarry = 0

        shiprestrict(x).equipwboomer = 0
    End Sub

    Public Function getbaserestrictname(ByVal x As Integer) As String
        getbaserestrictname = ""
        For a = 1 To brestrictlist.Count
            brestrict = brestrictlist(a)
            If brestrict.id = x Then
                getbaserestrictname = brestrict.name
                Exit For
            End If
        Next
    End Function

    Public Function getrestrictattribute(ByVal x As Integer, ByVal y As Integer)
        getrestrictattribute = -1
        If y = 1 Then
            getrestrictattribute = shiprestrict(x).name
        ElseIf y = 2 Then
            getrestrictattribute = shiprestrict(x).id
        ElseIf y = 3 Then
            getrestrictattribute = shiprestrict(x).anycplane
        ElseIf y = 4 Then
            getrestrictattribute = shiprestrict(x).needcfighter
        ElseIf y = 5 Then
            getrestrictattribute = shiprestrict(x).needcboomer
        ElseIf y = 6 Then
            getrestrictattribute = shiprestrict(x).needctorpedo
        ElseIf y = 7 Then
            getrestrictattribute = shiprestrict(x).needlso
        ElseIf y = 8 Then
            getrestrictattribute = shiprestrict(x).anywplane
        ElseIf y = 9 Then
            getrestrictattribute = shiprestrict(x).needwboomer
        ElseIf y = 10 Then
            getrestrictattribute = shiprestrict(x).equiplso
        ElseIf y = 11 Then
            getrestrictattribute = shiprestrict(x).equipnfighter
        ElseIf y = 12 Then
            getrestrictattribute = shiprestrict(x).equipntorpedo
        ElseIf y = 13 Then
            getrestrictattribute = shiprestrict(x).equipnothers
        ElseIf y = 14 Then
            getrestrictattribute = shiprestrict(x).getcfightercarry
        ElseIf y = 15 Then
            getrestrictattribute = shiprestrict(x).equipcboomer
        ElseIf y = 16 Then
            getrestrictattribute = shiprestrict(x).equipctorpedo
        ElseIf y = 17 Then
            getrestrictattribute = shiprestrict(x).equipcjet
        ElseIf y = 18 Then
            getrestrictattribute = shiprestrict(x).equipwboomer
        End If
    End Function
    WriteOnly Property start As Integer
        Set(value As Integer)
            startvalue = value
        End Set
    End Property

    ReadOnly Property nfstate(ByVal x As carry_class) As Double
        Get
            nfstate = -1
            Dim code As String
            code = Mid(x.uniquecode, 1, 4)
            For a = 0 To 19
                If shiprestrict(a).shipuniquecode = code Then
                    If shiprestrict(a).equiplso = 1 Then
                        nfstate = 0
                        If shiprestrict(a).equipnfighter + shiprestrict(a).equipntorpedo > 0 Then
                            nfstate = 1
                            If shiprestrict(a).equipnfighter >= 2 And shiprestrict(a).equipntorpedo > 0 Then
                                nfstate = 1.25
                            ElseIf shiprestrict(a).equipnfighter = 1 And shiprestrict(a).equipntorpedo > 0 Then
                                nfstate = 1.2
                            ElseIf shiprestrict(a).equipnfighter > 0 And shiprestrict(a).equipnothers > 0 Then
                                nfstate = 1.18
                            End If
                        End If
                    End If
                    Exit For
                End If
            Next
        End Get
    End Property

    ReadOnly Property CIstate(ByVal x As carry_class) As Single
        Get
            CIstate = 0
            Dim code As String
            code = Mid(x.uniquecode, 1, 4)
            For a = 0 To 19
                If shiprestrict(a).shipuniquecode = code Then
                    If shiprestrict(a).needcfighter = 1 Then
                        If shiprestrict(a).equipcfighter > 0 And shiprestrict(a).equipcboomer > 0 And shiprestrict(a).equipctorpedo > 0 Then
                            CIstate = 1.25
                        End If
                    ElseIf shiprestrict(a).needcboomer >= 1 Then
                        If shiprestrict(a).equipcfighter > 0 And shiprestrict(a).equipcboomer > 0 And shiprestrict(a).equipctorpedo > 0 Then
                            CIstate = 1.25
                        ElseIf shiprestrict(a).equipcboomer >= 2 And shiprestrict(a).equipctorpedo > 0 Then
                            CIstate = 1.2
                        ElseIf shiprestrict(a).equipcboomer = 1 And shiprestrict(a).equipctorpedo > 0 Then
                            CIstate = 1.15
                        End If
                    End If
                End If
            Next
        End Get
    End Property

    Public Sub setcarry(ByVal x As carry_class, Optional ByVal y As Integer = 0)
        If y = 0 Then
            If plane.getattribute(x.planeid, 3) = 8 Then
                shiprestrict(x.shipid).equiplso = 1
            ElseIf plane.getattribute(x.planeid, 3) = 1 Then
                shiprestrict(x.shipid).equipcfighter = shiprestrict(x.shipid).equipcfighter + 1
                If plane.getattribute(x.planeid, 16) = 1 Then
                    shiprestrict(x.shipid).equipnfighter = shiprestrict(x.shipid).equipnfighter + 1
                ElseIf plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers + 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 2 Or plane.getattribute(x.planeid, 3) = 3 Then
                shiprestrict(x.shipid).equipcboomer = shiprestrict(x.shipid).equipcboomer + 1
                If plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers + 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 4 Then
                shiprestrict(x.shipid).equipctorpedo = shiprestrict(x.shipid).equipctorpedo + 1
                If plane.getattribute(x.planeid, 16) = 1 Then
                    shiprestrict(x.shipid).equipntorpedo = shiprestrict(x.shipid).equipntorpedo + 1
                ElseIf plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers + 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 5 Then
                shiprestrict(x.shipid).equipcjet = shiprestrict(x.shipid).equipcjet + 1
            ElseIf plane.getattribute(x.planeid, 3) = 12 Or plane.getattribute(x.planeid, 3) = 13 Then
                shiprestrict(x.shipid).equipwboomer = shiprestrict(x.shipid).equipwboomer + 1
            End If
        ElseIf y = 1 Then
            shiprestrict(x.shipid).getcfightercarry = shiprestrict(x.shipid).getcfightercarry + 1
        End If
    End Sub

    Public Sub removecarry(ByVal x As carry_class, Optional ByVal y As Integer = 0)
        If y = 0 Then
            If plane.getattribute(x.planeid, 3) = 8 Then
                shiprestrict(x.shipid).equiplso = 0
            ElseIf plane.getattribute(x.planeid, 3) = 1 Then
                shiprestrict(x.shipid).equipcfighter = shiprestrict(x.shipid).equipcfighter - 1
                If plane.getattribute(x.planeid, 16) = 1 Then
                    shiprestrict(x.shipid).equipnfighter = shiprestrict(x.shipid).equipnfighter - 1
                ElseIf plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers - 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 2 Or plane.getattribute(x.planeid, 3) = 3 Then
                shiprestrict(x.shipid).equipcboomer = shiprestrict(x.shipid).equipcboomer - 1
                If plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers - 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 4 Then
                shiprestrict(x.shipid).equipctorpedo = shiprestrict(x.shipid).equipctorpedo - 1
                If plane.getattribute(x.planeid, 16) = 1 Then
                    shiprestrict(x.shipid).equipntorpedo = shiprestrict(x.shipid).equipntorpedo - 1
                ElseIf plane.getattribute(x.planeid, 16) = 2 Then
                    shiprestrict(x.shipid).equipnothers = shiprestrict(x.shipid).equipnothers - 1
                End If
            ElseIf plane.getattribute(x.planeid, 3) = 5 Then
                shiprestrict(x.shipid).equipcjet = shiprestrict(x.shipid).equipcjet - 1
            ElseIf plane.getattribute(x.planeid, 3) = 12 Or plane.getattribute(x.planeid, 3) = 13 Then
                shiprestrict(x.shipid).equipwboomer = shiprestrict(x.shipid).equipwboomer - 1
            End If
        ElseIf y = 1 Then
            shiprestrict(x.shipid).getcfightercarry = shiprestrict(x.shipid).getcfightercarry - 1
        End If

    End Sub

    Public Function estimate(ByVal x As carry_class) As Boolean
        estimate = False
        Dim code As String
        code = Mid(x.uniquecode, 1, 4)
        For a = 0 + startvalue To 19
            If shiprestrict(a).shipuniquecode = code And plane.getattribute(x.planeid, 3) = 8 Then
                estimate = lsopossible(a)
            ElseIf shiprestrict(a).shipuniquecode = code And shiprestrict(a).equiplso = 1 Then  '判断夜战机装配
                If plane.getattribute(x.planeid, 16) = 1 And plane.getattribute(x.planeid, 3) = 4 Then  '判断类型1的夜攻
                    If shiprestrict(a).equipntorpedo = 0 Then '没装备夜攻
                        estimate = True
                    ElseIf shiprestrict(a).equipntorpedo > 0 And allnfpossible(a) Then '装备1格以上夜攻；全夜战可能
                        estimate = True
                    End If
                ElseIf plane.getattribute(x.planeid, 16) = 1 And plane.getattribute(x.planeid, 3) = 1 Then  '判断类型1的夜战
                    If shiprestrict(a).equipntorpedo + shiprestrict(a).equipnfighter = 0 Then '没装备夜攻；没装备夜战
                        estimate = True
                    ElseIf shiprestrict(a).equipntorpedo + shiprestrict(a).equipnfighter > 0 And allnfpossible(a) Then '装备1格以上的夜攻或夜战；全夜战可能
                        estimate = True
                    End If
                ElseIf plane.getattribute(x.planeid, 16) = 2 And plane.getattribute(x.planeid, 3) = 2 Then  '判断类型2的夜爆战
                    If shiprestrict(a).equipntorpedo + shiprestrict(a).equipnfighter > 0 Then  '装备1格以上的夜攻或夜战
                        estimate = True
                    End If
                ElseIf plane.getattribute(x.planeid, 16) = 2 And plane.getattribute(x.planeid, 3) = 4 Then  '判断类型2的夜攻
                    If shiprestrict(a).equipntorpedo + shiprestrict(a).equipnfighter > 0 Then  '装备1格以上的夜攻或夜战
                        estimate = True
                    End If
                End If
                Exit For
            End If
        Next
    End Function

    Public Function allnfpossible(Optional ByVal x As Integer = 0) As Boolean
        allnfpossible = True
        '以下部分可能没用
        'Dim needlosnum As Integer
        'Dim equiplosnum As Integer
        'For a = 0 + startvalue To 19
        '    If shiprestrict(a).needlso = 1 Then
        '        needlosnum = needlosnum + 1
        '    End If
        '    If shiprestrict(a).needlso = 1 And shiprestrict(a).equiplso = 1 Then
        '        equiplosnum = equiplosnum + 1
        '    End If
        'Next
        'If equiplosnum < plane.getnfequipnum(0) And equiplosnum < needlosnum Then
        '    allnfpossible = False
        'End If
        '以上部分可能没用
        For a = 0 + startvalue To 19
            If shiprestrict(a).equiplso = 1 Then
                If shiprestrict(a).equipnfighter + shiprestrict(a).equipntorpedo < 1 Then
                    allnfpossible = False
                End If
            End If
        Next
    End Function

    Private Function lsopossible(ByVal x As Integer) As Boolean
        Dim needlsonum As Integer = 0
        Dim shipcarry(0) As Integer
        Dim shipid(0) As Integer
        Dim tem As Integer
        Dim nfequipmun As Integer = plane.getnfequipnum(0)

        lsopossible = False

        For a = 0 + startvalue To 19
            If shiprestrict(a).needlso > 0 And ship(a).active = True Then
                ReDim Preserve shipcarry(needlsonum)
                ReDim Preserve shipid(needlsonum)
                ship(a).resetcarryobtained()
                For b = 0 To ship(a).carrycount - 1 - 1
                    shipcarry(needlsonum) = shipcarry(needlsonum) + ship(a).getcarry(10).amount
                Next
                shipid(needlsonum) = a
                needlsonum = needlsonum + 1
            End If
        Next
        If needlsonum > 0 Then
            For a = 0 To needlsonum - 1
                For b = a To needlsonum - 1
                    If shipcarry(b) > shipcarry(a) Then
                        tem = shipcarry(a)
                        shipcarry(a) = shipcarry(b)
                        shipcarry(b) = tem
                        tem = shipid(a)
                        shipid(a) = shipid(b)
                        shipid(b) = tem
                    End If
                Next
            Next
        End If

        If needlsonum <> 0 Then
            For a = 0 To needlsonum - 1
                If shipid(a) = x Then
                    If a < nfequipmun Then
                        lsopossible = True
                    End If
                End If
            Next
        End If
    End Function

    Public Sub resetequipcarry()
        For a = 0 To 19
            If shiprestrict(a).id >= 133 And shiprestrict(a).id <= 136 Then
                shiprestrict(a).equiplso = 1
            Else
                shiprestrict(a).equiplso = 0
            End If


            shiprestrict(a).equipcfighter = 0
            shiprestrict(a).equipcboomer = 0
            shiprestrict(a).equipctorpedo = 0
            shiprestrict(a).equipcjet = 0
            shiprestrict(a).equipnfighter = 0
            shiprestrict(a).equipntorpedo = 0
            shiprestrict(a).equipnothers = 0

            shiprestrict(a).getcfightercarry = 0

            shiprestrict(a).equipwboomer = 0

        Next
    End Sub

End Class

Public Class UIcontrol_class
    Dim stringcombo2 As New Collection
    Dim stringlist1 As New Collection
    Dim stringlist2 As New Collection

    Dim firstfleet As New Collection
    Dim secondfleet As New Collection

    Dim errorlist As New Collection

    Dim shipid As Integer
    Dim carryid As Integer

    Private Sub list1refresh()
        stringlist1.Clear()
        Dim situationstring As String
        situationstring = "<一队制空-" & ffcaavalue + ffwaavalue & "><联合制空-" & ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue & "><警告-" & errorlist.Count & "项>"
        stringlist1.Add(situationstring)
        If firstfleet.Count <> 0 Then
            For a = 1 To firstfleet.Count
                If firstfleet(a) > 999 Then
                    shipid = firstfleet(a) / 1000 - 1
                    stringlist1.Add("<F1>" & baseship.getattribute(ship(shipid).uniquecode, 1) & "[" & restrictcontrol.getrestrictattribute(shipid, 1) & "]") '& ship(shipid).uniquecode)
                ElseIf firstfleet(a) <= 999 And firstfleet(a) >= 0 Then
                    shipid = Int(firstfleet(a) / 10)
                    carryid = firstfleet(a) - (Int(firstfleet(a) / 10) * 10)
                    If ship(shipid).getcarry(carryid).planeid = 0 Then
                        stringlist1.Add("[" & Format(ship(shipid).getcarry(carryid).amount, "00") & "]" & "[无搭载]")
                    Else
                        stringlist1.Add("[" & Format(ship(shipid).getcarry(carryid).amount, "00") & "][" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, 1) & "]")
                    End If
                End If
            Next
        End If
        If secondfleet.Count <> 0 Then
            For a = 1 To secondfleet.Count
                If secondfleet(a) > 999 Then
                    shipid = secondfleet(a) / 1000 - 1
                    stringlist1.Add("<F2>" & baseship.getattribute(ship(shipid).uniquecode, 1) & "[" & restrictcontrol.getrestrictattribute(shipid, 1) & "]") ' & ship(shipid).uniquecode)
                ElseIf secondfleet(a) <= 999 And secondfleet(a) >= 0 Then
                    shipid = Int(secondfleet(a) / 10)
                    carryid = secondfleet(a) - (Int(secondfleet(a) / 10) * 10)
                    If ship(shipid).getcarry(carryid).planeid = 0 Then
                        stringlist1.Add("[" & Format(ship(shipid).getcarry(carryid).amount, "00") & "]" & "[无搭载]")
                    Else
                        stringlist1.Add("[" & Format(ship(shipid).getcarry(carryid).amount, "00") & "][" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, 1) & "]")
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub combo2refresh(ByVal x As Integer)
        If x > 0 And x <= firstfleet.Count Then
            stringcombo2.Clear()
            If firstfleet(x) > 999 Then
                shipid = firstfleet(x) / 1000 - 1
                Dim num As Integer = 0
                Do While restrictcontrol.getbaserestrictname(Int(ship(shipid).restrict / 10) * 10 + num) <> ""
                    stringcombo2.Add(restrictcontrol.getbaserestrictname(Int(ship(shipid).restrict / 10) * 10 + num))
                    num = num + 1
                Loop
            End If
        ElseIf x > firstfleet.Count Then
            stringcombo2.Clear()
            If secondfleet(x - firstfleet.Count) > 999 Then
                shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                Dim num As Integer = 0
                Do While restrictcontrol.getbaserestrictname(Int(ship(shipid).restrict / 10) * 10 + num) <> ""
                    stringcombo2.Add(restrictcontrol.getbaserestrictname(Int(ship(shipid).restrict / 10) * 10 + num))
                    num = num + 1
                Loop
            End If
        End If
    End Sub

    Private Sub list2refresh(Optional ByVal x As Integer = 0)
        Dim errorcode As Integer
        Dim errorship As Integer
        Dim situationstring As String
        Dim shipcount As Integer
        Dim planecount As Integer
        Dim firstminaviationfire As Integer
        Dim firstmaxaviationfire As Integer
        Dim combominaviationfire As Integer
        Dim combomaxaviationfire As Integer
        stringlist2.Clear()
        If x <= 0 Then
            For a = 0 To 19
                If ship(a).active Then
                    shipcount = shipcount + 1
                    For b = 0 To ship(a).carrycount - 1
                        If ship(a).getcarry(b).planeid <> 0 Then
                            If plane.getattribute(ship(a).getcarry(b).planeid, 3) <> 8 And plane.getattribute(ship(a).getcarry(b).planeid, 3) <> 21 Then
                                planecount = planecount + ship(a).getcarry(b).amount
                            End If
                        End If
                    Next
                    combominaviationfire = combominaviationfire + ship(a).minaviationfire
                    combomaxaviationfire = combomaxaviationfire + ship(a).maxaviationfire
                    If a <= 9 Then
                        firstminaviationfire = firstminaviationfire + ship(a).minaviationfire
                        firstmaxaviationfire = firstmaxaviationfire + ship(a).maxaviationfire
                    End If
                End If
            Next
            situationstring = "<舰娘数-" & shipcount & "><总舰载机数-" & planecount & ">"
            stringlist2.Add(situationstring)
            stringlist2.Add("<第一舰队制空/目标制空差-" & ffcaavalue + ffwaavalue & "/" & ffcaavalue + ffwaavalue - ordinaryfleetaavalue & ">")
            stringlist2.Add("<联合舰队制空/目标制空差-" & ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue & "/" & ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue - combinedfleetaavalue & ">")
            situationstring = "<航空开幕-第一舰队(" & firstminaviationfire & "-" & firstmaxaviationfire & ")-联合舰队(" & combominaviationfire & "-" & combomaxaviationfire & ")>"
            stringlist2.Add(situationstring)
            stringlist2.Add("====<警告-" & errorlist.Count & "项>=======================================")
            If errorlist.Count <> 0 Then
                For a = 1 To errorlist.Count
                    If errorlist(a) = 1 Then
                        situationstring = "<警告：舰战数量不足>"
                    ElseIf errorlist(a) = 3 Then
                        situationstring = "<警告：舰爆数量不足>"
                    ElseIf errorlist(a) = 4 Then
                        situationstring = "<警告：舰攻数量不足>"
                    ElseIf errorlist(a) = 6 Then
                        situationstring = "<警告：未能成功装载彩云以回避T不利>"
                    ElseIf errorlist(a) = 12 Then
                        situationstring = "<警告：水爆数量不足>"
                    ElseIf errorlist(a) = 21 Then
                        situationstring = "<警告：未能成功装载司令部设施>"
                    ElseIf errorlist(a) = 100 Then
                        situationstring = "<警告：未编入舰娘或无可用舰载机>"
                    ElseIf errorlist(a) = 101 Then
                        situationstring = "<警告：第一舰队无法满足制空需求>"
                    ElseIf errorlist(a) = 102 Then
                        situationstring = "<警告：联合舰队无法满足制空需求>"
                    Else
                        If errorlist(a) > 10000 Then
                            errorcode = Int(errorlist(a) / 100)
                            errorship = errorlist(a) - Int(errorlist(a) / 100) * 100
                            situationstring = "<警告：" & baseship.getattribute(ship(errorship).uniquecode, 1) & "-----"
                            If errorcode = 101 Then
                                situationstring = situationstring & "无法炮击>"
                            ElseIf errorcode = 102 Then
                                situationstring = situationstring & "未达成CI发动条件>"
                            ElseIf errorcode = 103 Then
                                situationstring = situationstring & "未达成夜战发动条件>"
                            ElseIf errorcode = 111 Then
                                situationstring = situationstring & "未装配水爆>"
                            End If
                        End If
                    End If
                    stringlist2.Add(situationstring)
                Next
            Else
                stringlist2.Add("<无警告>")
            End If
        ElseIf x > 0 Then
            If x <= firstfleet.Count Then
                If firstfleet(x) > 999 Then
                    shipid = firstfleet(x) / 1000 - 1
                    carryid = -1
                ElseIf firstfleet(x) < 999 Then
                    shipid = Int(firstfleet(x) / 10)
                    carryid = firstfleet(x) - Int(firstfleet(x) / 10) * 10
                End If
            ElseIf x > firstfleet.Count Then
                If secondfleet(x - firstfleet.Count) > 999 Then
                    shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                    carryid = -1
                ElseIf secondfleet(x - firstfleet.Count) < 999 Then
                    shipid = Int(secondfleet(x - firstfleet.Count) / 10)
                    carryid = secondfleet(x - firstfleet.Count) - Int(secondfleet(x - firstfleet.Count) / 10) * 10
                End If
            End If
            If carryid = -1 Then
                situationstring = "<第一舰队>" & baseship.getattribute(ship(shipid).uniquecode, 1) & "[" & restrictcontrol.getbaserestrictname(ship(shipid).restrict) & "]"
                For a = 101 To 103
                    If errorlist.Contains(Trim(Str(a * 100 + shipid))) Then
                        If a = 101 Then
                            situationstring = situationstring & "<无法炮击>"
                        ElseIf a = 102 Then
                            situationstring = situationstring & "<未达成CI发动条件>"
                        ElseIf a = 103 Then
                            situationstring = situationstring & "<未达成夜战发动条件>"
                        ElseIf a = 111 Then
                            situationstring = situationstring & "<未装配水爆>"
                        End If
                    End If
                Next
                stringlist2.Add(situationstring)
                situationstring = "<制空值-" & ship(shipid).AAvalue & "><航空开幕-(" & Int(ship(shipid).minaviationfire) & "-" & Int(ship(shipid).maxaviationfire) & ")>"
                stringlist2.Add(situationstring)
                If ship(shipid).shellingfire <> -1 Then
                    If ship(shipid).shellingfire > 0 Then
                        situationstring = "<航空炮击-" & ship(shipid).shellingfire & ">"
                    Else
                        situationstring = "<航空炮击-无法炮击"
                    End If
                Else
                    situationstring = "<航空炮击-Null>"
                End If
                If ship(shipid).CIshellingfire <> -1 Then
                    If restrictcontrol.CIstate(ship(shipid).getcarry(0)) <> 0 Then
                        situationstring = situationstring & "<航空CI-" & ship(shipid).CIshellingfire & "(" & restrictcontrol.CIstate(ship(shipid).getcarry(0)) & ")>"
                    ElseIf restrictcontrol.CIstate(ship(shipid).getcarry(0)) = 0 Then
                        situationstring = situationstring & "<航空CI-无法CI>"
                    End If
                Else
                    situationstring = situationstring & "<航空CI-Null>"
                End If
                stringlist2.Add(situationstring)
                If ship(shipid).nightfightfire <> -1 Then
                    If restrictcontrol.nfstate(ship(shipid).getcarry(0)) >= 1 Then
                        situationstring = "<夜战炮击-" & ship(shipid).nightfightfire & ">"
                    ElseIf restrictcontrol.nfstate(ship(shipid).getcarry(0)) < 1 Then
                        situationstring = "<夜战炮击-无法夜战>"
                    End If
                Else
                    situationstring = "<夜战炮击-Null>"
                End If
                If ship(shipid).CInightfightfire <> -1 Then
                    If restrictcontrol.nfstate(ship(shipid).getcarry(0)) > 1 Then
                        situationstring = situationstring & "<夜战CI-" & ship(shipid).CInightfightfire & "(" & restrictcontrol.nfstate(ship(shipid).getcarry(0)) & ")>"
                    ElseIf restrictcontrol.nfstate(ship(shipid).getcarry(0)) <= 1 Then
                        situationstring = situationstring & "<夜战CI-无法CI>"
                    End If
                Else
                    situationstring = situationstring & "<夜战CI-Null>"
                End If
                stringlist2.Add(situationstring)
                Dim carrycount As Integer = 0
                planecount = 0
                For a = 0 To ship(shipid).carrycount - 1
                    If ship(shipid).getcarry(a).planeid <> 0 Then
                        If plane.getattribute(ship(shipid).getcarry(a).planeid, 3) <> 8 And plane.getattribute(ship(shipid).getcarry(a).planeid, 3) <> 21 Then
                            carrycount = carrycount + 1
                            planecount = planecount + ship(shipid).getcarry(a).amount
                        End If
                    End If
                Next
                situationstring = "====<已搭载舰载机-" & carrycount & "格-" & planecount & "架>=========================="
                stringlist2.Add(situationstring)
                For a = 0 To ship(shipid).carrycount - 1
                    situationstring = "[" & Format(ship(shipid).getcarry(a).amount, "00") & "]["
                    If ship(shipid).getcarry(a).planeid <> 0 Then
                        situationstring = situationstring & plane.getattribute(ship(shipid).getcarry(a).planeid, 1)
                    Else
                        situationstring = situationstring & "无搭载"
                    End If
                    situationstring = situationstring & "]<" & ship(shipid).getcarry(a).AAvalue & "><(" & Int(ship(shipid).getcarry(a).minaviationfire) & "-" & Int(ship(shipid).getcarry(a).maxaviationfire) & ")>"
                    stringlist2.Add(situationstring)
                Next
            ElseIf carryid <> -1 Then
                situationstring = "<" & baseship.getattribute(ship(shipid).uniquecode, 1) & "-" & carryid + 1 & ">[" & ship(shipid).getcarry(carryid).amount & "]"
                If ship(shipid).getcarry(carryid).planeid <> 0 Then
                    situationstring = situationstring & plane.getattribute(ship(shipid).getcarry(carryid).planeid, 1)
                Else
                    situationstring = situationstring & "无搭载"
                End If
                situationstring = situationstring & "<" & ship(shipid).getcarry(carryid).AAvalue & "><(" & Int(ship(shipid).getcarry(carryid).minaviationfire) & "-" & Int(ship(shipid).getcarry(carryid).maxaviationfire) & ")>"
                stringlist2.Add(situationstring)
                situationstring = "<改修-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, 17) & ">"
                stringlist2.Add(situationstring)
                For a = 4 To 12
                    If plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) <> 0 Then
                        If a = 4 Then
                            situationstring = "<火力-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 5 Then
                            situationstring = "<雷装-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 6 Then
                            situationstring = "<对空-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 7 Then
                            situationstring = "<对潜-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 8 Then
                            situationstring = "<爆装-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 9 Then
                            situationstring = "<命中-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 10 Then
                            situationstring = "<装甲-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 11 Then
                            situationstring = "<回避-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 12 Then
                            situationstring = "<索敌-" & plane.getattribute(ship(shipid).getcarry(carryid).planeid, a) & ">"
                        End If
                        stringlist2.Add(situationstring)
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub setship(ByVal x As Integer)
        If x < 10 Then
            shipid = (x + 1) * 1000
            firstfleet.Add(shipid)
            shipid = shipid / 1000 - 1
            For a = 0 To ship(shipid).carrycount - 1
                firstfleet.Add(shipid * 10 + a)
            Next
        ElseIf x > 9 Then
            shipid = (x + 1) * 1000
            secondfleet.Add(shipid)
            shipid = shipid / 1000 - 1
            For a = 0 To ship(shipid).carrycount - 1
                secondfleet.Add(shipid * 10 + a)
            Next
        End If
    End Sub

    Public Sub setshipattribute(ByVal x As Integer, ByVal y As Integer)
        If x > 0 And x <= firstfleet.Count Then
            If firstfleet(x) > 999 Then
                shipid = firstfleet(x) / 1000 - 1
                ship(shipid).restrict = Int(ship(shipid).restrict / 10) * 10 + y
            End If
        ElseIf x > firstfleet.Count Then
            If secondfleet(x - firstfleet.Count) > 999 Then
                shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                ship(shipid).restrict = Int(ship(shipid).restrict / 10) * 10 + y
            End If
        End If
        Call list1refresh()
    End Sub

    Public Sub removeship(ByVal x As Integer)
        Dim num As Integer = 1
        If x > 0 And x <= firstfleet.Count Then
            If firstfleet(x) > 999 Then
                shipid = firstfleet(x) / 1000 - 1
                Do While num <= firstfleet.Count
                    If firstfleet(num) = (shipid + 1) * 1000 Then
                        firstfleet.Remove(num)
                    ElseIf Int(firstfleet(num) / 10) = shipid Then
                        firstfleet.Remove(num)
                    Else
                        num = num + 1
                    End If
                Loop
                ship(shipid).clear()
            End If
        ElseIf x > firstfleet.Count Then
            If secondfleet(x - firstfleet.Count) > 999 Then
                shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                Do While num <= secondfleet.Count
                    If secondfleet(num) = (shipid + 1) * 1000 Then
                        secondfleet.Remove(num)
                    ElseIf Int(secondfleet(num) / 10) = shipid Then
                        secondfleet.Remove(num)
                    Else
                        num = num + 1
                    End If
                Loop
                ship(shipid).clear()
            End If
        End If
    End Sub

    Public Function showstringlist1(ByVal x As Integer) As String
        If x = 1 Then Call list1refresh()
        Try
            showstringlist1 = stringlist1(x)
        Catch
            showstringlist1 = ""
        End Try
    End Function

    Public Function showstringcombo2(ByVal x As Integer, ByVal y As Integer) As String
        If x = 1 Then Call combo2refresh(y)
        Try
            showstringcombo2 = stringcombo2(x)
        Catch
            showstringcombo2 = ""
        End Try
    End Function

    Public Function getshipid(ByVal x As Integer) As Integer
        getshipid = -1
        If x > 0 And x <= firstfleet.Count Then
            If firstfleet(x) > 999 Then
                getshipid = firstfleet(x) / 1000 - 1
            End If
        ElseIf x > firstfleet.Count Then
            If secondfleet(x - firstfleet.Count) > 999 Then
                getshipid = secondfleet(x - firstfleet.Count) / 1000 - 1
            End If
        End If
    End Function

    Public Function showstringlist2(ByVal x As Integer, ByVal y As Integer) As String
        If x = 1 Then Call list2refresh(y)
        Try
            showstringlist2 = stringlist2(x)
        Catch
            showstringlist2 = ""
        End Try
    End Function

    Public Sub clearship()
        For a = 0 To 19
            ship(a).clear()
        Next
        firstfleet.Clear()
        secondfleet.Clear()
    End Sub

    Public Sub adderror(ByVal x As Integer)
        If errorlist.Count <> 0 Then
            If errorlist.Contains(x) = False Then
                errorlist.Add(x, x)
            End If
        Else
            errorlist.Add(x, x)
        End If
    End Sub

    Public Sub clearerror()
        errorlist.Clear()
    End Sub

    Public Sub exportconfigurecode()
        Dim configurationcode As String
        Dim ffcount As Integer = 0
        Dim sfcount As Integer = 0

        Dim itemcode As String = ""
        Dim itemcodecolle As New Collection

        configurationcode = "{" & Chr(34) & "version" & Chr(34) & ":4,"

        If firstfleet.Count <> 0 Then
            configurationcode = configurationcode & Chr(34) & "f1" & Chr(34) & ":{"
            For a = 1 To firstfleet.Count
                If firstfleet(a) > 999 Then
                    ffcount = ffcount + 1
                    itemcodecolle.Clear()
                    If ffcount <> 1 Then
                        configurationcode = configurationcode & ","
                    End If
                    shipid = firstfleet(a) / 1000 - 1
                    configurationcode = configurationcode & Chr(34) & "s" & ffcount & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & baseship.getattribute(ship(shipid).uniquecode, 2) & "," & Chr(34) & "luck" & Chr(34) & ":-1," & Chr(34) & "items" & Chr(34) & ":{"
                    For b = 0 To ship(shipid).carrycount - 1
                        If ship(shipid).getcarry(b).planeid <> 0 Then
                            itemcode = Chr(34) & "i" & b + 1 & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & plane.getattribute(ship(shipid).getcarry(b).planeid, 2) & "," & Chr(34) & "rf" & Chr(34) & ":" & plane.getattribute(ship(shipid).getcarry(b).planeid, 17) & "," & Chr(34) & "mas" & Chr(34) & ":"
                            If plane.getattribute(ship(shipid).getcarry(b).planeid, 3) = 8 Then
                                itemcode = itemcode & "0}"
                            Else
                                itemcode = itemcode & "7}"
                            End If
                            itemcodecolle.Add(itemcode)
                        End If
                    Next
                    If itemcodecolle.Count <> 0 Then
                        For b = 1 To itemcodecolle.Count
                            If b <> 1 Then
                                configurationcode = configurationcode & ","
                            End If
                            configurationcode = configurationcode & itemcodecolle(b)
                        Next
                    End If
                    configurationcode = configurationcode & "}}"
                End If
            Next
            configurationcode = configurationcode & "}"
        Else
            configurationcode = configurationcode & Chr(34) & "f1" & Chr(34) & ":{}"
        End If
        If secondfleet.Count <> 0 Then
            configurationcode = configurationcode & "," & Chr(34) & "f2" & Chr(34) & ":{"
            For a = 1 To secondfleet.Count
                If secondfleet(a) > 999 Then
                    sfcount = sfcount + 1
                    itemcodecolle.Clear()
                    If sfcount <> 1 Then
                        configurationcode = configurationcode & ","
                    End If
                    shipid = secondfleet(a) / 1000 - 1
                    configurationcode = configurationcode & Chr(34) & "s" & sfcount & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & baseship.getattribute(ship(shipid).uniquecode, 2) & "," & Chr(34) & "luck" & Chr(34) & ":-1," & Chr(34) & "items" & Chr(34) & ":{"
                    For b = 0 To ship(shipid).carrycount - 1
                        If ship(shipid).getcarry(b).planeid <> 0 Then
                            itemcode = Chr(34) & "i" & b + 1 & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & plane.getattribute(ship(shipid).getcarry(b).planeid, 2) & "," & Chr(34) & "rf" & Chr(34) & ":" & plane.getattribute(ship(shipid).getcarry(b).planeid, 22) & "," & Chr(34) & "mas" & Chr(34) & ":"
                            If plane.getattribute(ship(shipid).getcarry(b).planeid, 3) = 8 Then
                                itemcode = itemcode & "0}"
                            Else
                                itemcode = itemcode & "7}"
                            End If
                            itemcodecolle.Add(itemcode)
                        End If
                    Next
                    If itemcodecolle.Count <> 0 Then
                        For b = 1 To itemcodecolle.Count
                            If b <> 1 Then
                                configurationcode = configurationcode & ","
                            End If
                            configurationcode = configurationcode & itemcodecolle(b)
                        Next
                    End If
                    configurationcode = configurationcode & "}}"
                End If
            Next
            configurationcode = configurationcode & "}"
        Else
            configurationcode = configurationcode & "," & Chr(34) & "f2" & Chr(34) & ":{}"
        End If
        configurationcode = configurationcode & "," & Chr(34) & "f3" & Chr(34) & ":{}," & Chr(34) & "f4" & Chr(34) & ":{}," & Chr(34) & "fField" & Chr(34) & ":{" & Chr(34) & "f1" & Chr(34) & ":{}," & Chr(34) & "f2" & Chr(34) & ":{}," & Chr(34) & "f3" & Chr(34) & ":{}}}"
        Try
            Clipboard.Clear()
            Clipboard.SetText(configurationcode)
            MsgBox("配置代码已复制到剪贴板")
        Catch
            Clipboard.Clear()
            Clipboard.SetText(configurationcode)
            MsgBox("配置代码已复制到剪贴板")
        End Try
    End Sub
End Class

Public Class filecontrol_class
    Sub New()
        saabpd.Load(Application.StartupPath + "\data\SAAbaseplanedata.xml")
        For Each bpnode As Xml.XmlNode In saabpd.DocumentElement.ChildNodes
            If bpnode.Name = "plane" Then
                bplane.name = bpnode.Attributes("name").Value
                bplane.id = bpnode.Attributes("id").Value
                bplane.classification = bpnode.Attributes("classification").Value
                bplane.fire = bpnode.Attributes("fire").Value
                bplane.torpedo = bpnode.Attributes("torpedo").Value
                bplane.antiaircraft = bpnode.Attributes("antiaircraft").Value
                bplane.antisubmarine = bpnode.Attributes("antisubmarine").Value
                bplane.bombing = bpnode.Attributes("bombing").Value
                bplane.hit = bpnode.Attributes("hit").Value
                bplane.armor = bpnode.Attributes("armor").Value
                bplane.avoid = bpnode.Attributes("avoid").Value
                bplane.spotting = bpnode.Attributes("spotting").Value
                bplane.airrange = bpnode.Attributes("airrange").Value
                bplane.antibombing = bpnode.Attributes("antibombing").Value
                bplane.intercept = bpnode.Attributes("intercept").Value
                bplane.nightfighting = bpnode.Attributes("nightfighting").Value

                bpcolle.Add(bplane, bplane.id)
            End If
        Next
    End Sub

    Public Sub checkplaneupdata()
        Dim bpdrootnode As Xml.XmlElement = saabpd.DocumentElement
        Dim updrootnode As Xml.XmlElement = saaupd.DocumentElement
        If updrootnode.Attributes("version").Value <> bpdrootnode.Attributes("version").Value Then
            updrootnode.SetAttribute("version", bpdrootnode.Attributes("version").Value)
            Dim antiaircraft As Double
            For Each upnode As Xml.XmlElement In saaupd.DocumentElement.ChildNodes
                If bpcolle.Contains(upnode.Attributes("id").Value) Then
                    bplane = bpcolle.Item(upnode.Attributes("id").Value)
                    upnode.SetAttribute("classification", bplane.classification)
                    upnode.SetAttribute("fire", bplane.fire)
                    upnode.SetAttribute("torpedo", bplane.torpedo)
                    If upnode.Attributes("classification").Value = 1 Then
                        antiaircraft = bplane.antiaircraft + upnode.Attributes("improve").Value * 0.2
                    ElseIf upnode.Attributes("classification").Value = 2 Then
                        antiaircraft = bplane.antiaircraft + upnode.Attributes("improve").Value * 0.25
                    Else
                        antiaircraft = bplane.antiaircraft
                    End If
                    upnode.SetAttribute("antiaircraft", antiaircraft)
                    upnode.SetAttribute("antisubmarine", bplane.antisubmarine)
                    upnode.SetAttribute("bombing", bplane.bombing)
                    upnode.SetAttribute("hit", bplane.hit)
                    upnode.SetAttribute("armor", bplane.armor)
                    upnode.SetAttribute("avoid", bplane.avoid)
                    upnode.SetAttribute("spotting", bplane.spotting)
                    upnode.SetAttribute("airrange", bplane.airrange)
                    upnode.SetAttribute("antibombing", bplane.antibombing)
                    upnode.SetAttribute("intercept", bplane.intercept)
                    upnode.SetAttribute("nightfighting", bplane.nightfighting)
                End If
            Next
            saaupd.Save(Application.StartupPath + "\data\SAAuserplanedata.xml")
        End If
    End Sub
End Class