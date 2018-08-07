Public Class planeunit_class
    Inherits equipmentunit_class

    Sub New(ByVal newuniquecode As String)
        MyBase.New(newuniquecode)
    End Sub

    ReadOnly Property classification As Double
        Get
            classification = cvequipmentgroup.getattribute(MyBase.baseid, 2)
        End Get
    End Property

    ReadOnly Property antibombing As Double
        Get
            antibombing = cvequipmentgroup.getattribute(MyBase.baseid, 3)
        End Get
    End Property

    ReadOnly Property intercept As Double
        Get
            intercept = cvequipmentgroup.getattribute(MyBase.baseid, 4)
        End Get
    End Property

    ReadOnly Property nightfighting As Double
        Get
            nightfighting = cvequipmentgroup.getattribute(MyBase.baseid, 5)
        End Get
    End Property

    ReadOnly Property name As String
        Get
            name = usercvequipmentgroup.getattribute(MyBase.extendid, 1)
        End Get
    End Property

    ReadOnly Property improve As Double
        Get
            improve = usercvequipmentgroup.getattribute(MyBase.extendid, 2)
        End Get
    End Property

    Overrides ReadOnly Property antiaircraft As Double
        Get
            antiaircraft = usercvequipmentgroup.getattribute(MyBase.extendid, 5)
        End Get
    End Property
End Class

Public Class plane_class
    Dim uselist As New Collection
    Dim temuselist As New Collection

    Dim planeunit() As planeunit_class
    Dim planelist As New Collection

    Sub New()
        'Call loadplaneunit(0)
    End Sub

    Public Sub loadplaneunit(ByVal x As Integer)
        ReDim planeunit(-1)
        planelist.Clear()
        Dim unitcount As Integer = 0
        Dim sameplanecount As Integer
        Dim newuniquecode As String
        If usercvequipmentgroup IsNot Nothing Then
            If usercvequipmentgroup.length > 0 Then
                For a = 0 To usercvequipmentgroup.length - 1
                    If usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), x + 3) <> 0 Then
                        unitcount = unitcount + usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), x + 3)
                        ReDim Preserve planeunit(unitcount - 1)
                        sameplanecount = 1
                        For b = unitcount - usercvequipmentgroup.getattribute(usercvequipmentgroup.getid(a), x + 3) To unitcount - 1
                            newuniquecode = Trim(Str(Val(usercvequipmentgroup.getid(a)) * 100 + sameplanecount))
                            planeunit(b) = New planeunit_class(newuniquecode)
                            planelist.Add(b, newuniquecode)
                            sameplanecount = sameplanecount + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Function getplane(ByVal x As Integer, Optional ByVal y As Double = 0, Optional ByVal z As Double = 0, Optional ByVal zz As Double = 0) As Integer
        getplane = 0
        Dim fattribute As Double = y
        Dim sattribute As Double = 0

        If x = 0 Then fattribute = 0


        Dim unuselist As New Collection
        Dim fattributegroup As New Collection
        Dim sattributegroup As New Collection

        If planeunit.Length > 0 Then
            For a = 0 To planeunit.Length - 1
                If used(planeunit(a).uniquecode) = False Then
                    If x = 0 And planeunit(a).nightfighting >= 1 Then '取得夜战机
                        If y = 0 And planeunit(a).classification = 1 Then 'y0取得夜战   
                            unuselist.Add(planeunit(a).uniquecode)
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If z = 0 Then 'z0优先对潜 
                                sattributegroup.Add(planeunit(a).antisubmarine)
                            End If
                        ElseIf y = 1 And planeunit(a).classification = 2 Then 'y1取得夜爆战
                            unuselist.Add(planeunit(a).uniquecode)
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If z = 0 Then 'z0优先爆装  
                                sattributegroup.Add(planeunit(a).bombing)
                            End If
                        ElseIf y = 2 And planeunit(a).classification = 3 Then 'y1取得夜爆
                            unuselist.Add(planeunit(a).uniquecode)
                            fattributegroup.Add(planeunit(a).bombing)
                            If z = 0 Then 'z0优先对空  
                                sattributegroup.Add(planeunit(a).antiaircraft)
                            End If
                        ElseIf y = 3 And planeunit(a).classification = 4 Then 'y1取得夜攻
                            unuselist.Add(planeunit(a).uniquecode)
                            fattributegroup.Add(planeunit(a).torpedo)
                            If z = 0 Then 'z0优先对空  
                                sattributegroup.Add(planeunit(a).antiaircraft)
                            End If
                        End If

                    ElseIf x = 1 And planeunit(a).classification = 1 Then  '取得舰战
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then 'zz0副优先闪避
                                sattributegroup.Add(planeunit(a).avoid)
                            End If
                        ElseIf z = 1 Then 'z1主优先闪避
                            fattributegroup.Add(planeunit(a).avoid)
                            If zz = 0 Then 'zz0副优先对空
                                sattributegroup.Add(planeunit(a).antiaircraft)
                            End If
                        End If
                    ElseIf x = 2 And planeunit(a).classification = 2 Then  '取得爆战
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then 'zz0副优先爆装
                                sattributegroup.Add(planeunit(a).bombing)
                            End If
                        End If
                    ElseIf x = 3 And planeunit(a).classification = 3 Then  '取得舰爆
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先爆装
                            fattributegroup.Add(planeunit(a).bombing)
                            If zz = 0 Then  'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        End If
                    ElseIf x = 4 And planeunit(a).classification = 4 Then  '取得舰攻
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先雷装
                            fattributegroup.Add(planeunit(a).torpedo)
                            If zz = 0 Then 'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        End If
                    ElseIf x = 5 And planeunit(a).classification = 5 Then  '取得喷气机
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then 'zz0主优先爆装
                                sattributegroup.Add(planeunit(a).bombing)
                            End If
                        End If
                    ElseIf x = 6 And planeunit(a).classification = 6 Then  '取得彩云
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先索敌
                            fattributegroup.Add(planeunit(a).spotting)
                            If zz = 0 Then 'zz0副优先火力
                                sattributegroup.Add(planeunit(a).fire)
                            End If
                        End If
                    ElseIf x = 7 And planeunit(a).classification = 7 Then  '取得二式舰侦
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先索敌
                            fattributegroup.Add(planeunit(a).spotting)
                            If zz = 0 Then 'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        End If
                    ElseIf x = 8 And planeunit(a).classification = 8 Then  '取得夜战要员
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先火力
                            fattributegroup.Add(planeunit(a).fire)
                            If zz = 0 Then 'zz0副优先回避
                                sattributegroup.Add(planeunit(a).avoid)
                            End If
                        End If
                    ElseIf x = 11 And planeunit(a).classification = 11 Then  '取得水战
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then 'zz0副优先回避
                                sattributegroup.Add(planeunit(a).avoid)
                            End If
                        End If
                    ElseIf x = 12 And planeunit(a).classification = 12 Then  '取得水爆
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then  'zz0副优先爆装
                                sattributegroup.Add(planeunit(a).bombing)
                            End If
                        ElseIf z = 1 Then 'z1主优先爆装
                            fattributegroup.Add(planeunit(a).bombing)
                            If zz = 0 Then 'zz0副优先对潜
                                sattributegroup.Add(planeunit(a).antisubmarine)
                            End If
                        End If
                    ElseIf x = 13 And planeunit(a).classification = 13 Then  '取得飞行艇
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先索敌
                            fattributegroup.Add(planeunit(a).spotting)
                            If zz = 0 Then 'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        End If
                    ElseIf x = 21 And planeunit(a).classification = 21 Then  '取得司令部
                        unuselist.Add(planeunit(a).uniquecode)
                        If z = 0 Then 'z0主优先对空
                            fattributegroup.Add(planeunit(a).antiaircraft)
                            If zz = 0 Then 'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        ElseIf z = 1 Then 'z1主优先装甲
                            fattributegroup.Add(planeunit(a).armor)
                            If zz = 0 Then 'zz0副优先命中
                                sattributegroup.Add(planeunit(a).hit)
                            End If
                        End If
                    End If
                End If
            Next
        End If


        Dim planeid As Integer
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
                        planeid = unuselist(a)
                        sattribute = sattributegroup(a)
                    End If
                End If
            Next
            getplane = planeid
        End If

    End Function

    Public Function getlbplane(ByVal classification As Integer, ByVal airrange As Integer, Optional ByVal mode As Integer = 0) As Double
        getlbplane = 0

        Dim fattribute As Double = 0
        Dim sattribute As Double = 0

        Dim unuselist As New Collection
        Dim fattributegroup As New Collection
        Dim sattributegroup As New Collection

        If planeunit.Length > 0 Then
            For a = 0 To planeunit.Length - 1
                If used(planeunit(a).uniquecode) = False Then
                    If planeunit(a).classification = classification And planeunit(a).airrange >= airrange Then
                        If classification = 31 Then   '取得陆攻
                            If mode = 0 Then   '雷装优先，对空次要
                                fattributegroup.Add(planeunit(a).torpedo)
                                sattributegroup.Add(planeunit(a).antiaircraft)
                            ElseIf mode = 1 Then   '对空优先，雷装次要
                                fattributegroup.Add(planeunit(a).antiaircraft)
                                sattributegroup.Add(planeunit(a).torpedo)
                            End If
                        ElseIf classification = 32 Or classification = 1 Then   '取得陆战局战舰战
                            If mode = 0 Then     '优先出击制空
                                fattributegroup.Add(planeunit(a).intercept * 1.5 + planeunit(a).antiaircraft)
                                sattributegroup.Add(50 - planeunit(a).airrange)
                            ElseIf mode = 1 Then     '优先防空制空
                                fattributegroup.Add(planeunit(a).intercept + planeunit(a).antiaircraft + planeunit(a).antibombing * 2)
                                sattributegroup.Add(50 - planeunit(a).airrange)
                            End If
                        ElseIf classification = 13 Then   '取得大艇
                            If mode = 0 Then   '优先航程
                                fattributegroup.Add(planeunit(a).airrange)
                                sattributegroup.Add(planeunit(a).spotting)
                            End If
                        End If
                    End If

                End If
            Next
        End If

        Dim planeid As Integer
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
                        planeid = unuselist(a)
                        sattribute = sattributegroup(a)
                    End If
                End If
            Next
            getlbplane = planeid
        End If
    End Function

    Public Function getattribute(ByVal x As Integer, ByVal y As Integer)
        getattribute = 0
        If x <> 0 Then
            Dim uniquecode As String = Trim(x)
            If y = 1 Then
                getattribute = planeunit(planelist(uniquecode)).name
            ElseIf y = 2 Then
                getattribute = planeunit(planelist(uniquecode)).baseid
            ElseIf y = 3 Then
                getattribute = planeunit(planelist(uniquecode)).classification
            ElseIf y = 4 Then
                getattribute = planeunit(planelist(uniquecode)).fire
            ElseIf y = 5 Then
                getattribute = planeunit(planelist(uniquecode)).torpedo
            ElseIf y = 6 Then
                getattribute = planeunit(planelist(uniquecode)).antiaircraft
            ElseIf y = 7 Then
                getattribute = planeunit(planelist(uniquecode)).antisubmarine
            ElseIf y = 8 Then
                getattribute = planeunit(planelist(uniquecode)).bombing
            ElseIf y = 9 Then
                getattribute = planeunit(planelist(uniquecode)).hit
            ElseIf y = 10 Then
                getattribute = planeunit(planelist(uniquecode)).armor
            ElseIf y = 11 Then
                getattribute = planeunit(planelist(uniquecode)).avoid
            ElseIf y = 12 Then
                getattribute = planeunit(planelist(uniquecode)).spotting
            ElseIf y = 13 Then
                getattribute = planeunit(planelist(uniquecode)).airrange
            ElseIf y = 14 Then
                getattribute = planeunit(planelist(uniquecode)).antibombing
            ElseIf y = 15 Then
                getattribute = planeunit(planelist(uniquecode)).intercept
            ElseIf y = 16 Then
                getattribute = planeunit(planelist(uniquecode)).nightfighting
            ElseIf y = 17 Then
                getattribute = planeunit(planelist(uniquecode)).improve
            End If
        End If
    End Function

    ReadOnly Property unit(ByVal id As String) As planeunit_class
        Get
            unit = planeunit(planelist(id))
        End Get
    End Property

    Public Function getnfequipnum(ByVal x As Integer) As Integer
        getnfequipnum = 0
        If planeunit.Length > 0 Then
            If x = 0 Then
                For a = 0 To planeunit.Length - 1
                    If planeunit(a).classification = 8 Then
                        getnfequipnum = getnfequipnum + 1
                    End If
                Next
            ElseIf x > 0 Then
                For a = 1 To planeunit.Length - 1
                    If planeunit(a).nightfighting = x Then
                        getnfequipnum = getnfequipnum + 1
                    End If
                Next
            End If
        End If
    End Function

    Public Function calculationminairrange(ByVal airrange As Integer, ByVal planeid As String) As Integer
        calculationminairrange = airrange
        Dim result As Double
        If planeunit(planelist(planeid)).airrange >= airrange Then
            result = (-(1 - airrange * 2) - Math.Sqrt((1 - airrange * 2) ^ 2 - 4 * (airrange ^ 2 - planeunit(planelist(planeid)).airrange))) / 2
            If airrange - Math.Round(result) > 4 Then
                calculationminairrange = airrange - 3
            Else
                calculationminairrange = result
            End If
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

Public Class cvshiptype_class
    Dim cvshiptypegroup As basedatagroup_class
    Dim cvshiptypedate As basedata_class
    Dim saa_cv_std As New Xml.XmlDocument
    Dim modificationvalue As Integer = 1

    Sub New()
        filecontrol.loadfile("\base\SAA_cv_std.xml", saa_cv_std, cvshiptypegroup)
    End Sub

    Public Function getattribute(ByVal uniquecode As String, ByVal attributeindex As Integer)
        getattribute = 0
        Dim id As String = Trim(Val(uniquecode))
        If attributeindex = 1 Then   'name
            getattribute = cvshiptypegroup.getattribute(id, 1)
        ElseIf attributeindex = 2 Then   'id
            getattribute = cvshiptypegroup.getattribute(id, 0)
        ElseIf attributeindex = 3 Then   'fire
            getattribute = baseshiptypegroup.getattribute(id, 3)
        ElseIf attributeindex = 4 Then   'armor
            getattribute = baseshiptypegroup.getattribute(id, 6)
        ElseIf attributeindex = 5 Then   'antisubmarine99
            getattribute = baseshiptypegroup.getattribute(id, 7)
        ElseIf attributeindex = 6 Then   'avoid99
            getattribute = baseshiptypegroup.getattribute(id, 9)
        ElseIf attributeindex = 7 Then   'spotting99
            getattribute = baseshiptypegroup.getattribute(id, 11)
        ElseIf attributeindex = 8 Then   'carry1
            getattribute = cvshiptypegroup.getattribute(id, 2)
        ElseIf attributeindex = 9 Then   'carry2
            getattribute = cvshiptypegroup.getattribute(id, 3)
        ElseIf attributeindex = 10 Then   'carry3
            getattribute = cvshiptypegroup.getattribute(id, 4)
        ElseIf attributeindex = 11 Then   'carry4
            getattribute = cvshiptypegroup.getattribute(id, 5)
        ElseIf attributeindex = 12 Then   'carry5
            getattribute = cvshiptypegroup.getattribute(id, 6)
        ElseIf attributeindex = 13 Then   'modification
            getattribute = cvshiptypegroup.getattribute(id, 7)
        ElseIf attributeindex = 14 Then   'ncfighter
            getattribute = cvshiptypegroup.getattribute(id, 8)
        ElseIf attributeindex = 15 Then   'nctorpedo
            getattribute = cvshiptypegroup.getattribute(id, 9)
        ElseIf attributeindex = 16 Then   'ncboomer
            getattribute = cvshiptypegroup.getattribute(id, 10)
        ElseIf attributeindex = 17 Then   'ncsount
            getattribute = cvshiptypegroup.getattribute(id, 11)
        ElseIf attributeindex = 18 Then   'ncjet
            getattribute = cvshiptypegroup.getattribute(id, 12)
        ElseIf attributeindex = 19 Then   'nwfighter
            getattribute = cvshiptypegroup.getattribute(id, 13)
        ElseIf attributeindex = 20 Then   'nwboomer
            getattribute = cvshiptypegroup.getattribute(id, 14)
        ElseIf attributeindex = 21 Then   'wfpriority
            getattribute = cvshiptypegroup.getattribute(id, 15)
        ElseIf attributeindex = 22 Then   'restrict
            getattribute = cvshiptypegroup.getattribute(id, 16)
        End If
    End Function

    Public Function getid(ByVal index As Integer) As Integer
        getid = "0"
        Dim count As Integer = 0
        For a = 0 To cvshiptypegroup.length - 1
            If cvshiptypegroup.getattribute(cvshiptypegroup.getid(a), 7) >= modificationvalue Then
                count = count + 1
                If count = index Then
                    getid = cvshiptypegroup.getattribute(cvshiptypegroup.getid(a), 0)
                    Exit For
                End If
            End If
        Next
    End Function

    WriteOnly Property modification As Integer
        Set(value As Integer)
            modificationvalue = value
        End Set
    End Property
End Class

Public Class cvship_class
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
                    Dim fire As Integer = cvshiptype.getattribute(code, 3)
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
                    Dim fire As Integer = cvshiptype.getattribute(code, 3)
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

    Public Function enabletype(ByVal x() As Integer, Optional ByVal y As Integer = 0) As Integer
        enabletype = 0
        Dim shipuniquecode As String = Mid(uniquecodevalue, 1, Len(uniquecodevalue) - 1)
        For a = 0 To x.Length - 1
            If x(a) = 1 Then
                If cvshiptype.getattribute(shipuniquecode, 14) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 2 Then
                If cvshiptype.getattribute(shipuniquecode, 16) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 3 Then
                If cvshiptype.getattribute(shipuniquecode, 16) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 4 Then
                If cvshiptype.getattribute(shipuniquecode, 15) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 5 Then
                If cvshiptype.getattribute(shipuniquecode, 18) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 6 Then
                If cvshiptype.getattribute(shipuniquecode, 17) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 11 Then
                If cvshiptype.getattribute(shipuniquecode, 19) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            ElseIf x(a) = 12 Then
                If cvshiptype.getattribute(shipuniquecode, 20) = 1 Then
                    enabletype = x(a)
                    Exit For
                End If
            End If
        Next
    End Function

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

Public Class shiprestrict_class
    Dim datavalue(19)

    Public Property id As String
        Get
            id = datavalue(0)
        End Get
        Set(value As String)
            datavalue(0) = value
        End Set
    End Property

    Public Property name As String
        Get
            name = datavalue(1)
        End Get
        Set(value As String)
            datavalue(1) = value
        End Set
    End Property

    Public Property anycplane As Integer
        Get
            anycplane = datavalue(2)
        End Get
        Set(value As Integer)
            datavalue(2) = value
        End Set
    End Property

    Public Property needcfighter As Integer
        Get
            needcfighter = datavalue(3)
        End Get
        Set(value As Integer)
            datavalue(3) = value
        End Set
    End Property

    Public Property needcboomer As Integer
        Get
            needcboomer = datavalue(4)
        End Get
        Set(value As Integer)
            datavalue(4) = value
        End Set
    End Property

    Public Property needctorpedo As Integer
        Get
            needctorpedo = datavalue(5)
        End Get
        Set(value As Integer)
            datavalue(5) = value
        End Set
    End Property

    Public Property needlso As Integer
        Get
            needlso = datavalue(6)
        End Get
        Set(value As Integer)
            datavalue(6) = value
        End Set
    End Property

    Public Property anywplane As Integer
        Get
            anywplane = datavalue(7)
        End Get
        Set(value As Integer)
            datavalue(7) = value
        End Set
    End Property

    Public Property needwboomer As Integer
        Get
            needwboomer = datavalue(8)
        End Get
        Set(value As Integer)
            datavalue(8) = value
        End Set
    End Property

    Public Property shipuniquecode As String
        Get
            shipuniquecode = datavalue(9)
        End Get
        Set(value As String)
            datavalue(9) = value
        End Set
    End Property

    Public Property equiplso As Integer
        Get
            equiplso = datavalue(10)
        End Get
        Set(value As Integer)
            datavalue(10) = value
        End Set
    End Property

    Public Property equipcfighter As Integer
        Get
            equipcfighter = datavalue(11)
        End Get
        Set(value As Integer)
            datavalue(11) = value
        End Set
    End Property

    Public Property equipcboomer As Integer
        Get
            equipcboomer = datavalue(12)
        End Get
        Set(value As Integer)
            datavalue(12) = value
        End Set
    End Property

    Public Property equipctorpedo As Integer
        Get
            equipctorpedo = datavalue(13)
        End Get
        Set(value As Integer)
            datavalue(13) = value
        End Set
    End Property

    Public Property equipcjet As Integer
        Get
            equipcjet = datavalue(14)
        End Get
        Set(value As Integer)
            datavalue(14) = value
        End Set
    End Property

    Public Property equipnfighter As Integer
        Get
            equipnfighter = datavalue(15)
        End Get
        Set(value As Integer)
            datavalue(15) = value
        End Set
    End Property

    Public Property equipntorpedo As Integer
        Get
            equipntorpedo = datavalue(16)
        End Get
        Set(value As Integer)
            datavalue(16) = value
        End Set
    End Property

    Public Property equipnothers As Integer
        Get
            equipnothers = datavalue(17)
        End Get
        Set(value As Integer)
            datavalue(17) = value
        End Set
    End Property

    Public Property getcfightercarry As Integer
        Get
            getcfightercarry = datavalue(18)
        End Get
        Set(value As Integer)
            datavalue(18) = value
        End Set
    End Property

    Public Property equipwboomer As Integer
        Get
            equipwboomer = datavalue(19)
        End Get
        Set(value As Integer)
            datavalue(19) = value
        End Set
    End Property

End Class

Public Class restrictcontrol_class
    Dim saa_cv_rd As New Xml.XmlDocument
    Dim cvrestricatdatagroup As basedatagroup_class
    Dim shiprestrict(19) As shiprestrict_class
    Dim startvalue As Integer = 0

    Public Sub New()
        filecontrol.loadfile("\base\SAA_cv_rd.xml", saa_cv_rd, cvrestricatdatagroup)
        For a = 0 To 19
            shiprestrict(a) = New shiprestrict_class
        Next
    End Sub

    Public Sub initialize(ByVal shipindex As Integer, ByVal shipuniquecode As String, ByVal restricatdataid As Integer)
        shiprestrict(shipindex).name = cvrestricatdatagroup.getattribute(restricatdataid, 1)
        shiprestrict(shipindex).id = cvrestricatdatagroup.getattribute(restricatdataid, 0)
        shiprestrict(shipindex).anycplane = cvrestricatdatagroup.getattribute(restricatdataid, 2)
        shiprestrict(shipindex).needcfighter = cvrestricatdatagroup.getattribute(restricatdataid, 3)
        shiprestrict(shipindex).needcboomer = cvrestricatdatagroup.getattribute(restricatdataid, 4)
        shiprestrict(shipindex).needctorpedo = cvrestricatdatagroup.getattribute(restricatdataid, 5)
        shiprestrict(shipindex).needlso = cvrestricatdatagroup.getattribute(restricatdataid, 6)
        shiprestrict(shipindex).anywplane = cvrestricatdatagroup.getattribute(restricatdataid, 7)
        shiprestrict(shipindex).needwboomer = cvrestricatdatagroup.getattribute(restricatdataid, 8)

        shiprestrict(shipindex).shipuniquecode = shipuniquecode

        If shiprestrict(shipindex).id >= 133 And shiprestrict(shipindex).id <= 136 Then
            shiprestrict(shipindex).equiplso = 1
        Else
            shiprestrict(shipindex).equiplso = 0
        End If

        shiprestrict(shipindex).equipcfighter = 0
        shiprestrict(shipindex).equipcboomer = 0
        shiprestrict(shipindex).equipctorpedo = 0
        shiprestrict(shipindex).equipcjet = 0
        shiprestrict(shipindex).equipnfighter = 0
        shiprestrict(shipindex).equipntorpedo = 0
        shiprestrict(shipindex).equipnothers = 0

        shiprestrict(shipindex).getcfightercarry = 0

        shiprestrict(shipindex).equipwboomer = 0
    End Sub

    Public Function getbaserestrictname(ByVal id As Integer) As String
        getbaserestrictname = cvrestricatdatagroup.getattribute(id, 1)
        If getbaserestrictname = "0" Then
            getbaserestrictname = ""
        End If
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
            If shiprestrict(a).needlso > 0 And cvship(a).active = True Then
                ReDim Preserve shipcarry(needlsonum)
                ReDim Preserve shipid(needlsonum)
                cvship(a).resetcarryobtained()
                For b = 0 To cvship(a).carrycount - 1 - 1
                    shipcarry(needlsonum) = shipcarry(needlsonum) + cvship(a).getcarry(10).amount
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


Public Class lbcarry_class
    Dim planeidvalue As String = "0"
    Dim skilledbuffvalue As Integer = 25

    Public Property planeid As String
        Get
            planeid = planeidvalue
        End Get
        Set(value As String)
            planeidvalue = value
        End Set
    End Property

    Public Property skilledbuff As Integer
        Get
            skilledbuff = skilledbuffvalue
        End Get
        Set(value As Integer)
            skilledbuffvalue = value
        End Set
    End Property

    ReadOnly Property amount As Integer
        Get
            amount = 18
            If planeidvalue <> "0" Then
                If plane.unit(planeidvalue).classification = 13 Then
                    amount = 4
                End If
            End If
        End Get
    End Property

    ReadOnly Property damage As Double
        Get
            damage = 0
            If planeidvalue <> "0" Then
                If plane.unit(planeidvalue).classification = 31 Then
                    damage = Int(0.8 * plane.unit(planeidvalue).torpedo * Math.Sqrt(1.8 * amount)) * 1.8
                ElseIf plane.unit(planeidvalue).classification = 2 Or plane.unit(planeidvalue).classification = 3 Or plane.unit(planeidvalue).classification = 4 Or plane.unit(planeidvalue).classification = 12 Then
                    damage = Int(1 * plane.unit(planeidvalue).torpedo * Math.Sqrt(1.8 * amount)) * 1
                End If
            End If
        End Get
    End Property

    ReadOnly Property attackAAvalue As Integer
        Get
            attackAAvalue = 0
            If planeidvalue <> "0" Then
                attackAAvalue = Int((plane.unit(planeidvalue).antiaircraft + plane.unit(planeidvalue).intercept * 1.5) * Math.Sqrt(amount)) + skilledbuffvalue
            End If
        End Get
    End Property

    ReadOnly Property defenseAAvalue As Integer
        Get
            defenseAAvalue = 0
            If planeidvalue <> "0" Then
                defenseAAvalue = Int((plane.unit(planeidvalue).antiaircraft + plane.unit(planeidvalue).intercept + plane.unit(planeidvalue).antibombing * 2) * Math.Sqrt(amount)) + skilledbuffvalue
            End If
        End Get
    End Property
End Class

Public Class plUIcontrol_class
    Dim stringcombo2 As New Collection
    Dim stringlist1 As New Collection
    Dim stringlist2 As New Collection

    Dim firstfleet As New Collection
    Dim secondfleet As New Collection

    Dim errorlist As New Collection

    Dim shipid As Integer
    Dim carryid As Integer

    Dim enhancedmode As Boolean = True
    Dim landbasemode As Boolean = True

    Private Sub list1refresh()
        stringlist1.Clear()
        Dim situationstring As String
        situationstring = "<一队制空-" & ffcaavalue + ffwaavalue & "><联合制空-" & ffcaavalue + ffwaavalue + sfcaavalue + sfwaavalue & "><警告-" & errorlist.Count & "项>"
        stringlist1.Add(situationstring)
        If firstfleet.Count <> 0 Then
            For a = 1 To firstfleet.Count
                If firstfleet(a) > 999 Then
                    shipid = firstfleet(a) / 1000 - 1
                    stringlist1.Add("<F1>" & cvshiptype.getattribute(cvship(shipid).uniquecode, 1) & "[" & restrictcontrol.getrestrictattribute(shipid, 1) & "]") '& ship(shipid).uniquecode)
                ElseIf firstfleet(a) <= 999 And firstfleet(a) >= 0 Then
                    shipid = Int(firstfleet(a) / 10)
                    carryid = firstfleet(a) - (Int(firstfleet(a) / 10) * 10)
                    If cvship(shipid).getcarry(carryid).planeid = 0 Then
                        stringlist1.Add("[" & Format(cvship(shipid).getcarry(carryid).amount, "00") & "]" & "[无搭载]")
                    Else
                        stringlist1.Add("[" & Format(cvship(shipid).getcarry(carryid).amount, "00") & "][" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, 1) & "]")
                    End If
                End If
            Next
        End If
        If secondfleet.Count <> 0 Then
            For a = 1 To secondfleet.Count
                If secondfleet(a) > 999 Then
                    shipid = secondfleet(a) / 1000 - 1
                    stringlist1.Add("<F2>" & cvshiptype.getattribute(cvship(shipid).uniquecode, 1) & "[" & restrictcontrol.getrestrictattribute(shipid, 1) & "]") ' & ship(shipid).uniquecode)
                ElseIf secondfleet(a) <= 999 And secondfleet(a) >= 0 Then
                    shipid = Int(secondfleet(a) / 10)
                    carryid = secondfleet(a) - (Int(secondfleet(a) / 10) * 10)
                    If cvship(shipid).getcarry(carryid).planeid = 0 Then
                        stringlist1.Add("[" & Format(cvship(shipid).getcarry(carryid).amount, "00") & "]" & "[无搭载]")
                    Else
                        stringlist1.Add("[" & Format(cvship(shipid).getcarry(carryid).amount, "00") & "][" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, 1) & "]")
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
                Do While restrictcontrol.getbaserestrictname(Int(cvship(shipid).restrict / 10) * 10 + num) <> ""
                    stringcombo2.Add(restrictcontrol.getbaserestrictname(Int(cvship(shipid).restrict / 10) * 10 + num))
                    num = num + 1
                Loop
            End If
        ElseIf x > firstfleet.Count Then
            stringcombo2.Clear()
            If secondfleet(x - firstfleet.Count) > 999 Then
                shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                Dim num As Integer = 0
                Do While restrictcontrol.getbaserestrictname(Int(cvship(shipid).restrict / 10) * 10 + num) <> ""
                    stringcombo2.Add(restrictcontrol.getbaserestrictname(Int(cvship(shipid).restrict / 10) * 10 + num))
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
                If cvship(a).active Then
                    shipcount = shipcount + 1
                    For b = 0 To cvship(a).carrycount - 1
                        If cvship(a).getcarry(b).planeid <> 0 Then
                            If plane.getattribute(cvship(a).getcarry(b).planeid, 3) <> 8 And plane.getattribute(cvship(a).getcarry(b).planeid, 3) <> 21 Then
                                planecount = planecount + cvship(a).getcarry(b).amount
                            End If
                        End If
                    Next
                    combominaviationfire = combominaviationfire + cvship(a).minaviationfire
                    combomaxaviationfire = combomaxaviationfire + cvship(a).maxaviationfire
                    If a <= 9 Then
                        firstminaviationfire = firstminaviationfire + cvship(a).minaviationfire
                        firstmaxaviationfire = firstmaxaviationfire + cvship(a).maxaviationfire
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
                            situationstring = "<警告：" & cvshiptype.getattribute(cvship(errorship).uniquecode, 1) & "-----"
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
                situationstring = "<第一舰队>" & cvshiptype.getattribute(cvship(shipid).uniquecode, 1) & "[" & restrictcontrol.getbaserestrictname(cvship(shipid).restrict) & "]"
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
                situationstring = "<制空值-" & cvship(shipid).AAvalue & "><航空开幕-(" & Int(cvship(shipid).minaviationfire) & "-" & Int(cvship(shipid).maxaviationfire) & ")>"
                stringlist2.Add(situationstring)
                If cvship(shipid).shellingfire <> -1 Then
                    If cvship(shipid).shellingfire > 0 Then
                        situationstring = "<航空炮击-" & cvship(shipid).shellingfire & ">"
                    Else
                        situationstring = "<航空炮击-无法炮击"
                    End If
                Else
                    situationstring = "<航空炮击-Null>"
                End If
                If cvship(shipid).CIshellingfire <> -1 Then
                    If restrictcontrol.CIstate(cvship(shipid).getcarry(0)) <> 0 Then
                        situationstring = situationstring & "<航空CI-" & cvship(shipid).CIshellingfire & "(" & restrictcontrol.CIstate(cvship(shipid).getcarry(0)) & ")>"
                    ElseIf restrictcontrol.CIstate(cvship(shipid).getcarry(0)) = 0 Then
                        situationstring = situationstring & "<航空CI-无法CI>"
                    End If
                Else
                    situationstring = situationstring & "<航空CI-Null>"
                End If
                stringlist2.Add(situationstring)
                If cvship(shipid).nightfightfire <> -1 Then
                    If restrictcontrol.nfstate(cvship(shipid).getcarry(0)) >= 1 Then
                        situationstring = "<夜战炮击-" & cvship(shipid).nightfightfire & ">"
                    ElseIf restrictcontrol.nfstate(cvship(shipid).getcarry(0)) < 1 Then
                        situationstring = "<夜战炮击-无法夜战>"
                    End If
                Else
                    situationstring = "<夜战炮击-Null>"
                End If
                If cvship(shipid).CInightfightfire <> -1 Then
                    If restrictcontrol.nfstate(cvship(shipid).getcarry(0)) > 1 Then
                        situationstring = situationstring & "<夜战CI-" & cvship(shipid).CInightfightfire & "(" & restrictcontrol.nfstate(cvship(shipid).getcarry(0)) & ")>"
                    ElseIf restrictcontrol.nfstate(cvship(shipid).getcarry(0)) <= 1 Then
                        situationstring = situationstring & "<夜战CI-无法CI>"
                    End If
                Else
                    situationstring = situationstring & "<夜战CI-Null>"
                End If
                stringlist2.Add(situationstring)
                Dim carrycount As Integer = 0
                planecount = 0
                For a = 0 To cvship(shipid).carrycount - 1
                    If cvship(shipid).getcarry(a).planeid <> 0 Then
                        If plane.getattribute(cvship(shipid).getcarry(a).planeid, 3) <> 8 And plane.getattribute(cvship(shipid).getcarry(a).planeid, 3) <> 21 Then
                            carrycount = carrycount + 1
                            planecount = planecount + cvship(shipid).getcarry(a).amount
                        End If
                    End If
                Next
                situationstring = "====<已搭载舰载机-" & carrycount & "格-" & planecount & "架>=========================="
                stringlist2.Add(situationstring)
                For a = 0 To cvship(shipid).carrycount - 1
                    situationstring = "[" & Format(cvship(shipid).getcarry(a).amount, "00") & "]["
                    If cvship(shipid).getcarry(a).planeid <> 0 Then
                        situationstring = situationstring & plane.getattribute(cvship(shipid).getcarry(a).planeid, 1)
                    Else
                        situationstring = situationstring & "无搭载"
                    End If
                    situationstring = situationstring & "]<" & cvship(shipid).getcarry(a).AAvalue & "><(" & Int(cvship(shipid).getcarry(a).minaviationfire) & "-" & Int(cvship(shipid).getcarry(a).maxaviationfire) & ")>"
                    stringlist2.Add(situationstring)
                Next
            ElseIf carryid <> -1 Then
                situationstring = "<" & cvshiptype.getattribute(cvship(shipid).uniquecode, 1) & "-" & carryid + 1 & ">[" & cvship(shipid).getcarry(carryid).amount & "]"
                If cvship(shipid).getcarry(carryid).planeid <> 0 Then
                    situationstring = situationstring & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, 1)
                Else
                    situationstring = situationstring & "无搭载"
                End If
                situationstring = situationstring & "<" & cvship(shipid).getcarry(carryid).AAvalue & "><(" & Int(cvship(shipid).getcarry(carryid).minaviationfire) & "-" & Int(cvship(shipid).getcarry(carryid).maxaviationfire) & ")>"
                stringlist2.Add(situationstring)
                situationstring = "<改修-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, 17) & ">"
                stringlist2.Add(situationstring)
                For a = 4 To 12
                    If plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) <> 0 Then
                        If a = 4 Then
                            situationstring = "<火力-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 5 Then
                            situationstring = "<雷装-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 6 Then
                            situationstring = "<对空-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 7 Then
                            situationstring = "<对潜-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 8 Then
                            situationstring = "<爆装-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 9 Then
                            situationstring = "<命中-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 10 Then
                            situationstring = "<装甲-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 11 Then
                            situationstring = "<回避-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
                        ElseIf a = 12 Then
                            situationstring = "<索敌-" & plane.getattribute(cvship(shipid).getcarry(carryid).planeid, a) & ">"
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
            For a = 0 To cvship(shipid).carrycount - 1
                firstfleet.Add(shipid * 10 + a)
            Next
        ElseIf x > 9 Then
            shipid = (x + 1) * 1000
            secondfleet.Add(shipid)
            shipid = shipid / 1000 - 1
            For a = 0 To cvship(shipid).carrycount - 1
                secondfleet.Add(shipid * 10 + a)
            Next
        End If
    End Sub

    Public Sub setshipattribute(ByVal x As Integer, ByVal y As Integer)
        If x > 0 And x <= firstfleet.Count Then
            If firstfleet(x) > 999 Then
                shipid = firstfleet(x) / 1000 - 1
                cvship(shipid).restrict = Int(cvship(shipid).restrict / 10) * 10 + y
            End If
        ElseIf x > firstfleet.Count Then
            If secondfleet(x - firstfleet.Count) > 999 Then
                shipid = secondfleet(x - firstfleet.Count) / 1000 - 1
                cvship(shipid).restrict = Int(cvship(shipid).restrict / 10) * 10 + y
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
                cvship(shipid).clear()
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
                cvship(shipid).clear()
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
            cvship(a).clear()
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
                    configurationcode = configurationcode & Chr(34) & "s" & ffcount & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & cvshiptype.getattribute(cvship(shipid).uniquecode, 2) & "," & Chr(34) & "luck" & Chr(34) & ":-1," & Chr(34) & "items" & Chr(34) & ":{"
                    For b = 0 To cvship(shipid).carrycount - 1
                        If cvship(shipid).getcarry(b).planeid <> 0 Then
                            itemcode = Chr(34) & "i" & b + 1 & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & plane.getattribute(cvship(shipid).getcarry(b).planeid, 2) & "," & Chr(34) & "rf" & Chr(34) & ":" & plane.getattribute(cvship(shipid).getcarry(b).planeid, 17) & "," & Chr(34) & "mas" & Chr(34) & ":"
                            If plane.getattribute(cvship(shipid).getcarry(b).planeid, 3) = 8 Then
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
                    configurationcode = configurationcode & Chr(34) & "s" & sfcount & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & cvshiptype.getattribute(cvship(shipid).uniquecode, 2) & "," & Chr(34) & "luck" & Chr(34) & ":-1," & Chr(34) & "items" & Chr(34) & ":{"
                    For b = 0 To cvship(shipid).carrycount - 1
                        If cvship(shipid).getcarry(b).planeid <> 0 Then
                            itemcode = Chr(34) & "i" & b + 1 & Chr(34) & ":{" & Chr(34) & "id" & Chr(34) & ":" & plane.getattribute(cvship(shipid).getcarry(b).planeid, 2) & "," & Chr(34) & "rf" & Chr(34) & ":" & plane.getattribute(cvship(shipid).getcarry(b).planeid, 22) & "," & Chr(34) & "mas" & Chr(34) & ":"
                            If plane.getattribute(cvship(shipid).getcarry(b).planeid, 3) = 8 Then
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

    Public Sub changeenhancedmode()
        If enhancedmode = False Then
            saa_plane.Height = saa_plane.Height + 40
            saa_plane.Label4.Top = saa_plane.Label4.Top + 40
            saa_plane.Label4.Text = "▲"
        ElseIf enhancedmode = True Then
            saa_plane.Height = saa_plane.Height - 40
            saa_plane.Label4.Top = saa_plane.Label4.Top - 40
            saa_plane.Label4.Text = "▼"
        End If
        enhancedmode = Not enhancedmode
    End Sub

    Public Sub changelandbasemode()
        If landbasemode = False Then
            saa_plane.Width = saa_plane.Width + 426
        ElseIf landbasemode = True Then
            saa_plane.Width = saa_plane.Width - 426
        End If
    End Sub
End Class

