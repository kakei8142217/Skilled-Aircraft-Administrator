Public Class smallgununit_class
    Inherits equipmentunit_class

    Sub New(ByVal newuniquecode As String)
        MyBase.New(newuniquecode)
    End Sub

    ReadOnly Property name As String
        Get
            name = userddequipmentgroup.getattribute(MyBase.extendid, 1)
        End Get
    End Property

    ReadOnly Property classification As Integer
        Get
            classification = ddequipmentgroup.getattribute(MyBase.baseid, 2)
        End Get
    End Property

    ReadOnly Property improve As Integer
        Get
            improve = userddequipmentgroup.getattribute(MyBase.extendid, 2)
        End Get
    End Property

    Overrides ReadOnly Property fire As Double
        Get
            fire = userddequipmentgroup.getattribute(MyBase.extendid, 5)
        End Get
    End Property

    Overrides ReadOnly Property torpedo As Double
        Get
            torpedo = userddequipmentgroup.getattribute(MyBase.extendid, 6)
        End Get
    End Property

End Class

Public Class smallgun_class
    Dim smallgununit() As smallgununit_class
    Dim smallgunlist As New Collection

    Dim usedlist As New Collection
    Dim temusedlist As New Collection


    Sub New()
        Call loadsmallgunnuit(0)
    End Sub

    Public Sub loadsmallgunnuit(ByVal listindex As Integer)
        ReDim smallgununit(-1)
        smallgunlist.Clear()
        Dim unitcount As Integer = 0
        Dim sameplanecount As Integer
        Dim newuniquecode As String
        If userddequipmentgroup IsNot Nothing Then
            If userddequipmentgroup.length > 0 Then
                For a = 0 To userddequipmentgroup.length - 1
                    If userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), listindex + 3) <> 0 Then
                        unitcount = unitcount + userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), listindex + 3)
                        ReDim Preserve smallgununit(unitcount - 1)
                        sameplanecount = 1
                        For b = unitcount - userddequipmentgroup.getattribute(userddequipmentgroup.getid(a), listindex + 3) To unitcount - 1
                            newuniquecode = Trim(Str(Val(userddequipmentgroup.getid(a)) * 100 + sameplanecount))
                            smallgununit(b) = New smallgununit_class(newuniquecode)
                            smallgunlist.Add(b, newuniquecode)
                            sameplanecount = sameplanecount + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    ReadOnly Property unit(ByVal uniquecode As String) As smallgununit_class
        Get
            unit = smallgununit(smallgunlist.Item(uniquecode))
        End Get
    End Property

    Public Function getbasesmallgunlist(ByVal classification As Integer) As String()
        Dim list() As String
        Dim exist As Boolean
        For a = 0 To smallgununit.Length - 1
            If smallgununit(a).classification = classification Then
                exist = False
                Try
                    If list.Length > 0 Then
                        For b = 0 To list.Length - 1
                            If list(b) = smallgununit(a).baseid Then
                                exist = True
                            End If
                        Next
                    End If
                Catch
                End Try
                If exist = False Then
                    Try
                        ReDim Preserve list(list.Length)
                    Catch
                        ReDim Preserve list(0)
                    End Try
                    list(list.Length - 1) = smallgununit(a).baseid
                End If
            End If
        Next
        getbasesmallgunlist = list
    End Function

    Public Overloads Function getsmallgun(ByVal classification As Integer, Optional ByVal limit As Integer = 0) As String
        getsmallgun = "0"

        Dim unuselist As New Collection
        Dim fattributegroup As New Collection
        Dim sattributegroup As New Collection


        If smallgununit IsNot Nothing Then
            For a = 0 To smallgununit.Length - 1
                If smallgununit(a).classification = classification Then
                    If used(smallgununit(a).uniquecode) = False Then
                        unuselist.Add(smallgununit(a).uniquecode)
                        Dim fattribute As Double = smallgununit(a).fire + smallgununit(a).torpedo
                        fattributegroup.Add(fattribute)
                        Dim sattribute As Double = smallgununit(a).hit + smallgununit(a).avoid + smallgununit(a).armor + smallgununit(a).spotting
                        sattributegroup.Add(sattribute)
                    End If
                End If

            Next
        End If


        If unuselist.Count <> 0 Then
            Dim smallgunid As String = "0"
            Dim buffid As String = "0"
            Dim fattribute As Double = limit
            Dim sattribute As Double
            If fattribute >= 0 Then
                For a = 1 To unuselist.Count
                    If fattributegroup(a) > fattribute Then
                        fattribute = fattributegroup(a)
                    End If
                Next
            End If
            For a = 1 To unuselist.Count
                If fattributegroup(a) = fattribute Then
                    If sattributegroup(a) >= sattribute Then
                        smallgunid = unuselist(a)
                        sattribute = sattributegroup(a)
                    End If
                End If
            Next
            getsmallgun = smallgunid
        End If
    End Function

    Public Overloads Function getsmallgun(ByVal baseid As String, Optional ByVal limit As Integer = 0) As String
        getsmallgun = "0"

        Dim unuselist As New Collection
        Dim fattributegroup As New Collection
        Dim sattributegroup As New Collection


        If smallgununit IsNot Nothing Then
            For a = 0 To smallgununit.Length - 1
                If smallgununit(a).baseid = baseid Then
                    If used(smallgununit(a).uniquecode) = False Then
                        unuselist.Add(smallgununit(a).uniquecode)
                        Dim fattribute As Double = smallgununit(a).fire + smallgununit(a).torpedo
                        fattributegroup.Add(fattribute)
                        Dim sattribute As Double = smallgununit(a).hit + smallgununit(a).avoid + smallgununit(a).armor + smallgununit(a).spotting
                        sattributegroup.Add(sattribute)
                    End If
                End If

            Next
        End If


        If unuselist.Count <> 0 Then
            Dim smallgunid As String = "0"
            Dim buffid As String = "0"
            Dim fattribute As Double = limit
            Dim sattribute As Double
            If fattribute >= 0 Then
                For a = 1 To unuselist.Count
                    If fattributegroup(a) > fattribute Then
                        fattribute = fattributegroup(a)
                    End If
                Next
            End If
            For a = 1 To unuselist.Count
                If fattributegroup(a) = fattribute Then
                    If sattributegroup(a) >= sattribute Then
                        smallgunid = unuselist(a)
                        sattribute = sattributegroup(a)
                    End If
                End If
            Next
            getsmallgun = smallgunid
        End If
    End Function

    Public Function extrabuff(ByVal ship As ddship_class) As Integer()
        Dim bufflist(ship.length - 1) As Integer
        Dim id As String
        For a = 0 To ship.length - 1
            id = ship.carry(a).equipmentid
            If id <> "0" Then
                Dim baseid As String
                baseid = smallgun.unit(id).baseid
                For b = 0 To ddextrabuffgroup.length - 1
                    Dim buffid As String = ddextrabuffgroup.getid(b)
                    If ddextrabuffgroup.getattribute(buffid, 1) = 1 Then
                        If ddextrabuffgroup.getattribute(buffid, 2) = baseid Then
                            If smallgun.unit(id).improve >= ddextrabuffgroup.getattribute(buffid, 3) And smallgun.unit(id).improve <= ddextrabuffgroup.getattribute(buffid, 4) Then
                                Dim shiptypematch As Boolean = False
                                If userddshiptype.unit(ship.id).baseid = ddextrabuffgroup.getattribute(buffid, 5) Then
                                    shiptypematch = True
                                ElseIf userddshiptype.unit(ship.id).type >= ddextrabuffgroup.getattribute(buffid, 6) And userddshiptype.unit(ship.id).type <= ddextrabuffgroup.getattribute(buffid, 7) Then
                                    shiptypematch = True
                                End If
                                If shiptypematch Then
                                    Dim cheshi = ship.equipmentcount(a + 1, {baseid}, {ddextrabuffgroup.getattribute(buffid, 3)}, {ddextrabuffgroup.getattribute(buffid, 4)})
                                    bufflist(a) = ddextrabuffgroup.getattribute(buffid, cheshi + 7)
                                End If
                            End If
                        End If
                    ElseIf ddextrabuffgroup.getattribute(buffid, 1) = 2 Then
                        If ddextrabuffgroup.getattribute(buffid, 2) = baseid Then
                            If smallgun.unit(id).improve >= ddextrabuffgroup.getattribute(buffid, 3) And smallgun.unit(id).improve <= ddextrabuffgroup.getattribute(buffid, 4) Then

                                For c = 0 To ddextrabuffgroup.getdata(buffid).getdatagroup.length - 1
                                    Dim class2buffid As String = ddextrabuffgroup.getdata(buffid).getdatagroup.getid(c)
                                    Dim shiptypematch As Boolean = False
                                    If userddshiptype.unit(ship.id).baseid = ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 1) Then
                                        shiptypematch = True
                                    ElseIf userddshiptype.unit(ship.id).type >= ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 2) And userddshiptype.unit(ship.id).type <= ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 3) Then
                                        shiptypematch = True
                                    End If
                                    If shiptypematch Then
                                        Dim baseidlist(ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.length - 1) As String
                                        Dim improvemin(ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.length - 1) As Integer
                                        Dim improvemax(ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.length - 1) As Integer
                                        For d = 0 To ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.length - 1
                                            Dim class3buffid As String = ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.getid(d)
                                            baseidlist(d) = ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.getattribute(class3buffid, 1)
                                            improvemin(d) = ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.getattribute(class3buffid, 2)
                                            improvemax(d) = ddextrabuffgroup.getdata(buffid).getdatagroup.getdata(class2buffid).getdatagroup.getattribute(class3buffid, 3)
                                        Next
                                        Dim count As Integer = ship.equipmentcount(a + 1, baseidlist, improvemin, improvemax)
                                        If count > 0 Then
                                            bufflist(2) = bufflist(2) + (ddextrabuffgroup.getdata(buffid).getdatagroup.getattribute(class2buffid, 3 + count))
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next
            End If
        Next
        extrabuff = bufflist
    End Function

    Private Function used(ByVal uniquecode As String) As Boolean
        used = False
        If temusedlist.Contains(uniquecode) Then
            used = True
        End If
        If usedlist.Contains(uniquecode) Then
            used = True
        End If
    End Function

    Public Sub addcollection(ByVal uniquecode As String, Optional ByVal mode As Integer = 0)
        If mode = 0 Then
            temusedlist.Add(uniquecode, uniquecode)
        Else
            usedlist.Add(uniquecode, uniquecode)
        End If
    End Sub

    Public Sub clearcollection(Optional ByVal mode As Integer = 0)
        temusedlist.Clear()
        If mode = 1 Then
            usedlist.Clear()
        End If
    End Sub
End Class



Public Class ddcarry_class
    Dim typevalue As Integer
    Dim shipidvalue As String
    Dim equipmentidvalue As String = "0"
    Dim buffvalue As Integer

    Sub New(ByVal newshipid As String, ByVal newtype As Integer)
        shipidvalue = newshipid
        typevalue = newtype
    End Sub

    ReadOnly Property shipid As String
        Get
            shipid = shipidvalue
        End Get
    End Property

    ReadOnly Property type As Integer
        Get
            type = typevalue
        End Get
    End Property

    Public Property equipmentid As String
        Get
            equipmentid = equipmentidvalue
        End Get
        Set(value As String)
            equipmentidvalue = value
        End Set
    End Property

    Public Property buff As Integer
        Get
            buff = buffvalue
        End Get
        Set(value As Integer)
            buffvalue = value
        End Set
    End Property

    ReadOnly Property damage As Double
        Get
            If equipmentidvalue <> "0" Then
                damage = smallgun.unit(equipmentidvalue).fire + smallgun.unit(equipmentidvalue).torpedo + buffvalue
            Else
                damage = 0
            End If
        End Get
    End Property


    ReadOnly Property showstring As String
        Get
            showstring = "[未装备]"
            If equipmentidvalue <> "0" Then
                showstring = "[" & smallgun.unit(equipmentidvalue).name & "][" & Format(smallgun.unit(equipmentidvalue).fire + smallgun.unit(equipmentidvalue).torpedo, "0.##") & "][" & buffvalue & "]"
            End If
        End Get
    End Property
End Class

Public Class ddship_class
    Dim ddcarry() As ddcarry_class
    Dim shipidvalue As String
    Dim nfmodevalue As Integer = 0
    Dim buffmultiplevalue As Double = 1
    Dim restrictmodevalue As Integer = 0
    Dim enhancedlockvalue As Integer = 0

    Sub New(ByVal newshipid As String)
        shipidvalue = newshipid

        ReDim ddcarry(userddshiptype.unit(shipidvalue).gridcount - 1)
        For a = 0 To ddcarry.Length - 1
            ddcarry(a) = New ddcarry_class(shipidvalue, 0)
        Next
        If userddshiptype.unit(shipidvalue).enhanced = 1 Then
            ReDim Preserve ddcarry(ddcarry.Length)
            ddcarry(ddcarry.Length - 1) = New ddcarry_class(shipidvalue, 1)
        End If
    End Sub

    ReadOnly Property id As String
        Get
            id = shipidvalue
        End Get
    End Property

    ReadOnly Property carry(ByVal index As Integer) As ddcarry_class
        Get
            carry = ddcarry(index)
        End Get
    End Property

    ReadOnly Property length As Integer
        Get
            If ddcarry IsNot Nothing Then
                length = ddcarry.Length
            Else
                length = 0
            End If
        End Get
    End Property

    Public Property nfmode As Integer
        Get
            nfmode = nfmodevalue
        End Get
        Set(value As Integer)
            nfmodevalue = value
        End Set
    End Property

    Public Property buffmultiple As Double
        Get
            buffmultiple = buffmultiplevalue
        End Get
        Set(value As Double)
            buffmultiplevalue = value
        End Set
    End Property

    Public Property restrictmode As Integer
        Get
            restrictmode = restrictmodevalue
        End Get
        Set(value As Integer)
            restrictmodevalue = value
        End Set
    End Property

    Public Property enhancedlock As Integer
        Get
            enhancedlock = enhancedlockvalue
        End Get
        Set(value As Integer)
            enhancedlockvalue = value
        End Set
    End Property

    ReadOnly Property damage As Double
        Get
            Dim value() As Double = basedamage()
            If value(1) > 1 Then
                damage = value(0) * 2
            Else
                damage = value(0)
            End If
        End Get
    End Property

    ReadOnly Property showstring As String
        Get
            Dim stringvalue As String
            Dim value() As Double = basedamage()
            stringvalue = "<" & value(0)
            If value(1) > 1 Then
                stringvalue = stringvalue & "×2"
            Else
                stringvalue = stringvalue & "×1"
            End If
            stringvalue = stringvalue & ">[" & userddshiptype.unit(shipidvalue).name & "]"
            If nfmodevalue = 0 Then
                If equipmentcount(0) = 0 Then
                    stringvalue = stringvalue & "[自动判断]"
                ElseIf equipmentcount(1) < 2 And equipmentcount(2) < 2 Then
                    stringvalue = stringvalue & "[自动判断(单击)]"
                ElseIf equipmentcount(1) >= 2 Then
                    stringvalue = stringvalue & "[自动判断(二连)]"
                ElseIf equipmentcount(2) >= 2 Then
                    stringvalue = stringvalue & "[自动判断(CI)]"
                End If
            ElseIf nfmodevalue = 1 Then
                stringvalue = stringvalue & "[强制二连]"
            ElseIf nfmodevalue = 2 Then
                stringvalue = stringvalue & "[强制CI]"
            End If

            If restrictmodevalue = 0 Then
                stringvalue = stringvalue & "[无限制]"
            ElseIf restrictmodevalue = 1 Then
                stringvalue = stringvalue & "[锁定一格]"
            ElseIf restrictmodevalue = 2 Then
                stringvalue = stringvalue & "[锁定两格]"
            End If

            If userddshiptype.unit(shipidvalue).enhanced = 1 Then
                If enhancedlockvalue = 0 Then
                    stringvalue = stringvalue & "[×]"
                ElseIf enhancedlockvalue = 1 Then
                    stringvalue = stringvalue & "[○]"
                End If
            End If

            stringvalue = stringvalue & "[" & Format(buffmultiplevalue, "0.00") & "]"

            showstring = stringvalue
        End Get
    End Property

    Public Overloads Function equipmentcount(ByVal classification As Integer) As Integer
        Dim count As Integer = 0
        For a = 0 To ddcarry.Length - 1
            If ddcarry(a).equipmentid <> "0" Then
                If classification = 0 Then
                    count = count + 1
                ElseIf smallgun.unit(ddcarry(a).equipmentid).classification = classification Then
                    count = count + 1
                End If
            End If
        Next
        equipmentcount = count
    End Function

    Public Overloads Function equipmentcount(ByVal firstto As Integer, ByVal baseid As String(), ByVal improvemin As Integer(), ByVal improvemax As Integer()) As Integer
        Dim count As Integer = 0
        For a = 0 To firstto - 1
            For b = 0 To baseid.Length - 1
                If ddcarry(a).equipmentid <> "0" Then
                    If smallgun.unit(ddcarry(a).equipmentid).baseid = baseid(b) Then
                        If smallgun.unit(ddcarry(a).equipmentid).improve <= improvemax(b) And smallgun.unit(ddcarry(a).equipmentid).improve >= improvemin(b) Then
                            count = count + 1
                        End If
                    End If
                End If
            Next
        Next
        equipmentcount = count
    End Function

    Public Sub removeequipment(Optional ByVal mode As Integer = 0)
        For a = 0 To length - 1
            ddcarry(a).equipmentid = "0"
            ddcarry(a).buff = 0
        Next
    End Sub

    Private Function basedamage() As Double()
        Dim value(1) As Double
        value(0) = userddshiptype.unit(shipidvalue).damage
        For a = 0 To length - 1
            value(0) = value(0) + ddcarry(a).damage
        Next
        If equipmentcount(1) >= 2 Then
            value(1) = 1.2
        ElseIf equipmentcount(2) >= 2 Then
            value(1) = 1.5
        Else
            value(1) = 1
        End If
        value(0) = value(0) * value(1)
        If value(0) > 300 Then
            value(0) = Math.Sqrt(value(0) - 300) + 300
        End If
        value(0) = Int(value(0) * buffmultiplevalue)
        basedamage = value
    End Function
End Class

Public Class ddshipgroup_class
    Dim ddship() As ddship_class

    ReadOnly Property ship(ByVal index As Integer) As ddship_class
        Get
            ship = ddship(index)
        End Get
    End Property

    ReadOnly Property length As Integer
        Get
            If ddship IsNot Nothing Then
                length = ddship.Length
            Else
                length = 0
            End If
        End Get
    End Property

    Public Sub addship(ByVal newshipid As String, Optional ByVal mode As Integer = 0)
        If Not (exist(newshipid) = True And mode = 0) Then
            ReDim Preserve ddship(length)
            ddship(ddship.Length - 1) = New ddship_class(newshipid)
        End If
    End Sub

    Public Sub removeship(ByVal index As Integer)
        If 0 <= index < length Then
            If index <> length - 1 Then

                For a = index + 1 To length - 1
                    ddship(a - 1) = ddship(a)
                Next
            End If
            ReDim Preserve ddship(ddship.Length - 2)
        End If
    End Sub

    Public Sub removeallship()
        ReDim Preserve ddship(-1)
    End Sub

    Private Function exist(ByVal shipid As String) As Boolean
        exist = False
        If ddship IsNot Nothing Then
            For a = 0 To ddship.Length - 1
                If ddship(a).id = shipid Then exist = True
            Next
        End If
    End Function

    Public Sub sort()
        If ddship IsNot Nothing Then
            Dim cache As ddship_class
            For a = 0 To length - 1
                For b = a To length - 1
                    If ddship(a).damage < ddship(b).damage Then
                        cache = ddship(a)
                        ddship(a) = ddship(b)
                        ddship(b) = cache
                    End If
                Next
            Next

        End If
    End Sub
End Class



Public Class ddshiptypeunit_class
    Inherits shiptypeunit_class

    Sub New(ByVal newuniquecode As String)
        MyBase.New(newuniquecode)
    End Sub

    ReadOnly Property modification As Integer
        Get
            modification = ddshiptypegroup.getattribute(MyBase.baseid, 2)
        End Get
    End Property

    ReadOnly Property type As Integer
        Get
            type = ddshiptypegroup.getattribute(MyBase.baseid, 3)
        End Get
    End Property

    ReadOnly Property extraarmour As Integer
        Get
            extraarmour = ddshiptypegroup.getattribute(MyBase.baseid, 4)
        End Get
    End Property

    ReadOnly Property dropship As Integer
        Get
            dropship = ddshiptypegroup.getattribute(MyBase.baseid, 5)
        End Get
    End Property

    ReadOnly Property tank As Integer
        Get
            tank = ddshiptypegroup.getattribute(MyBase.baseid, 6)
        End Get
    End Property

    ReadOnly Property headquarters As Integer
        Get
            headquarters = ddshiptypegroup.getattribute(MyBase.baseid, 7)
        End Get
    End Property

    ReadOnly Property name As String
        Get
            name = userddshiptypegroup.getattribute(MyBase.extendid, 1)
        End Get
    End Property

    ReadOnly Property level As Integer
        Get
            level = userddshiptypegroup.getattribute(MyBase.extendid, 2)
        End Get
    End Property

    ReadOnly Property luck As Integer
        Get
            luck = userddshiptypegroup.getattribute(MyBase.extendid, 3)
        End Get
    End Property

    ReadOnly Property enhanced As Integer
        Get
            enhanced = userddshiptypegroup.getattribute(MyBase.extendid, 4)
        End Get
    End Property

    ReadOnly Property damage As Integer
        Get
            damage = MyBase.fire + MyBase.torpedo
        End Get
    End Property

    ReadOnly Property antisubmarine As Integer
        Get
            antisubmarine = Int(MyBase.antisubmarine99 + (MyBase.antisubmarine165 - MyBase.antisubmarine99) / (Me.level - 99) * 66)
        End Get
    End Property

    ReadOnly Property avoid As Integer
        Get
            avoid = Int(MyBase.avoid99 + (MyBase.avoid165 - MyBase.avoid99) / (Me.level - 99) * 66)
        End Get
    End Property

    ReadOnly Property spotting As Integer
        Get
            spotting = Int(MyBase.spotting99 + (MyBase.spotting165 - MyBase.spotting99) / (Me.level - 99) * 66)
        End Get
    End Property
End Class

Public Class userddshiptype_class
    Dim ddshiptypeunit() As ddshiptypeunit_class
    Dim ddshiptypelist As New Collection

    Dim modificationvalue As Integer = 1

    Sub New()
        Call loadddshiptypeunit()
    End Sub

    Public Sub loadddshiptypeunit()
        ReDim ddshiptypeunit(-1)
        ddshiptypelist.Clear()
        Dim unitcount As Integer = 0
        Dim sameshiptypecount As Integer
        Dim newuniquecode As String
        If userddshiptypegroup IsNot Nothing Then
            If userddshiptypegroup.length > 0 Then
                For a = 0 To userddshiptypegroup.length - 1
                    If userddshiptypegroup.getattribute(userddshiptypegroup.getid(a), 5) <> 0 Then
                        unitcount = unitcount + userddshiptypegroup.getattribute(userddshiptypegroup.getid(a), 5)
                        ReDim Preserve ddshiptypeunit(unitcount - 1)
                        sameshiptypecount = 1
                        For b = unitcount - userddshiptypegroup.getattribute(userddshiptypegroup.getid(a), 5) To unitcount - 1
                            newuniquecode = Trim(Str(Val(userddshiptypegroup.getid(a)) * 10 + sameshiptypecount))
                            ddshiptypeunit(b) = New ddshiptypeunit_class(newuniquecode)
                            ddshiptypelist.Add(b, newuniquecode)
                            sameshiptypecount = sameshiptypecount + 1
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub unitsort(ByVal mode As String)
        If ddshiptypeunit IsNot Nothing Then
            Dim cache As ddshiptypeunit_class
            ddshiptypelist.Clear()
            For a = 0 To ddshiptypeunit.Length - 1
                For b = a To ddshiptypeunit.Length - 1
                    If mode = "level" Then
                        If ddshiptypeunit(a).level < ddshiptypeunit(b).level Then
                            cache = ddshiptypeunit(a)
                            ddshiptypeunit(a) = ddshiptypeunit(b)
                            ddshiptypeunit(b) = cache
                        End If
                    ElseIf mode = "damage" Then
                        If ddshiptypeunit(a).damage < ddshiptypeunit(b).damage Then
                            cache = ddshiptypeunit(a)
                            ddshiptypeunit(a) = ddshiptypeunit(b)
                            ddshiptypeunit(b) = cache
                        End If
                    ElseIf mode = "type" Then
                        If ddshiptypeunit(a).type > ddshiptypeunit(b).type Then
                            cache = ddshiptypeunit(a)
                            ddshiptypeunit(a) = ddshiptypeunit(b)
                            ddshiptypeunit(b) = cache
                        End If
                    End If
                Next
                ddshiptypelist.Add(a, ddshiptypeunit(a).uniquecode)
            Next
        End If
    End Sub

    Public Sub edituserddshiptype(ByVal targetid As String, ByVal newid As String)
        Dim userddshiptypedata As New basedata_class(6)
        If userddshiptypegroup Is Nothing Then
            userddshiptypegroup = New basedatagroup_class(6)
        End If

        If newid <> "0" Then
            userddshiptypedata.setattribute(0, newid)
            Dim baseid As String = Int(newid / 10000000)
            Dim level As Integer = Int((newid - baseid * 10000000) / 10000)
            Dim luck As Integer = Int((newid - baseid * 10000000 - level * 10000) / 10)
            Dim enhanced As Integer = newid - baseid * 10000000 - level * 10000 - luck * 10
            Dim name As String = ddshiptypegroup.getattribute(baseid, 1)
            name = name & "(" & level & "|" & luck
            If enhanced = 1 Then
                name = name & "|○)"
            Else
                name = name & ")"
            End If
            userddshiptypedata.setattribute(1, name)
            userddshiptypedata.setattribute(2, level)
            userddshiptypedata.setattribute(3, luck)
            userddshiptypedata.setattribute(4, enhanced)
        End If

        If targetid <> "0" Then
            If userddshiptypegroup.exist(targetid) Then
                Dim editdata As basedata_class = userddshiptypegroup.getdata(targetid)
                If editdata.getattribute(5) > 1 Then
                    editdata.setattribute(5, editdata.getattribute(5) - 1)
                Else
                    userddshiptypegroup.removedata(userddshiptypegroup.getdata(targetid))
                End If
                If newid <> "0" Then
                    userddshiptypedata.setattribute(5, 1)
                    userddshiptypegroup.setdata(userddshiptypedata)
                End If
            End If
        Else
            If userddshiptypegroup.exist(newid) Then
                userddshiptypedata.setattribute(5, userddshiptypegroup.getattribute(newid, 5) + 1)
            Else
                userddshiptypedata.setattribute(5, 1)
            End If
            userddshiptypegroup.setdata(userddshiptypedata)
        End If




        Dim insert As Boolean = False
        If saa_u_dd_std.DocumentElement.ChildNodes.Count <> 0 Then
            For Each node As Xml.XmlElement In saa_u_dd_std.DocumentElement.ChildNodes
                If targetid <> "0" Then
                    If node.Attributes("id").Value = targetid Then
                        If node.Attributes("amountlist1").Value > 1 Then
                            node.SetAttribute("amountlist1", node.Attributes("amountlist1").Value - 1)
                        Else
                            saa_u_dd_std.DocumentElement.RemoveChild(node)
                        End If
                        If newid = "0" Then
                                insert = True
                            End If
                        End If
                    End If
                If newid <> "0" Then
                    If node.Attributes("id").Value = userddshiptypedata.getattribute(0) Then
                        node.SetAttribute("amountlist1", userddshiptypedata.getattribute(5))
                        insert = True
                        Exit For
                    End If
                End If
            Next
        End If
        If insert = False Then
            Dim node As Xml.XmlElement = saa_u_dd_std.CreateElement("myddshiptype")
            Dim attname() As String = {"id", "name", "level", "luck", "enhanced", "amountlist1"}
            For a = 0 To attname.Length - 1
                node.SetAttribute(attname(a), userddshiptypedata.getattribute(a))
            Next
            saa_u_dd_std.DocumentElement.AppendChild(node)
        End If
        saa_u_dd_std.Save(Application.StartupPath + "\data\user\SAA_u_dd_std.xml")
    End Sub

    Public Property modification As Integer
        Get
            modification = modificationvalue
        End Get
        Set(value As Integer)
            modificationvalue = value
        End Set
    End Property

    Public Function getbaseid(ByVal index As Integer) As String
        getbaseid = "0"
        Dim count As Integer = 0
        For a = 0 To ddshiptypegroup.length - 1
            If ddshiptypegroup.getattribute(ddshiptypegroup.getid(a), 2) >= modificationvalue Then
                count = count + 1
                If count = index Then
                    getbaseid = ddshiptypegroup.getattribute(ddshiptypegroup.getid(a), 0)
                    Exit For
                End If
            End If
        Next
    End Function

    Public Function getbaseattribute(ByVal id As String, ByVal attributeindex As Integer)
        getbaseattribute = "0"
        If attributeindex = 1 Then
            getbaseattribute = ddshiptypegroup.getattribute(id, 1)
        ElseIf attributeindex = 2 Then
            getbaseattribute = baseshiptypegroup.getattribute(id, 13)
        End If
    End Function

    ReadOnly Property unit(ByVal uniquecode As String) As ddshiptypeunit_class
        Get
            unit = ddshiptypeunit(ddshiptypelist.Item(uniquecode))
        End Get
    End Property

    Public Function getid(ByVal index As Integer) As String
        getid = "0"
        If ddshiptypeunit.Length > index Then
            getid = ddshiptypeunit(index).uniquecode
        End If
    End Function

    '==========以下函数观察后删除

    Public Function getattribute(ByVal uniquecode As String, ByVal attributeindex As Integer)
        getattribute = 0
        If uniquecode <> "0" Then
            uniquecode = Trim(uniquecode)
            If attributeindex = 1 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).name
            ElseIf attributeindex = 2 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).baseid
            ElseIf attributeindex = 3 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).type
            ElseIf attributeindex = 4 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).gridcount
            ElseIf attributeindex = 5 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).enhanced
            ElseIf attributeindex = 6 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).level
            ElseIf attributeindex = 7 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).luck
            ElseIf attributeindex = 8 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).hp
            ElseIf attributeindex = 9 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).fire
            ElseIf attributeindex = 10 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).torpedo
            ElseIf attributeindex = 11 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).antiaircraft
            ElseIf attributeindex = 12 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).armor
            ElseIf attributeindex = 13 Then   '对潜
                getattribute = Int(ddshiptypeunit(ddshiptypelist(uniquecode)).antisubmarine99 + (ddshiptypeunit(ddshiptypelist(uniquecode)).antisubmarine165 - ddshiptypeunit(ddshiptypelist(uniquecode)).antisubmarine99) / (ddshiptypeunit(ddshiptypelist(uniquecode)).level - 99) * 66)
            ElseIf attributeindex = 14 Then     '回避
                getattribute = Int(ddshiptypeunit(ddshiptypelist(uniquecode)).avoid99 + (ddshiptypeunit(ddshiptypelist(uniquecode)).avoid165 - ddshiptypeunit(ddshiptypelist(uniquecode)).avoid99) / (ddshiptypeunit(ddshiptypelist(uniquecode)).level - 99) * 66)
            ElseIf attributeindex = 15 Then   '索敌
                getattribute = Int(ddshiptypeunit(ddshiptypelist(uniquecode)).spotting99 + (ddshiptypeunit(ddshiptypelist(uniquecode)).spotting165 - ddshiptypeunit(ddshiptypelist(uniquecode)).spotting99) / (ddshiptypeunit(ddshiptypelist(uniquecode)).level - 99) * 66)
            ElseIf attributeindex = 16 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).speed
            ElseIf attributeindex = 17 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).firerange
            ElseIf attributeindex = 18 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).fuel
            ElseIf attributeindex = 19 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).ammunition
            ElseIf attributeindex = 20 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).modification
            ElseIf attributeindex = 21 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).extraarmour
            ElseIf attributeindex = 22 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).dropship
            ElseIf attributeindex = 23 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).tank
            ElseIf attributeindex = 24 Then
                getattribute = ddshiptypeunit(ddshiptypelist(uniquecode)).headquarters
            End If
        End If
    End Function


End Class



Public Class sgUIcontrol_class
    Dim list1string As New Collection

    Public Function checklist1refresh(ByVal index As Integer) As String
        checklist1refresh = userddshiptype.getattribute(userddshiptype.getid(index), 1)
    End Function

    Private Sub list1refresh()
        list1string.Clear()
        If ddshipgroup.length > 0 Then
            For a = 0 To ddshipgroup.length - 1
                list1string.Add(ddshipgroup.ship(a).showstring)
                If ddshipgroup.ship(a).length > 0 Then
                    For b = 0 To ddshipgroup.ship(a).length - 1
                        list1string.Add(ddshipgroup.ship(a).carry(b).showstring)
                    Next
                End If
            Next
        End If
    End Sub

    Public Function showlist1(ByVal index As Integer) As String
        showlist1 = ""
        If index = 1 Then Call list1refresh()
        If index <= list1string.Count Then
            showlist1 = list1string(index)
        End If
    End Function

    Public Function list1clickid(ByVal index As Integer) As String()
        list1clickid = {"-1", "-1"}
        Dim count As Integer = -1
        If ddshipgroup.length > 0 Then
            For a = 0 To ddshipgroup.length - 1
                count = count + 1
                If count = index Then
                    list1clickid = {a, "-1"}
                End If
                For b = 0 To ddshipgroup.ship(a).length - 1
                    count = count + 1
                    If count = index Then
                        list1clickid = {a, b}
                    End If
                Next
            Next
        End If
    End Function

    Public Sub changeshipnfset(ByVal shipindex As Integer, ByVal nfmode As Integer, ByVal restrictmode As Integer, ByVal enhancedlock As Integer, ByVal buffmultiple As Double)
        ddshipgroup.ship(shipindex).nfmode = nfmode
        ddshipgroup.ship(shipindex).restrictmode = restrictmode
        ddshipgroup.ship(shipindex).enhancedlock = enhancedlock
        ddshipgroup.ship(shipindex).buffmultiple = buffmultiple
    End Sub

    Public Sub changesortmode(ByVal mode As Integer)
        If mode = 0 Then
            userddshiptype.unitsort("level")
        ElseIf mode = 1 Then
            userddshiptype.unitsort("damage")
        ElseIf mode = 2 Then
            userddshiptype.unitsort("type")
        End If
    End Sub
End Class











