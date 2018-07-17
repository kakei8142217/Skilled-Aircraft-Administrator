Module SAA_module
    '============整体

    Public saa_plane As SAA_plane_form
    Public saa_ddrank As SAA_ddrank_form
    Public setting_mode As Integer

    '==============公共数据

    Public baseshiptypegroup As basedatagroup_class
    Public baseequipmentgroup As basedatagroup_class

    Public cvequipmentgroup As basedatagroup_class
    Public ddequipmentgroup As basedatagroup_class

    Public saa_u_pl_emd As New Xml.XmlDocument
    Public usercvequipmentgroup As basedatagroup_class

    Public saa_u_sg_emd As New Xml.XmlDocument
    Public userddequipmentgroup As basedatagroup_class

    Public saa_dd_std As New Xml.XmlDocument
    Public ddshiptypegroup As basedatagroup_class

    Public saa_u_dd_std As New Xml.XmlDocument
    Public userddshiptypegroup As basedatagroup_class


    Public ddextrabuffgroup As basedatagroup_class



    Public mapgroup As basedatagroup_class

    Public filecontrol As filecontrol_class

    '==============舰载机装配


    Public plane As New plane_class

    Public cvshiptype As cvshiptype_class

    Public cvship(19) As cvship_class
    Public vessel As New carry_class

    Public restrictcontrol As restrictcontrol_class

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

    Public plUIcontrol As New plUIcontrol_class


    '==============DD火雷排行

    Public userddshiptype As userddshiptype_class

    Public smallgun As smallgun_class

    Public sgUIcontrol As sgUIcontrol_class

    Public ddshipgroup As ddshipgroup_class



    '===========






End Module

Public Class filecontrol_class
    Dim saa_b_std As New Xml.XmlDocument
    Dim saa_b_emd As New Xml.XmlDocument
    Dim saa_pl_emd As New Xml.XmlDocument
    Dim saa_sg_emd As New Xml.XmlDocument
    Dim saa_dd_ebd As New Xml.XmlDocument

    Dim saa_md As New Xml.XmlDocument

    Sub New()
        Call checkfilexist()
        Call checkoldversionfile()

        Call loadfile("\base\SAA_b_std.xml", saa_b_std, baseshiptypegroup)
        Call loadfile("\base\SAA_b_emd.xml", saa_b_emd, baseequipmentgroup)
        Call loadfile("\base\SAA_pl_emd.xml", saa_pl_emd, cvequipmentgroup)
        Call loadfile("\base\SAA_sg_emd.xml", saa_sg_emd, ddequipmentgroup)

        Call loadfile("\user\SAA_u_pl_emd.xml", saa_u_pl_emd, usercvequipmentgroup)
        Call loadfile("\user\SAA_u_sg_emd.xml", saa_u_sg_emd, userddequipmentgroup)

        Call loadfile("\base\SAA_dd_std.xml", saa_dd_std, ddshiptypegroup)
        Call loadfile("\user\SAA_u_dd_std.xml", saa_u_dd_std, userddshiptypegroup)

        Call loadfile("\base\SAA_dd_ebd.xml", saa_dd_ebd, ddextrabuffgroup)

        Call loadfile("\base\SAA_md.xml", saa_md, mapgroup)

    End Sub

    Public Overloads Sub loadfile(ByVal filename As String, ByRef xmldoc As Xml.XmlDocument, ByRef newbasedategroup As basedatagroup_class)
        xmldoc.Load(Application.StartupPath + "\data" & filename)
        Dim node As Xml.XmlNode = xmldoc.DocumentElement
        loaddata(node, newbasedategroup)
    End Sub

    Private Sub loaddata(ByVal fathernode As Xml.XmlNode, ByRef newbasedategroup As basedatagroup_class)
        If fathernode.ChildNodes(0) IsNot Nothing Then
            Dim length As Integer = fathernode.ChildNodes(0).Attributes.Count
            Dim lengthchilds As Integer = fathernode.ChildNodes.Count
            newbasedategroup = New basedatagroup_class(length)
            Dim newbasedata As basedata_class
            For Each node As Xml.XmlNode In fathernode.ChildNodes
                newbasedata = New basedata_class(length)
                For a = 0 To length - 1
                    newbasedata.setattribute(a, node.Attributes(a).Value)
                Next
                newbasedategroup.setdata(newbasedata)
                loaddata(node, newbasedata.getdatagroup)
            Next
        End If
    End Sub

    Private Sub checkfilexist()
        If Not IO.Directory.Exists(Application.StartupPath + "\data\user") Then
            IO.Directory.CreateDirectory(Application.StartupPath + "\data\user")
        End If
        If Not IO.File.Exists(Application.StartupPath + "\data\user\SAA_u_pl_emd.xml") Then
            Dim xmlstring As New System.Text.StringBuilder
            xmlstring.Append("<?xml version=""1.0"" encoding=""gb2312""?>")
            xmlstring.Append("<userplaneequipmentdata>")
            xmlstring.Append("</userplaneequipmentdata>")
            saa_u_pl_emd.LoadXml(xmlstring.ToString)
            'saa_u_pl_emd.DocumentElement.SetAttribute("version", "00000000")
            saa_u_pl_emd.Save(Application.StartupPath + "\data\user\SAA_u_pl_emd.xml")
        End If
        If Not IO.File.Exists(Application.StartupPath + "\data\user\SAA_u_sg_emd.xml") Then
            Dim xmlstring As New System.Text.StringBuilder
            xmlstring.Append("<?xml version=""1.0"" encoding=""gb2312""?>")
            xmlstring.Append("<usersmallgunequipmentdata>")
            xmlstring.Append("</usersmallgunequipmentdata>")
            saa_u_pl_emd.LoadXml(xmlstring.ToString)
            ''saa_u_pl_emd.DocumentElement.SetAttribute("version", "00000000")
            saa_u_pl_emd.Save(Application.StartupPath + "\data\user\SAA_u_sg_emd.xml")
        End If
        If Not IO.File.Exists(Application.StartupPath + "\data\user\SAA_u_dd_std.xml") Then
            Dim xmlstring As New System.Text.StringBuilder
            xmlstring.Append("<?xml version=""1.0"" encoding=""gb2312""?>")
            xmlstring.Append("<userddshiptypedata>")
            xmlstring.Append("</userddshiptypedata>")
            saa_u_pl_emd.LoadXml(xmlstring.ToString)
            'saa_u_pl_emd.DocumentElement.SetAttribute("version", "00000000")
            saa_u_pl_emd.Save(Application.StartupPath + "\data\user\SAA_u_dd_std.xml")
        End If
    End Sub

    Private Sub checkoldversionfile()
        Dim result As DialogResult
        For Each file As String In IO.Directory.GetFiles(Application.StartupPath + "\data")
            If Mid(IO.Path.GetFileName(file), 1, 3) = "SAA" And IO.Path.GetExtension(file) = ".xml" Then
                result = MessageBox.Show("检测到旧版本记录文件，是否转换格式并覆写到新版本？（此操作将清除新版本记录文件中的所有内容）", "警告", MessageBoxButtons.YesNo)
                Exit For
            End If
        Next
        If result = DialogResult.Yes Then
            If IO.File.Exists(Application.StartupPath + "\data\SAAuserplanedata.xml") Then
                Dim oldxml As New Xml.XmlDocument
                Dim newxml As New Xml.XmlDocument
                oldxml.Load(Application.StartupPath + "\data\SAAuserplanedata.xml")
                Dim xmlstring As New System.Text.StringBuilder
                xmlstring.Append("<?xml version=""1.0"" encoding=""gb2312""?>")
                xmlstring.Append("<userplaneequipmentdata>")
                xmlstring.Append("</userplaneequipmentdata>")
                newxml.LoadXml(xmlstring.ToString)
                'saa_u_pl_emd.DocumentElement.SetAttribute("version", "00000000")
                newxml.Save(Application.StartupPath + "\data\user\SAA_u_pl_emd.xml")



                For Each node As Xml.XmlNode In oldxml.DocumentElement.ChildNodes
                    Dim newnode As Xml.XmlElement = newxml.CreateElement("plane")
                    Dim id As String = Trim(Str(node.Attributes("id").Value) & Format(node.Attributes("improve").Value, "00"))
                    newnode.SetAttribute("id", id)
                    newnode.SetAttribute("name", node.Attributes("name").Value)
                    newnode.SetAttribute("improve", node.Attributes("improve").Value)
                    newnode.SetAttribute("amountlist1", node.Attributes("amountlist1").Value)
                    newnode.SetAttribute("amountlist2", node.Attributes("amountlist2").Value)
                    newnode.SetAttribute("antiaircraft", node.Attributes("antiaircraft").Value)
                    newxml.DocumentElement.AppendChild(newnode)
                Next
                newxml.Save(Application.StartupPath + "\data\user\SAA_u_pl_emd.xml")
            End If

            For Each file As String In IO.Directory.GetFiles(Application.StartupPath + "\data")
                If Mid(IO.Path.GetFileName(file), 1, 3) = "SAA" And IO.Path.GetExtension(file) = ".xml" Then
                    IO.File.Delete(file)
                End If
            Next

        End If
    End Sub
End Class

Public Class basedata_class
    Dim datavalue()
    Dim datagroup As basedatagroup_class

    Sub New(ByVal newdatalength As Integer)
        ReDim datavalue(newdatalength - 1)
    End Sub

    Public Sub setattribute(ByVal index, ByVal value)
        datavalue(index) = value
    End Sub

    Public Function getattribute(ByVal index As Integer)
        getattribute = -9999
        If index < datavalue.Length Then
            getattribute = datavalue(index)
        End If
    End Function

    ReadOnly Property datalength As Integer
        Get
            Try
                datalength = datavalue.Length
            Catch
                datalength = 0
            End Try
        End Get
    End Property

    Public Property getdatagroup As basedatagroup_class
        Get
            getdatagroup = datagroup
        End Get
        Set(value As basedatagroup_class)
            datagroup = value
        End Set
    End Property

End Class

Public Class basedatagroup_class
    Dim basedata() As basedata_class
    Dim list As New Collection
    Dim datalengthvalue As Integer

    Sub New(ByVal newdatalenth As Integer)
        datalengthvalue = newdatalenth
    End Sub

    Public Function exist(ByVal id As String) As Boolean
        exist = list.Contains(Trim(Str(id)))
    End Function

    'Public Function getdata(ByVal id As String) As basedata_class
    '    Dim target As New basedata_class(datalengthvalue)
    '    copy(basedata(list(Trim(Str(id)))), target)
    '    getdata = target
    'End Function

    Public Property getdata(ByVal id As String) As basedata_class
        Get
            getdata = basedata(list(Trim(id)))
        End Get
        Set(value As basedata_class)

        End Set
    End Property

    'Public Sub setdata(ByVal from As basedata_class)
    '    If list.Contains(from.getattribute(0)) Then
    '        Call copy(from, basedata(list(Trim(Str(from.getattribute(0))))))
    '    Else
    '        ReDim Preserve basedata(length)
    '        basedata(length - 1) = New basedata_class(datalengthvalue)
    '        Call copy(from, basedata(length - 1), length - 1)
    '    End If
    'End Sub

    Public Sub setdata(ByVal from As basedata_class)
        If list.Contains(from.getattribute(0)) Then
            basedata(list(Trim(Str(from.getattribute(0))))) = from
        Else
            ReDim Preserve basedata(length)
            basedata(length - 1) = from
            list.Add(length - 1, basedata(length - 1).getattribute(0))
        End If
    End Sub

    Public Sub removedata(ByVal from As basedata_class)
        If list.Contains(from.getattribute(0)) Then
            If list(Trim(Str(from.getattribute(0)))) < length - 1 Then
                For a = list(Trim(Str(from.getattribute(0)))) + 1 To length - 1
                    basedata(a - 1) = basedata(a)
                Next
            End If
            ReDim Preserve basedata(length - 2)
            list.Clear()
            For a = 0 To length - 1
                list.Add(a, basedata(a).getattribute(0))
            Next
        End If
    End Sub

    Public Function getattribute(ByVal id As String, ByVal attributeindex As Integer)
        getattribute = 0
        If list.Contains(id) Then
            Dim indexstring As String = Trim(id)
            getattribute = basedata(list(indexstring)).getattribute(attributeindex)
        End If
    End Function

    Public Function getid(ByVal index As Integer) As String
        getid = "0"
        If index < basedata.Length Then
            getid = basedata(index).getattribute(0)
        End If
    End Function

    ReadOnly Property datalength As Integer
        Get
            datalength = datalengthvalue
        End Get
    End Property

    ReadOnly Property length As Integer
        Get
            Try
                length = basedata.Length
            Catch
                length = 0
            End Try
        End Get
    End Property

    Public Sub clear()
        ReDim basedata(-1)
        list.Clear()
    End Sub

    Private Sub copy(ByVal from As basedata_class, ByVal target As basedata_class, Optional ByVal index As Integer = -1)
        If from.datalength <= target.datalength Then
            If index <> -1 Then
                If target.getattribute(0) IsNot Nothing Then
                    list.Remove(Trim(target.getattribute(0)))
                    If list.Contains(Trim(from.getattribute(0))) Then
                        list.Remove(Trim(from.getattribute(0)))
                    End If
                End If
                list.Add(index, Trim(from.getattribute(0)))
            End If
            For a = 0 To from.datalength - 1
                target.setattribute(a, from.getattribute(a))
            Next
        End If
    End Sub
End Class

Public MustInherit Class equipmentunit_class
    Dim uniquecodevalue As String
    Dim baseidvalue As String
    Dim extendidvalue As String

    Sub New(ByVal newuniquecode As String)
        uniquecodevalue = newuniquecode
        baseidvalue = Int(Val(uniquecodevalue) / 10000)
        extendidvalue = Int(Val(uniquecodevalue) / 100)
    End Sub

    ReadOnly Property extendid As String
        Get
            extendid = extendidvalue
        End Get
    End Property

    ReadOnly Property baseid As String
        Get
            baseid = baseidvalue
        End Get
    End Property

    ReadOnly Property uniquecode As String
        Get
            uniquecode = uniquecodevalue
        End Get
    End Property

    Overridable ReadOnly Property fire As Double
        Get
            fire = baseequipmentgroup.getattribute(baseidvalue, 2)
        End Get
    End Property

    Overridable ReadOnly Property torpedo As Double
        Get
            torpedo = baseequipmentgroup.getattribute(baseidvalue, 3)
        End Get
    End Property

    Overridable ReadOnly Property antiaircraft As Double
        Get
            antiaircraft = baseequipmentgroup.getattribute(baseidvalue, 4)
        End Get
    End Property

    Overridable ReadOnly Property armor As Double
        Get
            armor = baseequipmentgroup.getattribute(baseidvalue, 5)
        End Get
    End Property

    Overridable ReadOnly Property antisubmarine As Double
        Get
            antisubmarine = baseequipmentgroup.getattribute(baseidvalue, 6)
        End Get
    End Property

    Overridable ReadOnly Property avoid As Double
        Get
            avoid = baseequipmentgroup.getattribute(baseidvalue, 7)
        End Get
    End Property

    Overridable ReadOnly Property spotting As Double
        Get
            spotting = baseequipmentgroup.getattribute(baseidvalue, 8)
        End Get
    End Property

    Overridable ReadOnly Property luck As Double
        Get
            luck = baseequipmentgroup.getattribute(baseidvalue, 9)
        End Get
    End Property

    Overridable ReadOnly Property hit As Double
        Get
            hit = baseequipmentgroup.getattribute(baseidvalue, 10)
        End Get
    End Property

    Overridable ReadOnly Property bombing As Double
        Get
            bombing = baseequipmentgroup.getattribute(baseidvalue, 11)
        End Get
    End Property

    Overridable ReadOnly Property firerange As Double
        Get
            firerange = baseequipmentgroup.getattribute(baseidvalue, 12)
        End Get
    End Property

    Overridable ReadOnly Property airrange As Double
        Get
            airrange = baseequipmentgroup.getattribute(baseidvalue, 13)
        End Get
    End Property

    Overridable ReadOnly Property setupal As Double
        Get
            setupal = baseequipmentgroup.getattribute(baseidvalue, 14)
        End Get
    End Property

End Class

Public MustInherit Class shiptypeunit_class
    Dim uniquecodevalue As String
    Dim baseidvalue As String
    Dim extendidvalue As String

    Sub New(ByVal newuniquecode As String)
        uniquecodevalue = newuniquecode
        baseidvalue = Int(Val(uniquecodevalue) / 100000000)
        extendidvalue = Int(Val(uniquecodevalue) / 10)
    End Sub

    ReadOnly Property extendid As String
        Get
            extendid = extendidvalue
        End Get
    End Property

    ReadOnly Property baseid As String
        Get
            baseid = baseidvalue
        End Get
    End Property

    ReadOnly Property uniquecode As String
        Get
            uniquecode = uniquecodevalue
        End Get
    End Property

    Overridable ReadOnly Property hp As Double
        Get
            hp = baseshiptypegroup.getattribute(baseidvalue, 2)
        End Get
    End Property

    Overridable ReadOnly Property fire As Double
        Get
            fire = baseshiptypegroup.getattribute(baseidvalue, 3)
        End Get
    End Property

    Overridable ReadOnly Property torpedo As Double
        Get
            torpedo = baseshiptypegroup.getattribute(baseidvalue, 4)
        End Get
    End Property

    Overridable ReadOnly Property antiaircraft As Double
        Get
            antiaircraft = baseshiptypegroup.getattribute(baseidvalue, 5)
        End Get
    End Property

    Overridable ReadOnly Property armor As Double
        Get
            armor = baseshiptypegroup.getattribute(baseidvalue, 6)
        End Get
    End Property

    Overridable ReadOnly Property antisubmarine99 As Double
        Get
            antisubmarine99 = baseshiptypegroup.getattribute(baseidvalue, 7)
        End Get
    End Property

    Overridable ReadOnly Property antisubmarine165 As Double
        Get
            antisubmarine165 = baseshiptypegroup.getattribute(baseidvalue, 8)
        End Get
    End Property

    Overridable ReadOnly Property avoid99 As Double
        Get
            avoid99 = baseshiptypegroup.getattribute(baseidvalue, 9)
        End Get
    End Property

    Overridable ReadOnly Property avoid165 As Double
        Get
            avoid165 = baseshiptypegroup.getattribute(baseidvalue, 10)
        End Get
    End Property

    Overridable ReadOnly Property spotting99 As Double
        Get
            spotting99 = baseshiptypegroup.getattribute(baseidvalue, 11)
        End Get
    End Property

    Overridable ReadOnly Property spotting165 As Double
        Get
            spotting165 = baseshiptypegroup.getattribute(baseidvalue, 12)
        End Get
    End Property

    Overridable ReadOnly Property initialluck As Double
        Get
            initialluck = baseshiptypegroup.getattribute(baseidvalue, 13)
        End Get
    End Property

    Overridable ReadOnly Property speed As Double
        Get
            speed = baseshiptypegroup.getattribute(baseidvalue, 14)
        End Get
    End Property

    Overridable ReadOnly Property firerange As Double
        Get
            firerange = baseshiptypegroup.getattribute(baseidvalue, 15)
        End Get
    End Property

    Overridable ReadOnly Property gridcount As Double
        Get
            gridcount = baseshiptypegroup.getattribute(baseidvalue, 16)
        End Get
    End Property

    Overridable ReadOnly Property fuel As Double
        Get
            fuel = baseshiptypegroup.getattribute(baseidvalue, 17)
        End Get
    End Property

    Overridable ReadOnly Property ammunition As Double
        Get
            ammunition = baseshiptypegroup.getattribute(baseidvalue, 18)
        End Get
    End Property

End Class