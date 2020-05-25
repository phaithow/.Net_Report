
Imports System.Net
Imports System.Net.Mail
Imports System.Web
Imports System.Xml
Imports System.Globalization
Imports System.Threading
Imports System.IO

Public Class PLMI0050

    Dim dtAll As DataTable
    Dim dtIng As DataTable
    Dim dtPack As DataTable
    Dim dtComplete As DataTable
    Dim dtDistIng As DataTable
    Dim dtDistPK As DataTable

    Dim TextFromPath As String
    Dim TextToPath As String
    Dim toAddAdmin As String
    Dim fromAddAdmin As String
    Dim industIng As String
    Dim industPK As String
    Dim smtpIPServer As String
    Dim culName As String = Thread.CurrentThread.CurrentCulture.Name

    Dim PKID_COMPANY As String
    Dim pNUMBER As String
    Dim pFNUMBER As String
    Dim PKID_USERS As String
    Dim pPHONE As String
    Dim PKID_COMPANYNAME As String
    Dim pNAME As String
    Dim PKID_COMPANYADDRESS As String
    Dim pCITY As String
    Dim pCOUNTRY As String
    Dim PKID_COUNTRY As String
    Dim pSTATEORPROVINCE As String
    Dim pPOSTALCODE As String
    Dim pSTREET1 As String
    Dim pSTREET2 As String
    Dim COMPANY_SpecLegacySpecJoin As String
    Dim PKID_SPECLEGACYPROFILE As String
    Dim pSAPCODE As String
    Dim COMPANY_scrmBUsRelationship As String
    Dim PKID_ENTITYSTATUS As String
    Dim PKID_BUSINESSUNIT As String
    Dim PKID_Facility As String
    Dim PKID_FacilityName As String
    Dim PKID_FacilityAddress As String
    Dim FACILITY_scrmBUsRelationship As String
    Dim COMPANYCREATEDDATE As String

    Dim File_Error As String = "PLMI0050.LOG"

    Private Sub PLMI0050_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            If culName <> "en-GB" Then
                System.Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo("en-GB")
            End If

            Dim clsDocument As New System.Xml.XmlDocument
            Dim path As String = Application.StartupPath

            clsDocument.Load(path & "\CPFSetting.config")
            Dim clsNode As System.Xml.XmlNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/PLMDBConnector/add")
            Dim PLMDBKey As String = clsNode.Attributes("key").Value
            Dim PLMDBvalue As String = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/CpfPLMDBConnector/add")
            Dim CpfPLMDBKey As String = clsNode.Attributes("key").Value
            Dim CpfPLMDBvalue As String = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/TextFromFolder/add")
            TextFromPath = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/TextToFolder/add")
            TextToPath = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/FromPLMAdminTeam/add")
            fromAddAdmin = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/ToPLMAdminTeam/add")
            toAddAdmin = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/IndustryCodeIngredient/add")
            industIng = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/IndustryCodePackaging/add")
            industPK = clsNode.Attributes("value").Value

            clsNode = clsDocument.SelectSingleNode("//CPFInterfaceSetting/smtpIPServer/add")
            smtpIPServer = clsNode.Attributes("value").Value

            Dim aCpfDB() As String = CpfPLMDBvalue.Split(",")
            Dim aPLMDB() As String = PLMDBvalue.Split(",")

            DatabaseBrand = "ORACLE"
            UserSchema1 = aPLMDB(0)
            DatabaseName1 = aPLMDB(1)
            DatabaseHost1 = aPLMDB(1)

            UserSchema2 = aCpfDB(0)
            DatabaseName2 = aCpfDB(1)
            DatabaseHost2 = aCpfDB(1)

            dtIng = New DataTable
            CreateSchema(dtIng)

            dtPack = New DataTable
            CreateSchema(dtPack)

            dtComplete = New DataTable
            dtComplete.Columns.Add("STATUS")
            dtComplete.Columns.Add("SAP_CODE")
            dtComplete.Columns.Add("NAME")
            dtComplete.Columns.Add("COMPANY_NUMBER")
            dtComplete.Columns.Add("FACILITY_NUMBER")

            Dim fileEntries As String() = Directory.GetFiles(TextFromPath, "*.txt")

            Dim fileName As String
            For Each fileName In fileEntries

                If Read_TextFile(fileName) = False Then

                    'Send Mail Error
                    SendMailToPLMAdmin("Not Complete : " & fileName)

                End If
            Next

            'Interface ข้อมูล Ingredient
            If dtIng.Rows.Count > 0 Then
                'Distinct ข้อมูล
                dtDistIng = dtIng.DefaultView.ToTable(True)
                If InterfaceData(dtDistIng, "I") = False Then

                    Dim WriteFile As StreamWriter
                    WriteFile = File.AppendText(File_Error)
                    WriteFile.WriteLine("Interface not Complete!! " & Now & " ")
                    WriteFile.Flush()
                    WriteFile.Close()

                    SendMailToPLMAdmin("Interface not Complete!!")

                    Me.Dispose()
                    Me.Close()

                End If
            End If

            'Interface ข้อมูล Ingredient
            If dtPack.Rows.Count > 0 Then
                'Distinct ข้อมูล
                dtDistPK = dtPack.DefaultView.ToTable(True)
                If InterfaceData(dtDistPK, "P") = False Then

                    Dim WriteFile As StreamWriter
                    WriteFile = File.AppendText(File_Error)
                    WriteFile.WriteLine("Interface not Complete!! " & Now & " ")
                    WriteFile.Flush()
                    WriteFile.Close()

                    SendMailToPLMAdmin("Interface not Complete!!")

                    Me.Dispose()
                    Me.Close()

                End If
            End If

            'Copy File
            For Each fileName In fileEntries

                IO.File.Copy(fileName, TextToPath & IO.Path.GetFileName(fileName), True)
                File.Delete(fileName)

            Next

            System.Threading.Thread.CurrentThread.CurrentCulture = New CultureInfo(culName)

            Me.Dispose()
            Me.Close()

        Catch ex As Exception

            Dim WriteFile As StreamWriter
            WriteFile = File.AppendText(File_Error)
            WriteFile.WriteLine("" & ex.Message.ToString & " " & Now & " ")
            WriteFile.Flush()
            WriteFile.Close()

            SendMailToPLMAdmin(ex.Message.ToString)

            Me.Dispose()
            Me.Close()
        End Try
    End Sub

    Private Function InterfaceData(ByVal dtInterface As DataTable, ByVal Flag As String) As Boolean

        Try

            Dim fileEntries As String() = Directory.GetFiles(TextFromPath, "*.txt")

            'กรณีที่บันทึกข้อมูลสำเร็จ
            If SaveData(dtInterface) = True Then

                Dim CompleteMsg As String = ""
                Dim companyAdd As String = ""
                Dim companyUpdate As String = ""
                Dim FacilityAdd As String = ""
                Dim FacilityUpdate As String = ""

                If dtComplete.Rows.Count > 0 Then
                    '1. Send Mail Complete
                    'Create mail message

                    For Each dr As DataRow In dtComplete.Rows

                        If dr("STATUS") = "A" Then

                            companyAdd = companyAdd & dr("SAP_CODE") & vbTab & dr("COMPANY_NUMBER") & "  " & dr("NAME") & vbNewLine

                            If dr("FACILITY_NUMBER") <> "" Then
                                FacilityAdd = FacilityAdd & "  " & vbTab & dr("FACILITY_NUMBER") & "  " & dr("NAME") & vbNewLine
                            End If


                        ElseIf dr("STATUS") = "U" Then

                            companyUpdate = companyUpdate & dr("SAP_CODE") & vbTab & dr("COMPANY_NUMBER") & "  " & dr("NAME") & vbNewLine

                            If dr("FACILITY_NUMBER") <> "" Then
                                FacilityUpdate = FacilityUpdate & "  " & vbTab & dr("FACILITY_NUMBER") & "  " & dr("NAME") & vbNewLine
                            End If

                        End If

                    Next

                    'Read Email Address
                    Dim dtAddress As New DataTable
                    If Flag = "I" Then

                        dtAddress = ReadMailAddressIng()

                    ElseIf Flag = "P" Then
                        dtAddress = ReadMailAddressPK()
                    End If

                    'Send Mail

                    Dim mailSubject As String
                    If Flag = "I" Then
                        mailSubject = "Vendor Master Interface in PLM System is Success(Ingredient)"
                    Else
                        mailSubject = "Vendor Master Interface in PLM System is Success(Packaging)"
                    End If

                    If dtAddress.Rows.Count > 0 Then
                        For Each rr As DataRow In dtAddress.Rows

                            Dim mailMsg As String = CompleteMailMessage(rr("NAME"), companyAdd, companyUpdate, FacilityAdd, FacilityUpdate, Flag)
                            SendMail(rr("EMAIL"), fromAddAdmin, mailSubject, mailMsg)

                        Next

                    End If

                End If
            End If

            Return True

        Catch ex As Exception
            Dim WriteFile As StreamWriter
            WriteFile = File.AppendText(File_Error)
            WriteFile.WriteLine("" & ex.Message.ToString & " " & Now & " ")
            WriteFile.Flush()
            WriteFile.Close()

            SendMailToPLMAdmin(ex.Message.ToString)

            Me.Dispose()
            Me.Close()
        End Try

    End Function

    Private Sub ClearData()
        PKID_COMPANY = ""
        pNUMBER = ""
        pFNUMBER = ""
        PKID_USERS = ""
        pPHONE = ""
        PKID_COMPANYNAME = ""
        pNAME = ""
        PKID_COMPANYADDRESS = ""
        pCITY = ""
        pCOUNTRY = ""
        PKID_COUNTRY = ""
        pSTATEORPROVINCE = ""
        pPOSTALCODE = ""
        pSTREET1 = ""
        pSTREET2 = ""
        COMPANY_SpecLegacySpecJoin = ""
        PKID_SPECLEGACYPROFILE = ""
        pSAPCODE = ""
        COMPANY_scrmBUsRelationship = ""
        PKID_ENTITYSTATUS = ""
        PKID_BUSINESSUNIT = ""
        PKID_Facility = ""
        PKID_FacilityName = ""
        PKID_FacilityAddress = ""
        FACILITY_scrmBUsRelationship = ""
        COMPANYCREATEDDATE = ""
    End Sub
    Private Function SaveData(ByVal dt As DataTable) As Boolean
        Dim Comm As OleDb.OleDbCommand
        Dim DataTran As OleDb.OleDbTransaction
        Dim status As String = ""
        Dim dr As DataRow

        Try

            dtComplete.Rows.Clear()

            dbConnection = New dbConnector.dbSelector.dbConn
            dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
            dbConnection.Open()
            dbConnection.set_Command(DataTran, Comm)

            For Each row As DataRow In dt.Rows

                If Not IsDBNull(row("PHONE")) Then
                    pPHONE = row("PHONE")
                Else
                    pPHONE = ""
                End If

                If Not IsDBNull(row("NAME")) Then
                    pNAME = row("NAME")
                Else
                    pNAME = ""
                End If

                If Not IsDBNull(row("CITY")) Then
                    pCITY = row("CITY")
                Else
                    pCITY = ""
                End If

                If Not IsDBNull(row("COUNTRY")) Then
                    pCOUNTRY = row("COUNTRY")
                Else
                    pCOUNTRY = ""
                End If

                If Not IsDBNull(row("STATE_OR_PROVINCE")) Then
                    pSTATEORPROVINCE = row("STATE_OR_PROVINCE")
                Else
                    pSTATEORPROVINCE = ""
                End If

                If Not IsDBNull(row("POSTAL_CODE")) Then
                    pPOSTALCODE = row("POSTAL_CODE")
                Else
                    pPOSTALCODE = ""
                End If

                If Not IsDBNull(row("STREET1")) Then
                    pSTREET1 = row("STREET1")
                Else
                    pSTREET1 = ""
                End If

                If Not IsDBNull(row("STREET2")) Then
                    pSTREET2 = row("STREET2")
                Else
                    pSTREET2 = ""
                End If

                If Not IsDBNull(row("SAP_CODE")) Then
                    pSAPCODE = row("SAP_CODE")
                Else
                    pSAPCODE = ""
                End If


                'Get ExistData
                If GetExistCompanyData(row("SAP_CODE")) = True Then
                    'Update Old Data
                    status = "U"

                    '***********************************
                    '********* Update Company **********
                    '***********************************

                    'Update ค่าใส่ Table scrmCompany
                    Comm.CommandText = "UPDATE SCRMCOMPANY SET LASTEDITDT = SYSDATE ,PHONE = '" & pPHONE & "' WHERE PKID = '" & PKID_COMPANY & "'"
                    Comm.ExecuteNonQuery()

                    'Update ค่าใส่ Table scrmEntityFreeTextName
                    Comm.CommandText = "UPDATE SCRMENTITYFREETEXTNAME SET NAME = '" & pNAME & "' Where PKID = '" & PKID_COMPANYNAME & "'"
                    Comm.ExecuteNonQuery()

                    'Update ค่าใส่ Table scrmAddress
                    Comm.CommandText = "UPDATE SCRMADDRESS SET CITY = '" & pCITY & "',FKCOUNTRY = (Select distinct PKID  From Countries Where isocode = '" & pCOUNTRY & "') ,STATEORPROVINCE = '" & pSTATEORPROVINCE & "' ," & _
                                       "POSTALCODE = '" & pPOSTALCODE & "' ,STREET1 = '" & pSTREET1 & "',STREET2 = '" & pSTREET2 & "' WHERE PKID = '" & PKID_COMPANYADDRESS & "'"
                    Comm.ExecuteNonQuery()

                    If GetExistFacilityData() = True Then

                        'Update ค่าใส่ Table scrmFacility 
                        Comm.CommandText = "UPDATE SCRMFACILITY SET LASTEDITDT = SYSDATE ,PHONE = '" & pPHONE & "' WHERE PKID = '" & PKID_Facility & "'"
                        Comm.ExecuteNonQuery()

                        'Update ค่าใส่ Table scrmEntityFreeTextName 
                        Comm.CommandText = "UPDATE SCRMENTITYFREETEXTNAME SET NAME = '" & pNAME & "' Where PKID = '" & PKID_FacilityName & "'"
                        Comm.ExecuteNonQuery()

                        'Update ค่าใส่ Table scrmAddress
                        Comm.CommandText = "UPDATE SCRMADDRESS SET CITY = '" & pCITY & "' ,FKCOUNTRY = (Select distinct PKID  From Countries Where isocode = '" & pCOUNTRY & "') ,STATEORPROVINCE = '" & pSTATEORPROVINCE & "' ," & _
                                           "POSTALCODE = '" & pPOSTALCODE & "' ,STREET1 = '" & pSTREET1 & "' ,STREET2 = '" & pSTREET2 & "' WHERE PKID = '" & PKID_FacilityAddress & "'"
                        Comm.ExecuteNonQuery()

                    End If

                Else
                    'Insert New Data

                    status = "A"

                    If GetPKID(row("COUNTRY")) = True Then

                        'Get ค่า Company Number --> pNUMBER
                        pNUMBER = GetCompanyNumber()

                        If pNAME = "" Then
                            Comm.Transaction.Rollback()
                            SendMailToPLMAdmin("Can not load Company Number")
                            Return False
                        End If

                        'Get ค่า Facility Number --> pFNUMBER
                        pFNUMBER = GetFacilityNumber()

                        If pFNUMBER = "" Then
                            Comm.Transaction.Rollback()
                            SendMailToPLMAdmin("Can not load Facility Number")
                            Return False
                        End If

                        '********************************
                        '********* สร้าง Company **********
                        '********************************
                        'Insert ค่าใน Table scrmCompany 
                        Comm.CommandText = "INSERT INTO SCRMCOMPANY(PKID,NUM ,LASTEDITDT,FKCOUNTRYID,FKSTATUS,WEBSITE,CREATIONDATE,FKORIGINATOR,FAX,PHONE) " & _
                                           "values('" & PKID_COMPANY & "','" & pNUMBER & "',SYSDATE,NULL,NULL,NULL,SYSDATE,'" & PKID_USERS & "',NULL,'" & pPHONE & "')"
                        Comm.ExecuteNonQuery()


                        'Insert ค่าใส่ Table scrmEntityFreeTextName
                        Comm.CommandText = "INSERT INTO SCRMENTITYFREETEXTNAME(PKID,LANGID ,NAME,FKENTITY) " & _
                                           "values('" & PKID_COMPANYNAME & "',0,'" & pNAME & "','" & PKID_COMPANY & "')"
                        Comm.ExecuteNonQuery()

                        'Insert ค่าใส่ Table scrmAddress
                        Comm.CommandText = "INSERT INTO SCRMADDRESS(PKID,CITY,FKCOUNTRY,STATEORPROVINCE,POSTALCODE,FKPARENT,STREET1,STREET2,FKPOSTALCOUNTRY ,POSTALCITY,POSTALCODE2,POSTAL1 ,POSTALSTATEORPROVINCE ,POSTAL2) " & _
                                           "values('" & PKID_COMPANYADDRESS & "','" & pCITY & "','" & PKID_COUNTRY & "','" & pSTATEORPROVINCE & "','" & pPOSTALCODE & "','" & PKID_COMPANY & "','" & pSTREET1 & "','" & pSTREET2 & "',NULL,NULL,NULL,NULL,NULL,NULL)"
                        Comm.ExecuteNonQuery()

                        'Insert ค่าใส่ Table specLegacySpecJoin
                        Comm.CommandText = "INSERT INTO SPECLEGACYSPECJOIN(PKID,FKSPECID,FKLEGACYPROFILEID,EQUIVALENT,EXTMANAGED) " & _
                                           "values('" & COMPANY_SpecLegacySpecJoin & "','" & PKID_COMPANY & "','" & PKID_SPECLEGACYPROFILE & "','" & pSAPCODE & "','0')"
                        Comm.ExecuteNonQuery()

                        'Insert ค่าใส่ Table scrmEntityStatBusRelationship 
                        Comm.CommandText = "INSERT INTO SCRMENTITYSTATBUSRELATIONSHIP ( PKID, FKSTATUS,FKENTITY) " & _
                                           "values('" & COMPANY_scrmBUsRelationship & "','" & PKID_ENTITYSTATUS & "','" & PKID_COMPANY & "')"
                        Comm.ExecuteNonQuery()


                        'Insert ค่าใส่ Table scrmEntityStatusBusRelbuJoin 
                        Comm.CommandText = "INSERT INTO SCRMENTITYSTATUSBUSRELBUJOIN (FKRELATIONSHIP, FKBUSINESSUNIT) " & _
                                           "values('" & COMPANY_scrmBUsRelationship & "','" & PKID_BUSINESSUNIT & "')"
                        Comm.ExecuteNonQuery()


                        '**********************************
                        '********* สร้าง Facility ***********
                        '**********************************

                        'Insert ค่าใส่ Table scrmFacility
                        Comm.CommandText = "INSERT INTO SCRMFACILITY(PKID, FKCOMPANY,FKSTATUS  ,FKCOUNTRY  ,NUM, LASTEDITDT, WEBSITE, CREATIONDATE, FKORIGINATOR, FAX, PHONE) " & _
                                           "values('" & PKID_Facility & "','" & PKID_COMPANY & "',NULL,NULL,'" & pFNUMBER & "',SYSDATE,NULL,SYSDATE,'" & PKID_USERS & "',NULL,'" & pPHONE & "')"
                        Comm.ExecuteNonQuery()

                        'Insert ค่าใส่ Table scrmEntityFreeTextName
                        Comm.CommandText = "INSERT INTO SCRMENTITYFREETEXTNAME(PKID,LANGID,NAME,FKENTITY) " & _
                                           "values('" & PKID_FacilityName & "',0,'" & pNAME & "','" & PKID_Facility & "')"
                        Comm.ExecuteNonQuery()

                        'Insert ค่าใส่ Table scrmAddress
                        Comm.CommandText = "INSERT INTO SCRMADDRESS(PKID,CITY,FKCOUNTRY,STATEORPROVINCE,POSTALCODE,FKPARENT,STREET1,STREET2,FKPOSTALCOUNTRY,POSTALCITY,POSTALCODE2,POSTAL1,POSTALSTATEORPROVINCE,POSTAL2) " & _
                                           "values('" & PKID_FacilityAddress & "','" & pCITY & "','" & PKID_COUNTRY & "','" & pSTATEORPROVINCE & "','" & pPOSTALCODE & "','" & PKID_Facility & "','" & pSTREET1 & "','" & pSTREET2 & "',NULL,NULL,NULL,NULL,NULL,NULL)"
                        Comm.ExecuteNonQuery()


                        'Insert ค่าใส่ Table scrmEntityStatBusRelationship
                        Comm.CommandText = "INSERT INTO SCRMENTITYSTATBUSRELATIONSHIP(PKID, FKSTATUS,FKENTITY) " & _
                                           "values('" & FACILITY_scrmBUsRelationship & "','" & PKID_ENTITYSTATUS & "','" & PKID_Facility & "')"
                        Comm.ExecuteNonQuery()


                        'Insert ค่าใส่ Table scrmEntityStatusBusRelbuJoin
                        Comm.CommandText = "INSERT INTO SCRMENTITYSTATUSBUSRELBUJOIN (FKRELATIONSHIP, FKBUSINESSUNIT) " & _
                                           "values('" & FACILITY_scrmBUsRelationship & "','" & PKID_BUSINESSUNIT & "')"
                        Comm.ExecuteNonQuery()

                    Else
                        'Get PKID ไม่ได้ให้ส่งเมล์แจ้ง Admin
                        Comm.Transaction.Rollback()
                        SendMailToPLMAdmin("Can not load PKID")

                        Return False

                    End If

                End If

                'เก็บค่าข้อมูลสำหรับส่ง Email

                dr = dtComplete.NewRow

                dr("STATUS") = status
                dr("SAP_CODE") = pSAPCODE
                dr("NAME") = pNAME
                dr("COMPANY_NUMBER") = pNUMBER
                If IsNothing(pFNUMBER) Then
                    dr("FACILITY_NUMBER") = ""
                Else
                    dr("FACILITY_NUMBER") = pFNUMBER
                End If

                dtComplete.Rows.Add(dr)

                'เคลียร์ข้อมูลเก่าสำหรับรับค่าใน Loop ต่อไป
                ClearData()
            Next


            Comm.Transaction.Commit()
            'Comm.Transaction.Rollback()

        Catch ex As Exception
            'Error ส่งเมล์แจ้ง Admin
            Comm.Transaction.Rollback()

            SendMailToPLMAdmin(ex.Message.ToString)
            Return False

        Finally
            dbConnection.Close()
        End Try

        Return True
    End Function

    Private Function GetCompanyNumber() As String

        Dim String_Select As String
        Dim ComNumner As String = ""

        String_Select = "SELECT STARTINGBLOCKNUMBER as pNUMBER FROM SCRMENTITYNUMBERMANAGER "
        ConnectDataBase(String_Select, UserSchema1, DatabaseName1, DatabaseHost1)
        dbConnection.Open()
        Try
            ComNumner = Connect_Command.ExecuteScalar()
        Catch ex As Exception
            SendMailToPLMAdmin(ex.Message.ToString)
        Finally
            dbConnection.Close()
        End Try

        If ComNumner <> "" Then

            Dim Comm As OleDb.OleDbCommand
            Dim DataTran As OleDb.OleDbTransaction

            Try

                dbConnection = New dbConnector.dbSelector.dbConn
                dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
                dbConnection.Open()
                dbConnection.set_Command(DataTran, Comm)

                Comm.CommandText = "UPDATE SCRMENTITYNUMBERMANAGER SET STARTINGBLOCKNUMBER = STARTINGBLOCKNUMBER + 1"
                Comm.ExecuteNonQuery()

                Comm.Transaction.Commit()

            Catch ex As Exception
                Comm.Transaction.Rollback()
                SendMailToPLMAdmin(ex.Message.ToString)
                Return ""
            Finally
                dbConnection.Close()
            End Try
        End If

        Return ComNumner

    End Function

    Private Function GetFacilityNumber() As String

        Dim String_Select As String
        Dim ComNumner As String = ""

        String_Select = "SELECT STARTINGBLOCKNUMBER as pNUMBER FROM SCRMENTITYNUMBERMANAGER"
        ConnectDataBase(String_Select, UserSchema1, DatabaseName1, DatabaseHost1)
        dbConnection.Open()
        Try
            ComNumner = Connect_Command.ExecuteScalar()
        Catch ex As Exception
            SendMailToPLMAdmin(ex.Message.ToString)
        Finally
            dbConnection.Close()
        End Try

        If ComNumner <> "" Then

            Dim Comm As OleDb.OleDbCommand
            Dim DataTran As OleDb.OleDbTransaction

            Try

                dbConnection = New dbConnector.dbSelector.dbConn
                dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
                dbConnection.Open()
                dbConnection.set_Command(DataTran, Comm)

                Comm.CommandText = "UPDATE SCRMENTITYNUMBERMANAGER SET STARTINGBLOCKNUMBER = STARTINGBLOCKNUMBER + 1"
                Comm.ExecuteNonQuery()

                Comm.Transaction.Commit()

            Catch ex As Exception
                Comm.Transaction.Rollback()
                SendMailToPLMAdmin(ex.Message.ToString)
                Return ""
            Finally
                dbConnection.Close()
            End Try
        End If

        Return ComNumner

    End Function

    Private Function GetExistCompanyData(ByVal pSAPCODE As String) As Boolean

        PKID_COMPANY = ""
        PKID_COMPANYNAME = ""
        PKID_COMPANYADDRESS = ""
        COMPANYCREATEDDATE = ""

        Dim foundFlag As Boolean = False

        Dim Comm As OleDb.OleDbCommand
        Dim DR As OleDb.OleDbDataReader
        Dim String_Select As String = ""

        String_Select = "SELECT distinct " & _
                        "t1.num as COMPANY_NUMBER, " & _
                        "t1.pkid as PKID_COMPANY " & _
                        ",t2.pkid as PKID_COMPANYNAME " & _
                        ",t3.pkid as PKID_COMPANYADDRESS " & _
                        ",to_char(T1.CREATIONDATE,'DD/MM/YYYY') as COMPANYCREATEDDATE " & _
                        "FROM scrmCompany t1  " & _
                        "INNER JOIN scrmEntityFreeTextName t2 ON t1.pkid = t2.fkEntity  and t2.langid = 0  " & _
                        "INNER JOIN scrmAddress t3 ON t1.pkid = t3.fkParent  " & _
                        "INNER JOIN specLegacySpecJoin t9 ON t1.pkid = t9.fkSpecID  " & _
                        "INNER JOIN specLegacyProfile t10 ON t9.fkLegacyProfileID = t10.pkid and t10.systemname = 'SAP' " & _
                        "where t9.equivalent = '" & pSAPCODE & "' "
        Try

            dbConnection = New dbConnector.dbSelector.dbConn
            dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
            dbConnection.Open()
            dbConnection.set_Command(Comm, String_Select)
            DR = Comm.ExecuteReader
            While DR.Read
                foundFlag = True

                If Not IsDBNull(DR("COMPANY_NUMBER")) Then
                    pNUMBER = DR("COMPANY_NUMBER")
                End If
                If Not IsDBNull(DR("PKID_COMPANY")) Then
                    PKID_COMPANY = DR("PKID_COMPANY")
                End If
                If Not IsDBNull(DR("PKID_COMPANYNAME")) Then
                    PKID_COMPANYNAME = DR("PKID_COMPANYNAME")
                End If
                If Not IsDBNull(DR("PKID_COMPANYADDRESS")) Then
                    PKID_COMPANYADDRESS = DR("PKID_COMPANYADDRESS")
                End If
                If Not IsDBNull(DR("COMPANYCREATEDDATE")) Then
                    COMPANYCREATEDDATE = DR("COMPANYCREATEDDATE")
                End If

            End While

        Catch ex As Exception
            dbConnection.Close()
            SendMailToPLMAdmin(ex.Message.ToString)
            Return False
        Finally
            dbConnection.Close()
        End Try

        Return foundFlag
    End Function

    Private Function GetExistFacilityData() As Boolean

        PKID_Facility = ""
        PKID_FacilityName = ""
        PKID_FacilityAddress = ""

        Dim foundFlag As Boolean = False
        Dim dateString As String = CDate(COMPANYCREATEDDATE).ToString("ddMMyyyy")

        Dim Comm As OleDb.OleDbCommand
        Dim DR As OleDb.OleDbDataReader
        Dim String_Select As String = ""

        String_Select = " Select t1.num as FACILITY_NUMBER, t1.PKID as PKID_FACILITY, t2.PKID as PKID_FACILITYNAME, t3.PKID as PKID_FACILITYADDRESS " & _
                        "From scrmFacility t1 " & _
                        "INNER JOIN scrmEntityFreeTextName t2 ON t1.pkid = t2.fkEntity  and t2.langid = 0  " & _
                        "INNER JOIN scrmAddress t3 ON t1.pkid = t3.fkParent  " & _
                        "WHERE FKCOMPANY = '" & PKID_COMPANY & "' " & _
                        "and to_char(t1.CREATIONDATE,'ddMMyyyy') ='" & dateString & "' and t1.creationdate <> to_date('31/12/9999 ','DD/MM/YYYY')"

        Try

            dbConnection = New dbConnector.dbSelector.dbConn
            dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
            dbConnection.Open()
            dbConnection.set_Command(Comm, String_Select)
            DR = Comm.ExecuteReader
            While DR.Read
                foundFlag = True

                If Not IsDBNull(DR("FACILITY_NUMBER")) Then
                    pFNUMBER = DR("FACILITY_NUMBER")
                End If

                If Not IsDBNull(DR("PKID_FACILITY")) Then
                    PKID_Facility = DR("PKID_FACILITY")
                End If

                If Not IsDBNull(DR("PKID_FACILITYNAME")) Then
                    PKID_FacilityName = DR("PKID_FACILITYNAME")
                End If

                If Not IsDBNull(DR("PKID_FACILITYADDRESS")) Then
                    PKID_FacilityAddress = DR("PKID_FACILITYADDRESS")
                End If
            End While

        Catch ex As Exception
            dbConnection.Close()
            SendMailToPLMAdmin(ex.Message.ToString)
            Return False
        Finally
            dbConnection.Close()
        End Try

        Return foundFlag

    End Function

    Private Function GetPKID(ByVal pCOUNTRY As String) As Boolean

        PKID_COMPANY = ""
        PKID_Facility = ""
        PKID_COMPANYNAME = ""
        PKID_FacilityName = ""
        PKID_COMPANYADDRESS = ""
        PKID_FacilityAddress = ""
        COMPANY_SpecLegacySpecJoin = ""
        COMPANY_scrmBUsRelationship = ""
        FACILITY_scrmBUsRelationship = ""
        PKID_COUNTRY = ""
        PKID_SPECLEGACYPROFILE = ""
        PKID_USERS = ""
        PKID_ENTITYSTATUS = ""
        PKID_BUSINESSUNIT = ""


        Dim foundFlag As Boolean = False

        Dim Comm As OleDb.OleDbCommand
        Dim DR As OleDb.OleDbDataReader
        Dim String_Select As String = ""

        String_Select = " Select " & _
                        "(Select '5002'||NEWID()  from dual) as PKID_COMPANY " & _
                        ",(Select '5001'||NEWID() from dual) as PKID_Facility " & _
                        ",(Select '5015'||NEWID() from dual) as PKID_COMPANYNAME " & _
                        ",(Select '5015'||NEWID() from dual) as PKID_FACILITYNAME " & _
                        ",(Select '5016'||NEWID() from dual) as PKID_COMPANYADDRESS " & _
                        ",(Select '5016'||NEWID() from dual) as PKID_FACILITYADDRESS " & _
                        ",(Select '2031'||NEWID() from dual) as COMPANY_SpecLegacySpecJoin " & _
                        ",(Select '5030'||NEWID() from dual) as COMPANY_scrmBUsRelationship " & _
                        ",(Select '5030'||NEWID() from dual) as FACILITY_scrmBUsRelationship " & _
                        ",(Select PKID From Countries Where isocode = '" & pCOUNTRY & "' )as PKID_COUNTRY " & _
                        ",(Select PKID From specLegacyProfile Where systemname = 'SAP') as PKID_SPECLEGACYPROFILE " & _
                        ",(Select PKID From Users Where username = 'prodikaadmin') as PKID_USERS " & _
                        ",(Select PKID FROM scrmEntityStatus Where name = 'Approved') as PKID_ENTITYSTATUS " & _
                        ",(Select c1.PKID From commonBusinessUnit c1, commonBusinessUnitName c2,commonBUNamespace c3 " & _
                        "Where c1.pkid = c2.fkspecbusinessunit and upper(c2.name) = 'THAILAND'  and c1.fkbunamespace = c3.pkid " & _
                        "and c3.namespaceid = 'scrm' and langid = 0) as PKID_BUSINESSUNIT " & _
                        "From dual "

        Try

            dbConnection = New dbConnector.dbSelector.dbConn
            dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
            dbConnection.Open()
            dbConnection.set_Command(Comm, String_Select)
            DR = Comm.ExecuteReader

            While DR.Read
                foundFlag = True

                If Not IsDBNull(DR("PKID_COMPANY")) Then
                    PKID_COMPANY = DR("PKID_COMPANY")
                End If

                If Not IsDBNull(DR("PKID_FACILITY")) Then
                    PKID_Facility = DR("PKID_FACILITY")
                End If

                If Not IsDBNull(DR("PKID_COMPANYNAME")) Then
                    PKID_COMPANYNAME = DR("PKID_COMPANYNAME")
                End If

                If Not IsDBNull(DR("PKID_FACILITYNAME")) Then
                    PKID_FacilityName = DR("PKID_FACILITYNAME")
                End If

                If Not IsDBNull(DR("PKID_COMPANYADDRESS")) Then
                    PKID_COMPANYADDRESS = DR("PKID_COMPANYADDRESS")
                End If

                If Not IsDBNull(DR("PKID_FACILITYADDRESS")) Then
                    PKID_FacilityAddress = DR("PKID_FACILITYADDRESS")
                End If

                If Not IsDBNull(DR("COMPANY_SPECLEGACYSPECJOIN")) Then
                    COMPANY_SpecLegacySpecJoin = DR("COMPANY_SPECLEGACYSPECJOIN")
                End If

                If Not IsDBNull(DR("COMPANY_SCRMBUSRELATIONSHIP")) Then
                    COMPANY_scrmBUsRelationship = DR("COMPANY_SCRMBUSRELATIONSHIP")
                End If

                If Not IsDBNull(DR("FACILITY_SCRMBUSRELATIONSHIP")) Then
                    FACILITY_scrmBUsRelationship = DR("FACILITY_SCRMBUSRELATIONSHIP")
                End If

                If Not IsDBNull(DR("PKID_COUNTRY")) Then
                    PKID_COUNTRY = DR("PKID_COUNTRY")
                End If

                If Not IsDBNull(DR("PKID_SPECLEGACYPROFILE")) Then
                    PKID_SPECLEGACYPROFILE = DR("PKID_SPECLEGACYPROFILE")
                End If

                If Not IsDBNull(DR("PKID_USERS")) Then
                    PKID_USERS = DR("PKID_USERS")
                End If

                If Not IsDBNull(DR("PKID_ENTITYSTATUS")) Then
                    PKID_ENTITYSTATUS = DR("PKID_ENTITYSTATUS")
                End If

                If Not IsDBNull(DR("PKID_BUSINESSUNIT")) Then
                    PKID_BUSINESSUNIT = DR("PKID_BUSINESSUNIT")
                End If

            End While

        Catch ex As Exception
            dbConnection.Close()
            SendMailToPLMAdmin(ex.Message.ToString)
            Return False
        Finally
            dbConnection.Close()
        End Try

        Return foundFlag

    End Function

    Private Function Read_TextFile(ByVal fileName As String) As Boolean

        Dim data_text() As String
        Dim dr As DataRow
        Dim line As Integer = 0
        Dim aIng() As String = industIng.Split(",")
        Dim aPK() As String = industPK.Split(",")



        Try

            Using reader As New StreamReader(fileName, System.Text.Encoding.Default, True)

                While reader.Peek <> -1

                    data_text = reader.ReadLine.Split(vbTab)

                    If data_text(0) <> "00" And data_text(0) <> "99" Then

                        Dim ingFlag As Boolean = False
                        Dim packFlag As Boolean = False

                        Dim text_spit As String = data_text(27)
                        Dim TrnType As String = data_text(55)

                        If text_spit <> "" Then
                            If text_spit.Length >= 2 Then
                                Dim checkText As String = data_text(27).Substring(0, 2)

                                For i As Integer = 0 To aIng.Length - 1
                                    If checkText = aIng(i) Then

                                        ingFlag = True
                                        Exit For
                                    End If
                                Next

                                If ingFlag = False Then
                                    For i As Integer = 0 To aPK.Length - 1
                                        If checkText = aPK(i) Then

                                            packFlag = True
                                            Exit For
                                        End If
                                    Next
                                End If

                                'If checkText = "02" Or checkText = "04" Or checkText = "05" Or checkText = "21" Then

                                If TrnType = "A" Or TrnType = "C" Then

                                    If ingFlag = True Then
                                        dr = dtIng.NewRow
                                        dr("SAP_CODE") = data_text(1)
                                        dr("NAME") = Replace(data_text(6), "'", "''")
                                        dr("STREET1") = Replace(data_text(10), "'", "''") & " " & Replace(data_text(11), "'", "''")
                                        dr("STREET2") = Replace(data_text(12), "'", "''")
                                        dr("CITY") = Replace(data_text(13), "'", "''")
                                        dr("POSTAL_CODE") = data_text(14)
                                        dr("COUNTRY") = data_text(15)
                                        dr("STATE_OR_PROVINCE") = Replace(data_text(16), "'", "''")
                                        dr("PHONE") = data_text(20)
                                        dtIng.Rows.Add(dr)

                                    ElseIf packFlag = True Then

                                        dr = dtPack.NewRow
                                        dr("SAP_CODE") = data_text(1)
                                        dr("NAME") = Replace(data_text(6), "'", "''")
                                        dr("STREET1") = Replace(data_text(10), "'", "''") & " " & Replace(data_text(11), "'", "''")
                                        dr("STREET2") = Replace(data_text(12), "'", "''")
                                        dr("CITY") = Replace(data_text(13), "'", "''")
                                        dr("POSTAL_CODE") = data_text(14)
                                        dr("COUNTRY") = data_text(15)
                                        dr("STATE_OR_PROVINCE") = Replace(data_text(16), "'", "''")
                                        dr("PHONE") = data_text(20)
                                        dtPack.Rows.Add(dr)

                                    End If

                                End If

                                'End If
                            End If

                        End If

                    End If

                End While

                reader.Close()

            End Using

            'dtAll.Merge(dtIng)
            'dtAll.Merge(dtPack)


        Catch ex As Exception
            'Dim WriteFile As StreamWriter
            'WriteFile = File.AppendText(File_Error)
            'WriteFile.WriteLine("" & ex.Message.ToString & " " & Now & " ")
            'WriteFile.Flush()
            'WriteFile.Close()
            SendMailToPLMAdmin(ex.Message.ToString)
            Return False
        End Try

        Return True
    End Function

    Sub CreateSchema(ByRef _dt As DataTable)
        With _dt.Columns
            .Add("SAP_CODE")
            .Add("NAME")
            .Add("STREET1")
            .Add("STREET2")
            .Add("CITY")
            .Add("POSTAL_CODE")
            .Add("COUNTRY")
            .Add("STATE_OR_PROVINCE")
            .Add("PHONE")
            '.Add("TRANSACTION_TYPE")
        End With

    End Sub

    Private Function ReadMailAddressIng() As DataTable

        myDataSet.Clear()
        myDataSet = New DataSet
        Dim dtAdd As DataTable
        Dim String_Select As String

        String_Select = " SELECT t1.firstname|| ' '||t1.lastname as Name, t1.Email FROM Users t1  " & _
                        "INNER JOIN UserGroupJoin t2 ON t1.pkid = t2.fkUsers  " & _
                        "INNER JOIN Groups t3 ON t2.fkGroups = t3.pkid " & _
                        "INNER JOIN GroupsML t4 ON t3.pkid = t4.fkGroup " & _
                        "where t4.name = 'Purchasing Admin - Ingredient'  "

        dbConnection = New dbConnector.dbSelector.dbConn
        dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
        dbConnection.set_Adapter(DataAdap, String_Select)

        Try
            dbConnection.Open()
            DataAdap.Fill(myDataSet, "PLM")
            dtAdd = myDataSet.Tables("PLM")

        Catch ex As Exception

        Finally
            dbConnection.Close()
        End Try

        Return dtAdd

    End Function

    Private Function ReadMailAddressPK() As DataTable

        myDataSet.Clear()
        myDataSet = New DataSet
        Dim dtAdd As DataTable
        Dim String_Select As String

        String_Select = " SELECT t1.firstname|| ' '||t1.lastname as Name, t1.Email FROM Users t1  " & _
                        "INNER JOIN UserGroupJoin t2 ON t1.pkid = t2.fkUsers  " & _
                        "INNER JOIN Groups t3 ON t2.fkGroups = t3.pkid " & _
                        "INNER JOIN GroupsML t4 ON t3.pkid = t4.fkGroup " & _
                        "where t4.name = 'Purchasing Admin - Packaging'  "

        dbConnection = New dbConnector.dbSelector.dbConn
        dbConnection.set_dbConnector(UserSchema1, DatabaseBrand, DatabaseName1, DatabaseHost1)
        dbConnection.set_Adapter(DataAdap, String_Select)

        Try
            dbConnection.Open()
            DataAdap.Fill(myDataSet, "PLM")
            dtAdd = myDataSet.Tables("PLM")

        Catch ex As Exception

        Finally
            dbConnection.Close()
        End Try

        Return dtAdd

    End Function

    Private Sub SendMailToPLMAdmin(ByVal errorList As String)

        Dim mailSubject As String = "Vendor Master Interface in PLM System is Error"
        Dim mailMsg As String = PLMAdminMailMessage(errorList)
        SendMail(toAddAdmin, fromAddAdmin, mailSubject, mailMsg)
    End Sub

    Private Function PLMAdminMailMessage(ByVal errorList As String) As String

        Return "เรียน PLM Admin " & vbNewLine & _
               "" & vbNewLine & _
               "ผลการ Interface ข้อมูลของ Company Profile และ Facility Profile เกิดปัญหา ไม่สามารถอัพเดทผลมายังระบบ PLM ได้" & vbNewLine & _
               "รบกวน PLM Admin ตรวจสอบปัญหาดังกล่าว และดำเนินการต่อไป " & vbNewLine & _
               "" & vbNewLine & _
               errorList & _
               "" & vbNewLine & _
               "" & vbNewLine & _
               "Dear PLM Admin" & vbNewLine & _
               "" & vbNewLine & _
               "Vendor Master Interface  has some errors. The result could not update in PLM successfully." & vbNewLine & _
               "Kindly check the errors." & vbNewLine & _
               "" & vbNewLine & _
               errorList & _
               "" & vbNewLine & _
               "" & vbNewLine & _
               "Note:   This message is intended only for the individual or entity to which it is addressed and may contain information that is confidential and/or" & vbNewLine & _
               "        privileged. If you received this email in error, please delete it and notify the sender immediately. Any dissemination, distribution or copying " & vbNewLine & _
               "        of this communication by someone other than the intended the recipient, is strictly prohibited."

    End Function

    Private Function CompleteMailMessage(ByVal userName As String, ByVal companyAdd As String, ByVal companyUpdate As String, ByVal FacilityAdd As String, ByVal FacilityUpdate As String, ByVal flag As String) As String

        Dim companyAddTHAI As String = ""
        Dim companyAddENG As String = ""
        Dim companyUpdateTHAI As String = ""
        Dim companyUpdateENG As String = ""

        Dim FacilityAddTHAI As String = ""
        Dim FacilityAddENG As String = ""
        Dim FacilityUpdateTHAI As String = ""
        Dim FacilityUpdateENG As String = ""

        Dim interfaceType As String = ""
        If flag = "I" Then
            interfaceType = "(Ingredient)"

        ElseIf flag = "P" Then
            interfaceType = "(Packaging)"

        End If

        If companyAdd <> "" Then
            companyAddTHAI = "Company ที่ถูกสร้างใหม่ : " & vbNewLine & companyAdd
            companyAddENG = "Created Companies : " & vbNewLine & companyAdd
        End If

        If companyUpdate <> "" Then
            companyUpdateTHAI = "Company ที่ถูกแก้ไข : " & vbNewLine & companyUpdate
            companyUpdateENG = "Updated Companies : " & vbNewLine & companyUpdate
        End If

        If FacilityAdd <> "" Then
            FacilityAddTHAI = "Facility ที่ถูกสร้างใหม่ : " & vbNewLine & FacilityAdd
            FacilityAddENG = "Created Facilities : " & vbNewLine & FacilityAdd
        End If

        If FacilityUpdate <> "" Then
            FacilityUpdateTHAI = "Facility ที่ถูกแก้ไข : " & vbNewLine & FacilityUpdate
            FacilityUpdateENG = "Updated Facilities : " & vbNewLine & FacilityUpdate
        End If

        Return "เรียน คุณ " & userName & vbNewLine & _
                       "" & vbNewLine & _
                       "ผลการ Interface ข้อมูล Vendor" & interfaceType & " จากระบบ SAP สำเร็จ ท่านสามารถเข้าไปดูรายละเอียดได้ที่ Company Profile และ Facility Profile ในระบบ PLM " & vbNewLine & _
                       "โดยข้อมูล Company และ Facility ที่ถูกสร้างหรือแก้ไขมีดังนี้: " & vbNewLine & _
                        vbNewLine & _
                       companyAddTHAI & vbNewLine & _
                       companyUpdateTHAI & vbNewLine & _
                       "" & vbNewLine & _
                       "" & vbNewLine & _
                       FacilityAddTHAI & vbNewLine & _
                       FacilityUpdateTHAI & vbNewLine & _
                       "" & vbNewLine & _
                       "" & vbNewLine & _
                       "หมายเหตุ:  ข้อมูลนี้มีวัตถุประสงค์เฉพาะสำหรับแต่ละบุคคลหรือนิติบุคคล และเป็นข้อมูลที่เป็นความลับ หากคุณได้รับอีเมลล์ฉบับนี้จากความ " & vbNewLine & _
                       "         ผิดพลาดในการส่ง โปรดลบ และแจ้งผู้ส่งทันที ห้าม!!! เผยแพร่ หรือคัดลอกการสื่อสารนี้โดยเด็ดขาด" & vbNewLine & _
                       "" & vbNewLine & _
                       "ขอขอบพระคุณอย่างสูง" & vbNewLine & _
                       "" & vbNewLine & _
                       "" & vbNewLine & _
                       "Dear " & userName & vbNewLine & _
                       "" & vbNewLine & _
                       "Vendor Master" & interfaceType & " Interface is successful. You can check the result of Interface at Company Profile and Facility Profile in PLM system. " & vbNewLine & _
                       "The results of interface are: " & vbNewLine & _
                        vbNewLine & _
                       companyAddENG & vbNewLine & _
                       companyUpdateENG & vbNewLine & _
                       "" & vbNewLine & _
                       "" & vbNewLine & _
                       FacilityAddENG & vbNewLine & _
                       FacilityUpdateENG & vbNewLine & _
                       "" & vbNewLine & _
                       "" & vbNewLine & _
                       "Thank you for your participation " & vbNewLine & _
                       "" & vbNewLine & _
                       "PLM Admin" & vbNewLine & _
                       "Tel: " & vbNewLine & _
                       "Email: " & vbNewLine & _
                       "Note:   This message is intended only for the individual or entity to which it is addressed and may contain information that is confidential and/or" & vbNewLine & _
                       "        privileged. If you received this email in error, please delete it and notify the sender immediately. Any dissemination, distribution or copying " & vbNewLine & _
                       "        of this communication by someone other than the intended the recipient, is strictly prohibited."


    End Function

    Private Sub SendMail(ByVal msgTo As String, ByVal msgFrom As String, ByVal msgSubject As String, ByVal msgBody As String)

        Try
            Dim smtpServer As String = smtpIPServer
            Dim insMail As New MailMessage(msgFrom, msgTo, msgSubject, msgBody)

            Dim smtpClient As New SmtpClient
            smtpClient.Host = smtpServer
            smtpClient.UseDefaultCredentials = False
            smtpClient.Send(insMail)

        Catch ex As Exception

            Dim WriteFile As StreamWriter
            WriteFile = File.AppendText(File_Error)
            WriteFile.WriteLine("" & ex.Message & " " & Now & " ")
            WriteFile.Flush()
            WriteFile.Close()
            Me.Dispose()
            Me.Close()
        End Try

    End Sub

End Class