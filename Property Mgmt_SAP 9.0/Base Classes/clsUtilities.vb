Imports System.IO
Imports System.Xml.XmlDocument
Public Class clsUtilities

    Private strThousSep As String = ","
    Private strDecSep As String = "."
    Private intQtyDec As Integer = 3
    Private FormNum As Integer
    Const An As String = " æ "
    Const Ab As String = " ÝÞÜØ"
    Dim oCombo, oCombobox1, oCombobox2 As SAPbouiCOM.ComboBox
    Dim oEdit As SAPbouiCOM.EditText
    Dim oExEdit As SAPbouiCOM.EditText
    Dim oGrid As SAPbouiCOM.Grid
    Dim oRecordSet As SAPbobsCOM.Recordset
    Dim sQuery As String
    Dim SmtpServer As New Net.Mail.SmtpClient()
    Dim mail As New Net.Mail.MailMessage
    Dim mailServer As String
    Dim mailPort As String
    Dim mailId As String
    Dim mailUser As String
    Dim mailPwd As String
    Dim mailSSL As String
    Dim toID As String
    Dim ccID As String
    Public Sub New()
        MyBase.New()
        FormNum = 1
    End Sub


    Public Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim dtFrom, dtTo As Date
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        Dim dblAnnualRent, dblNoofMonths, dblRentalInstallment As Double
        Dim oCheckbox As SAPbouiCOM.CheckBoxColumn
        dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "7"))
        dtTo = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "9"))
        Dim strRentType As String
        oCombo = aForm.Items.Item("17").Specific
        strRentType = oCombo.Selected.Value
        Dim intNoofDays As Double
        Dim intNoofMonths As Integer
        Dim otest As SAPbobsCOM.Recordset
        dblAnnualRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "13"))
        dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "11"))
        dblRentalInstallment = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "19"))
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from [@Z_CONTRACT] where DocEntry=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
        intNoofDays = otest.Fields.Item("U_Z_NoofDays").Value
        otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
        Dim dtTo1 As Date
        If 1 = 1 Then ' otest.RecordCount <= 0 Then
            oUserTable = oApplication.Company.UserTables.Item("Z_CONINS")
            oGrid = aForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
            Dim intCount As Integer = 0

            While dtFrom <= dtTo
                If intCount > 0 Then
                    Select Case strRentType
                        Case "D" '
                            If intNoofDays = 0 Then
                                intNoofDays = 1
                            End If
                            dtTo1 = DateAdd(DateInterval.Day, 1, dtFrom)
                        Case "W"
                            dtTo1 = DateAdd(DateInterval.Day, 7, dtFrom)
                        Case "M"
                            dtTo1 = DateAdd(DateInterval.Month, 1, dtFrom)
                        Case "Q"
                            dtTo1 = DateAdd(DateInterval.Month, 3, dtFrom)
                        Case "S"
                            dtTo1 = DateAdd(DateInterval.Month, 6, dtFrom)
                        Case "A"
                            dtTo1 = DateAdd(DateInterval.Year, 1, dtFrom)
                    End Select
                    dtTo1 = dtTo1.AddDays(-1)
                Else
                    Select Case strRentType
                        Case "D" '
                            If intNoofDays = 0 Then
                                intNoofDays = 1
                            End If
                            dtTo1 = DateAdd(DateInterval.Day, 1, dtFrom)
                        Case "W"
                            dtTo1 = DateAdd(DateInterval.Day, 7, dtFrom)
                        Case "M"
                            dtTo1 = DateAdd(DateInterval.Month, 0, dtFrom)
                        Case "Q"
                            dtTo1 = DateAdd(DateInterval.Month, 2, dtFrom)
                        Case "S"
                            dtTo1 = DateAdd(DateInterval.Month, 5, dtFrom)
                        Case "A"
                            dtTo1 = DateAdd(DateInterval.Year, 1, dtFrom)
                    End Select
                End If

                Dim intMOnth As Integer = dtTo1.Month
                intMOnth = DateTime.DaysInMonth(dtTo1.Year, dtTo1.Month)
                If strRentType <> "D" And strRentType <> "W" Then
                    dtTo1 = New DateTime(dtTo1.Year, dtTo1.Month, intMOnth)
                End If
                Dim dt12 As Date = dtTo1
                Select Case strRentType
                    Case "D" '
                        If intNoofDays = 0 Then
                            intNoofDays = 1
                        End If
                        dt12 = DateAdd(DateInterval.Day, 1, dtTo1)
                    Case "W"
                        dt12 = DateAdd(DateInterval.Day, 7, dtTo1)
                    Case "M"
                        dt12 = DateAdd(DateInterval.Month, 1, dtTo1)
                    Case "Q"
                        dt12 = DateAdd(DateInterval.Month, 3, dtTo1)
                    Case "S"
                        dt12 = DateAdd(DateInterval.Month, 6, dtTo1)
                    Case "A"
                        dt12 = DateAdd(DateInterval.Year, 1, dtTo1)
                End Select
                If dt12 >= dtTo Then
                    dtTo1 = dtTo
                End If

                otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount',U_Z_Manual from [@Z_CONINS] where U_Z_Month=" & Month(dtFrom) & " and U_Z_Year=" & Year(dtFrom) & " and   U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
                If otest.RecordCount <= 0 Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_CONINS", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                    oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                    oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                    oUserTable.UserFields.Fields.Item("U_Z_MonthRent").Value = oApplication.Utilities.getEdittextvalue(aForm, "19")

                    oUserTable.UserFields.Fields.Item("U_Z_StartDate1").Value = dtFrom
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate1").Value = dtTo1 '(oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_Month").Value = Month(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Year").Value = Year(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = dblAnnualRent / dblNoofMonths
                    oUserTable.UserFields.Fields.Item("U_Z_RentType").Value = oCombo.Selected.Value
                    If oUserTable.Add <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                Else
                    If otest.Fields.Item("U_Z_Manual").Value <> "Y" Then
                        strCode = otest.Fields.Item("Code").Value
                        oUserTable.GetByKey(strCode)
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode
                        oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                        oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                        oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                        oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                        oUserTable.UserFields.Fields.Item("U_Z_StartDate1").Value = dtFrom
                        oUserTable.UserFields.Fields.Item("U_Z_EndDate1").Value = dtTo1 '(oApplication.Utilities.getEdittextvalue(aForm, "9"))
                        oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                        oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                        oUserTable.UserFields.Fields.Item("U_Z_MonthRent").Value = oApplication.Utilities.getEdittextvalue(aForm, "19")
                        oUserTable.UserFields.Fields.Item("U_Z_Month").Value = Month(dtFrom)
                        oUserTable.UserFields.Fields.Item("U_Z_Year").Value = Year(dtFrom)
                        oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = dblAnnualRent / dblNoofMonths
                        oUserTable.UserFields.Fields.Item("U_Z_RentType").Value = oCombo.Selected.Value
                        If oUserTable.Update <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                    End If
                End If

                Select Case strRentType
                    Case "D" '
                        If intNoofDays = 0 Then
                            intNoofDays = 1
                        End If
                        dtFrom = DateAdd(DateInterval.Day, 1, dtFrom)
                    Case "W"
                        dtFrom = DateAdd(DateInterval.Day, 7, dtFrom)
                    Case "M"
                        dtFrom = DateAdd(DateInterval.Month, 1, dtFrom)
                    Case "Q"
                        dtFrom = DateAdd(DateInterval.Month, 3, dtFrom)
                    Case "S"
                        dtFrom = DateAdd(DateInterval.Month, 6, dtFrom)
                    Case "A"
                        dtFrom = DateAdd(DateInterval.Year, 1, dtFrom)
                End Select

                If strRentType <> "D" And strRentType <> "W" Then
                    dtFrom = dtTo1.AddDays(1)
                End If

                ' dtFrom = DateAdd(DateInterval.Month, 1, dtFrom)
            End While
        End If
        Return True

    End Function

    Public Sub cmdDraftToOrder_Click(ByVal aDocEntry As Integer, ByVal aobjectType As String)

        Dim pDraft As SAPbobsCOM.Documents
        pDraft = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDrafts)
        Dim pOrder As SAPbobsCOM.Documents
        ' oApplication.Company.XmlExportType = SAPbobsCOM.BoXmlExportTypes.xet_ExportImportMode 'xet_ExportImportMode
        '  oApplication.Company.XMLAsString = False
        pDraft.GetByKey(aDocEntry)
        If pDraft.SaveDraftToDocument() <> 0 Then
            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
        'pDraft.SaveXML("c:\drafts1.xml")
        'Dim Filename1 As String = "C:\drafts1.xml"
        ''   XmlDocument xml = new XmlDocument();
        ''xml.Load(fileName);

        ''/* put the <DocObjectCode> value into the <object> value in the xml */
        ''string objType = xml.SelectSingleNode("BOM/BO/Documents/row/DocObjectCode").InnerText;
        ''xml.SelectSingleNode("BOM/BO/AdmInfo/Object").InnerText = objType;

        ''/* remove the <DocObjectCode> element as the doc won't load as a valid doc with it */
        ''XmlNode node = xml.SelectSingleNode("BOM/BO/Documents/row/DocObjectCode");
        ''xml.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(node);

        ''/* remove the <DocNum> value from the xml so a new one gets assigned when it is saved */
        ''xml.SelectSingleNode("BOM/BO/Documents/row/DocNum").InnerText = "";

        ''xml.Save(fileName);


        'Dim xml As Xml.XmlDocument = New Xml.XmlDocument()
        'xml.Load("C:\drafts1.xml")
        ''       put the <DocObjectCode> value into the <object> value in the xml */


        'Dim objType As String = xml.SelectSingleNode("BOM/BO/Documents/row/DocObjectCode").InnerText
        'xml.SelectSingleNode("BOM/BO/AdmInfo/Object").InnerText = objType
        ''       remove the <DocObjectCode> element as the doc won't load as a valid doc with it */


        'Dim node As Xml.XmlNode = xml.SelectSingleNode("BOM/BO/Documents/row/DocObjectCode")
        'xml.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(node)
        ''       remove the <DocNum> value from the xml so a New one gets assigned when it is saved */
        ''ReqType
        'node = xml.SelectSingleNode("BOM/BO/Documents/row/ReqType")
        'xml.SelectSingleNode("BOM/BO/Documents/row").RemoveChild(node)


        'node = xml.SelectSingleNode("BOM/BO/Document_Lines/row/EnableReturnCost")
        'xml.SelectSingleNode("BOM/BO/Document_Lines/row").RemoveChild(node)

        'node = xml.SelectSingleNode("BOM/BO/Document_Lines/row/ReturnCost")
        'xml.SelectSingleNode("BOM/BO/Document_Lines/row").RemoveChild(node)

        'node = xml.SelectSingleNode("BOM/BO/Document_Lines/row/LineVendor")
        'xml.SelectSingleNode("BOM/BO/Document_Lines/row").RemoveChild(node)

        'xml.SelectSingleNode("BOM/BO/Documents/row/DocNum").InnerText = ""

        'xml.Save("C:\drafts1.xml")

        ''Here you should add a code that will change the Object's
        ''value from 112 (Drafts) to 17 (Orders) and also you should
        ''remove the DocObjectCode node from the xml. You can use any
        ''xml parser.
        ''
        ''Create a new order
        'pOrder = oApplication.Company.GetBusinessObjectFromXML("c:\drafts1.xml", 0)
        'If pOrder.Add() <> 0 Then

        '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        'Else

        '    'If pDraft.Close() <> 0 Then
        '    '    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    'End If
        '    'pDraft.ClosingOption = SAPbobsCOM.ClosingOptionEnum.coByCurrentSystemDate

        'End If

    End Sub

    Public Function createHRMainAuthorization() As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)
        '//Mandatory field, which is the key of the object.
        '//The partner namespace must be included as a prefix followed by _
        mUserPermission.PermissionID = "PM"
        '//The Name value that will be displayed in the General Authorization Tree
        mUserPermission.Name = "Property Management Addon"
        '//The permission that this object can get
        mUserPermission.Options = SAPbobsCOM.BoUPTOptions.bou_FullReadNone
        '//In case the level is one, there Is no need to set the FatherID parameter.
        '   mUserPermission.Levels = 1
        RetVal = mUserPermission.Add
        If RetVal = 0 Or -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Function addChildAuthorization(ByVal aChildID As String, ByVal aChildiDName As String, ByVal aorder As Integer, ByVal aFormType As String, ByVal aParentID As String, ByVal Permission As SAPbobsCOM.BoUPTOptions) As Boolean
        Dim RetVal As Long
        Dim ErrCode As Long
        Dim ErrMsg As String
        Dim mUserPermission As SAPbobsCOM.UserPermissionTree
        mUserPermission = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserPermissionTree)

        mUserPermission.PermissionID = aChildID
        mUserPermission.Name = aChildiDName
        mUserPermission.Options = Permission ' SAPbobsCOM.BoUPTOptions.bou_FullReadNone

        '//For level 2 and up you must set the object's father unique ID
        'mUserPermission.Level
        mUserPermission.ParentID = aParentID
        mUserPermission.UserPermissionForms.DisplayOrder = aorder
        '//this object manages forms
        ' If aFormType <> "" Then
        mUserPermission.UserPermissionForms.FormType = aFormType
        ' End If

        RetVal = mUserPermission.Add
        If RetVal = 0 Or RetVal = -2035 Then
            Return True
        Else
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End If


    End Function

    Public Sub AuthorizationCreation()
        addChildAuthorization("Setup", " Setup", 2, "", "PM", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("Trans", "Transactions", 2, "", "PM", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

        'Setup
        addChildAuthorization("Service", "Service Master", 3, frm_Earning, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("Location", "Location Master", 3, frm_Location, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("Unitprice", "Property UnitPrice List", 3, frm_PriceList, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("ProType", "Property Type", 3, frm_PropertyType, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("UnitType", "Proprty Unit Type", 3, frm_PropertyUnitType, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("ProFac", "Property Facility Details", 3, frm_ProFac, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("ProUnFac", "Property Unit Facility Details", 3, frm_ProUFac, "Setup", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)

        'Transactions

        addChildAuthorization("Property", "Properties", 3, frm_PropertyData, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("Eval", "Property Evaluation", 3, frm_Evaluation, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("ProUnit", "Property Unit", 3, frm_PropertyUnitDetails, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("Insurance", "Insurance", 3, frm_Insurance, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("Reser", "Reservation", 3, frm_Reservation, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("OwnContract", "Owner Contract", 3, frm_Contracts, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("TenContract", "Tenant Contract", 3, frm_TenContracts, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullReadNone)
        addChildAuthorization("RenWizard", "Contract Renewal Wizard", 3, frm_renewal, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("Search", "Search", 3, frm_Search, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        addChildAuthorization("BillWizard", "Bill Generation Wizar", 3, frm_BillGeneration, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        'addChildAuthorization("DPWizard", "Downpayment Wizard", 3, frm_PostingWizard, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)
        'addChildAuthorization("BillWizard", "Bill Generation Wizar", 3, frm_do, "Trans", SAPbobsCOM.BoUPTOptions.bou_FullNone)

    End Sub
    Public Sub AssignSerialNo(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 1 To aMatrix.RowCount
            aMatrix.Columns.Item("SlNo").Cells.Item(intRow).Specific.value = intRow
        Next
        aform.Freeze(False)
    End Sub

    Public Sub AssignRowNo(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        For intRow As Integer = 1 To aMatrix.VisualRowCount
            aMatrix.Columns.Item("V_-1").Cells.Item(intRow).Specific.value = intRow
        Next
        aform.Freeze(False)
    End Sub
    Public Function validateAuthorization(ByVal aUserId As String, ByVal aFormUID As String) As Boolean
        Dim oAuth As SAPbobsCOM.Recordset
        oAuth = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim struserid As String
        '    Return False
        struserid = oApplication.Company.UserName
        oAuth.DoQuery("select * from UPT1 where FormId='" & aFormUID & "'")
        If (oAuth.RecordCount <= 0) Then
            Return True
        Else
            Dim st As String
            st = oAuth.Fields.Item("PermId").Value
            st = "Select * from USR3 where PermId='" & st & "' and UserLink=" & aUserId
            oAuth.DoQuery(st)
            If oAuth.RecordCount > 0 Then
                If oAuth.Fields.Item("Permission").Value = "N" Then
                    Return False
                End If
                Return True
            Else
                Return True
            End If

        End If

        Return True

    End Function
#Region "ValidateCode"



    Public Function ValidateCode(ByVal aCode As String, ByVal aModule As String) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strqry As String = ""
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aModule = "PROFAC" Then
            strqry = "Select * from ""@Z_PROP2"" where ""U_Z_CODE""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Facility Already mapped in Property  Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If


        ElseIf aModule = "PROUFAC" Then
            strqry = "Select * from ""@Z_PROPUNIT2"" where ""U_Z_CODE""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Facility Already mapped in Property Unit Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If


        ElseIf aModule = "ALLOW" Then
            strqry = "Select * from ""@Z_HR_SALST1"" where ""U_Z_AllCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Allowance Code Already mapped in Salary Scale Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "RATING" Then
            strqry = "select * from ""@Z_HR_SEAPP1"" where ""U_Z_SelfRaCode""='" & aCode & "' or ""U_Z_MgrRaCode""='" & aCode & "' or ""U_Z_SMRaCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Rating Code Already mapped in Appraisals...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "EXPENCES" Then
            strqry = "select * from ""@Z_HR_TRAPL1"" where ""U_Z_ExpName""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Expences Already mapped in Travel Agenda...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "COURSE" Then
            strqry = "select * from ""@Z_HR_OTRIN"" where ""U_Z_CourseCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Course Code Already mapped in Training Agenda. Training Agenda Code : " & oTemp.Fields.Item("U_Z_TrainCode").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "TRAINER" Then
            strqry = "select * from ""@Z_HR_OTRIN"" where ""U_Z_InsName""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Trainer Code Already mapped in Training Agenda. Training Agenda Code : " & oTemp.Fields.Item("U_Z_TrainCode").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "TRAPLAN" Then
            strqry = "select * from ""@Z_HR_OASSTP"" where ""U_Z_TraCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Travel Agenda Code Already mapped in Employee Master. Employee Code : " & oTemp.Fields.Item("U_Z_EmpId").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "TRAINAGENDA" Then
            strqry = "select * from ""@Z_HR_TRIN1"" where ""U_Z_TrainCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Training Code Already mapped in Employee Master. Employee Code : " & oTemp.Fields.Item("U_Z_HREmpID").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "POSITION" Then
            strqry = "Select * from [@Z_HR_ORGST] where ""U_Z_PosCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Position Code already mapped in Organization Structure.Organization Code :" & oTemp.Fields.Item("U_Z_OrgCode").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

            strqry = "Select * from OHPS where ""Name""='" & aCode & "'"
            strqry = "SELECT *  FROM OHEM T0  INNER JOIN OHPS T1 ON T0.position = T1.posID WHERE T1.[name] ='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Position Code already mapped in Employee Master :", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

        ElseIf aModule = "JOBSCREEN" Then
            strqry = "Select * from ""@Z_HR_OPOSIN"" where ""U_Z_JobCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Job Code already mapped in Position Master.Position Code :" & oTemp.Fields.Item("U_Z_PosCode").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "SALARY" Then
            strqry = "Select * from ""@Z_HR_OPOSCO"" where ""U_Z_SalCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Salary Code Already mapped in Job Screen. Job Code  :" & oTemp.Fields.Item("U_Z_PosCode").Value, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "RECREQREASON" Then
            strqry = "select * from ""@Z_HR_ORMPREQ"" where ""U_Z_ReqReason""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Recruitment Request Reason Already mapped in Recruitment Requisition...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "INTRATING" Then
            strqry = "select * from ""@Z_HR_OHEM2"" where ""U_Z_Rating""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Interview Rating Code Already mapped in Interview Process form...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "INTERVIEWTYPE" Then
            strqry = "select * from ""@Z_HR_OHEM2"" where ""U_Z_InType""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Interview Type Already mapped in Interview Process form...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "RESPONSE" Then
            strqry = "Select * from ""@Z_HR_EXFORM1"" where ""U_Z_ResCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Responsibilities Code Already mapped in Employee exit initialization...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "QUS" Then
            strqry = "Select * from ""@Z_HR_EXFORM2"" where ""U_Z_QusCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Questionnaire Code Already mapped in Employee exit Interview form...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "LANG" Then
            strqry = "Select * from ""@Z_HR_RMPREQ5"" where ""U_Z_LanCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Language Code Already mapped in Recruitment Requisition...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "COUCAT" Then
            strqry = "Select * from ""@Z_HR_OCOUR"" where ""U_Z_CouCatCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Course Category Code Already mapped in Course Setup...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "COUTYP" Then
            strqry = "Select * from ""@Z_HR_OTRIN"" where ""U_Z_CourseTypeCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Course Type Code Already mapped in Training Agenda Setup...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "DEPT" Then
            strqry = "Select * from OUDP where ""Name""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                Return True
            End If
        ElseIf aModule = "BENEFIT" Then
            strqry = "Select * from ""@Z_HR_SALST2"" where ""U_Z_BeneCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Benefits Code Already mapped in Salary Scale Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "LEVEL" Then
            strqry = "Select * from ""@Z_HR_OSALST"" where ""U_Z_LevlCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Level Code Already mapped in Salary Scale Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "GRADE" Then
            strqry = "Select * from ""@Z_HR_OSALST"" where ""U_Z_GrdeCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Grade Code Already mapped in Salary Scale Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "OBJLOAN" Then
            strqry = "Select * from ""@Z_HR_OBJLOAN"" where ""U_Z_ObjCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Objects on Loan Code Already mapped in Employee Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "COMP" Then
            strqry = "Select * from ""@Z_HR_ORGST"" where ""U_Z_CompCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Company Code Already mapped in Organization Structure...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

            strqry = "Select * from ""@Z_HR_OPOSIN"" where ""U_Z_CompCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Company Code Already mapped in Position...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

        ElseIf aModule = "FUNC" Then
            strqry = "Select * from ""@Z_HR_ORGST"" where ""U_Z_FuncCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Function Code Already mapped in Organization Structure...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_OPOSIN"" where ""U_Z_DivCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Function Code Already mapped in Position...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

        ElseIf aModule = "UNIT" Then
            strqry = "Select * from ""@Z_HR_ORGST"" where ""U_Z_UnitCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Unit Code Already mapped in Organization Structure...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "LOC" Then
            strqry = "Select * from ""@Z_HR_ORGST"" where ""U_Z_LocCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Location Code Already mapped in Organization Structure...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

        ElseIf aModule = "ORG" Then
            'strqry = "Select * from ""@Z_HR_OPOSCO"" where ""U_Z_OrgCode""='" & aCode & "'"
            'oTemp.DoQuery(strqry)
            'If oTemp.RecordCount > 0 Then
            '    oApplication.Utilities.Message("Organizational Code Already mapped in Job Screen...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return True
            'End If
            strqry = "Select * from OHEM where ""U_Z_HR_OrgstCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Organizational Code Already mapped in Employee Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            'strqry = "Select * from ""@Z_HR_OPOSIN"" where ""U_Z_OrgCode""='" & aCode & "'"
            'oTemp.DoQuery(strqry)
            'If oTemp.RecordCount > 0 Then
            '    Return True
            'End If


        ElseIf aModule = "BUSINESS" Then
            strqry = "Select * from ""@Z_HR_SEAPP1"" where ""U_Z_BussCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Business Objective Code already mapped in Appraisal Business Objective....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_DEMA1"" where ""U_Z_BussCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Business Objective Code already mapped in Department Business Objective....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_COUR1"" where ""U_Z_BussCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                Return True
            End If
        ElseIf aModule = "PEOBJCAT" Then
            strqry = "Select * from ""@Z_HR_OPEOB"" where ""U_Z_PeoCategory""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Personal Category Code Already mapped in Personel Objectives...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "COMPLEVEL" Then
            strqry = "Select * from ""@Z_HR_RMPREQ3"" where ""U_Z_CompLevel""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Competence Level Code Already mapped in Recruitment Requisition...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_ECOLVL"" where ""U_Z_CompLevel""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Competence Level Code Already mapped in Employee Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If

            strqry = "Select * from ""@Z_HR_POSCO1"" where ""U_Z_CompLevel""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Competence Level Code Already mapped in Job Screen ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "PEOBJ" Then
            strqry = "Select * from ""@Z_HR_PEOBJ1"" where ""U_Z_HRPeoobjCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Personal Objective Code already mapped in Employee master Personal Objectives. ", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_COUR2"" where ""U_Z_PeopleCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                Return True
            End If
            'ElseIf aModule = "INTERVIEWTYPE" Then
            '    strqry = "Select * from ""@Z_HR_OITYP"" where ""U_Z_TypeCode""='" & aCode & "'"
            '    oTemp.DoQuery(strqry)
            '    If oTemp.RecordCount > 0 Then
            '        Return True
            '    End If
        ElseIf aModule = "REJECTIONMASTER" Then
            strqry = "select * from ""@Z_HR_OCRAPP"" where ""U_Z_RejResn""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Rejection Code already mapped in Applicant profile....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "OREJECTIONMASTER" Then
            strqry = "select * from ""@Z_HR_OHEM3"" where ""U_Z_RejReason""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Offer Rejection Code already mapped in Employement offer details....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "SEC" Then
            'strqry = "Select * from ""@Z_HR_OSEC"" where ""U_Z_SecCode""='" & aCode & "'"
            'oTemp.DoQuery(strqry)
            'If oTemp.RecordCount > 0 Then
            '    oApplication.Utilities.Message("Section Code Already Exits...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return True
            'End If
            strqry = "Select * from ""@Z_HR_ORGST"" where ""U_Z_SecCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Section Code Already mapped in Organizational Structure...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "RSTA" Then
            strqry = "Select * from ""@Z_HR_ORST"" where ""U_Z_StaCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                Return True
            End If
        ElseIf aModule = "COMPOBJ" Then
            strqry = "Select * from ""@Z_HR_COUR3"" where ""U_Z_CompCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Competence Code Already mapped in Course Master...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_POSCO1"" where ""U_Z_CompCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Competence Code Already mapped in Job Screen...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        ElseIf aModule = "RATE" Then
            strqry = "Select * from ""@Z_HR_SEAPP1"" where ""U_Z_SelfRaCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Rating Code Already mapped in Self Appraisal Rating...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_SEAPP2"" where ""U_Z_MgrRaCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Rating Code Already mapped in First Level Approval Rating...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
            strqry = "Select * from ""@Z_HR_SEAPP3"" where ""U_Z_SMRaCode""='" & aCode & "'"
            oTemp.DoQuery(strqry)
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Rating Code Already mapped in Second Level Approval Rating...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return True
            End If
        End If

        Return False
    End Function
#End Region
#Region "Create Cost Center"
    Public Function createCostCenter(ByVal aCode As String, ByVal aDesc As String, ByVal aStartDate As Date) As String
        Dim oCmpSrv As SAPbobsCOM.CompanyService
        oCmpSrv = oApplication.Company.GetCompanyService()

        Dim oPCService As SAPbobsCOM.IProfitCentersService
        Dim oPC As SAPbobsCOM.IProfitCenter
        Dim oPCParams As SAPbobsCOM.IProfitCenterParams
        Dim oPCsParams As SAPbobsCOM.IProfitCentersParams

        oPCService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProfitCentersService)
        oPCParams = oPCService.GetDataInterface(SAPbobsCOM.DimensionsServiceDataInterfaces.dsDimensionParams)

        oPC = oPCService.GetDataInterface(SAPbobsCOM.ProfitCentersServiceDataInterfaces.pcsProfitCenter)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select * from OPRC where ""PrcCode""='" & aCode & "' and ""DimCode""=1")
        Dim oStr As String = ""
        If oTest.RecordCount <= 0 Then
            oPC.CenterCode = aCode
            oPC.CenterName = aDesc
            oPC.InWhichDimension = 1
            oPC.Effectivefrom = aStartDate ' Now.Date
            oPCParams = oPCService.AddProfitCenter(oPC)
            oStr = oPCParams.CenterCode
        Else
            oStr = oTest.Fields.Item("PrcCode").Value
        End If

        Return oStr
        'If oStr <> "" Then
        '    Return True
        'Else
        '    Return False
        'End If



    End Function
#End Region

#Region "Convert Number to Arabic"


    Public Shared Function SFormatNumber(ByVal X As Double) As String
        Dim Letter1, Letter2, Letter3, Letter4, Letter5, Letter6 As String
        Dim c As String = Format(Math.Floor(X), "000000000000")
        Dim C1 As Double = Val(Mid(c, 12, 1))
        Select Case C1
            Case Is = 1 : Letter1 = "واحد"
            Case Is = 2 : Letter1 = "اثنان"
            Case Is = 3 : Letter1 = "ثلاثة"
            Case Is = 4 : Letter1 = "اربعة"
            Case Is = 5 : Letter1 = "خمسة"
            Case Is = 6 : Letter1 = "ستة"
            Case Is = 7 : Letter1 = "سبعة"
            Case Is = 8 : Letter1 = "ثمانية"
            Case Is = 9 : Letter1 = "تسعة"
        End Select


        Dim C2 As Double = Val(Mid(c, 11, 1))
        Select Case C2
            Case Is = 1 : Letter2 = "عشر"
            Case Is = 2 : Letter2 = "عشرون"
            Case Is = 3 : Letter2 = "ثلاثون"
            Case Is = 4 : Letter2 = "اربعون"
            Case Is = 5 : Letter2 = "خمسون"
            Case Is = 6 : Letter2 = "ستون"
            Case Is = 7 : Letter2 = "سبعون"
            Case Is = 8 : Letter2 = "ثمانون"
            Case Is = 9 : Letter2 = "تسعون"
        End Select


        If Letter1 <> "" And C2 > 1 Then Letter2 = Letter1 + " و" + Letter2
        If Letter2 = "" Or Letter2 Is Nothing Then
            Letter2 = Letter1
        End If
        If C1 = 0 And C2 = 1 Then Letter2 = Letter2 + "ة"
        If C1 = 1 And C2 = 1 Then Letter2 = "احدى عشر"
        If C1 = 2 And C2 = 1 Then Letter2 = "اثنى عشر"
        If C1 > 2 And C2 = 1 Then Letter2 = Letter1 + " " + Letter2
        Dim C3 As Double = Val(Mid(c, 10, 1))
        Select Case C3
            Case Is = 1 : Letter3 = "مائة"
            Case Is = 2 : Letter3 = "مئتان"
            Case Is > 2 : Letter3 = Left(SFormatNumber(C3), Len(SFormatNumber(C3)) - 1) + "مائة"
        End Select
        If Letter3 <> "" And Letter2 <> "" Then Letter3 = Letter3 + " و" + Letter2
        If Letter3 = "" Then Letter3 = Letter2


        Dim C4 As Double = Val(Mid(c, 7, 3))
        Select Case C4
            Case Is = 1 : Letter4 = "الف"
            Case Is = 2 : Letter4 = "الفان"
            Case 3 To 10 : Letter4 = SFormatNumber(C4) + " آلاف"
            Case Is > 10 : Letter4 = SFormatNumber(C4) + " الف"
        End Select
        If Letter4 <> "" And Letter3 <> "" Then Letter4 = Letter4 + " و" + Letter3
        If Letter4 = "" Then Letter4 = Letter3
        Dim C5 As Double = Val(Mid(c, 4, 3))
        Select Case C5
            Case Is = 1 : Letter5 = "مليون"
            Case Is = 2 : Letter5 = "مليونان"
            Case 3 To 10 : Letter5 = SFormatNumber(C5) + " ملايين"
            Case Is > 10 : Letter5 = SFormatNumber(C5) + " مليون"
        End Select
        If Letter5 <> "" And Letter4 <> "" Then Letter5 = Letter5 + " و" + Letter4
        If Letter5 = "" Then Letter5 = Letter4


        Dim C6 As Double = Val(Mid(c, 1, 3))
        Select Case C6
            Case Is = 1 : Letter6 = "مليار"
            Case Is = 2 : Letter6 = "ملياران"
            Case Is > 2 : Letter6 = SFormatNumber(C6) + " مليار"
        End Select
        If Letter6 <> "" And Letter5 <> "" Then Letter6 = Letter6 + " و" + Letter5
        If Letter6 = "" Then Letter6 = Letter5
        SFormatNumber = Letter6


    End Function

    Public Function getWords(ByVal myNumber As String) As String
        getWords = SpellNumber(myNumber) & Ab
    End Function
    Private Function GetIndex(ByVal S As String) As Byte
        If Val(S) = 1 Then
            GetIndex = 1
        ElseIf Val(S) = 2 Then
            GetIndex = 2
        ElseIf Val(S) >= 3 And Val(S) <= 10 Then
            GetIndex = 3
        Else
            GetIndex = 1

        End If

    End Function
    'Main Function
    Private Function SpellNumber(ByVal MyNumber As String)
        Dim intPound, intPiaster, Temp
        Dim DecimalPlace, Count

        Dim Place(9, 3) As String
        Place(2, 1) = " ÃáÝ "
        Place(2, 2) = " ÃáÝíä "
        Place(2, 3) = " ÂáÇÝ "

        Place(3, 1) = " ãáíæä "
        Place(3, 2) = " ãáíæäíä "
        Place(3, 3) = " ãáÇííä "

        Place(4, 1) = " Èáíæä "
        Place(4, 2) = " Èáíæäíä "
        Place(4, 3) = " ÈáÇííä "

        Place(5, 1) = " ÊÑíáíæä "
        Place(5, 2) = " ÊÑíáíæä "
        Place(5, 3) = " ÊÑíáíæä "

        ' String representation of amount
        MyNumber = Convert.ToString(MyNumber)

        ' Position of decimal place 0 if næÇÍÏ
        DecimalPlace = InStr(MyNumber, ".")
        'Convert intPiaster æ set MyNumber to Ìäíå amount
        If DecimalPlace > 0 Then
            intPiaster = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        Count = 1
        Do While MyNumber <> ""
            Temp = GetHundreds(Right(MyNumber, 3))
            Dim S As String = Right(MyNumber, 3)
            If Count = 1 Then
                If Temp <> "" Then intPound = Temp & Place(Count, GetIndex(S)) & intPound
            Else
                If Val(Right(MyNumber, 3)) <= 2 Then
                    If Temp <> "" Then intPound = Place(Count, GetIndex(S)) & An & intPound
                Else
                    If Temp <> "" Then intPound = Temp & Place(Count, GetIndex(S)) & An & intPound
                End If
            End If
            If Len(MyNumber) > 3 Then
                MyNumber = Left(MyNumber, Len(MyNumber) - 3)
            Else
                MyNumber = ""
            End If
            Count = Count + 1
        Loop
        If Right(intPound, 3) = An Then
            intPound = Left(intPound, Len(intPound) - 3)
        End If
        Select Case intPound
            Case ""
                intPound = "ÕÝÑ ÌäíåÇð"
            Case "æÇÍÏ"
                intPound = "Ìäíå æÇÍÏ"
            Case "ÇËäíä"
                intPound = "ÌäíåÇä"
            Case "ËáÇËÉ"
                intPound = "ËáÇËÉ ÌäíåÇÊ"
            Case "ÇÑÈÚÉ"
                intPound = "ÇÑÈÚÉ ÌäíåÇÊ"
            Case "ÎãÓÉ"
                intPound = "ÎãÓÉ ÌäíåÇÊ"
            Case "ÓÊÉ"
                intPound = "ÓÊÉ ÌäíåÇÊ"
            Case "ÓÈÚÉ"
                intPound = "ÓÈÚÉ ÌäíåÇÊ"
            Case "ËãÇäíÉ"
                intPound = "ËãÇäíÉ ÌäíåÇÊ"
            Case "ÊÓÚÉ"
                intPound = "ÊÓÚÉ ÌäíåÇÊ"
            Case "ÚÔÑÉ"
                intPound = "ÚÔÑÉ ÌäíåÇÊ"
            Case Else
                intPound = intPound & " ÌäíåÇð"
        End Select

        Select Case intPiaster
            Case ""
                intPiaster = ""
            Case "æÇÍÏ"
                intPiaster = " æ æÇÍÏ ÞÑÔ"
            Case Else
                intPiaster = " æ " & intPiaster & " ÞÑÔÇð"
        End Select

        SpellNumber = intPound & intPiaster
    End Function

    'Converts a number from 100-999 into text
    Private Function GetHundreds(ByVal MyNumber As String)
        Dim Result As String

        If Val(MyNumber) = 0 Then Exit Function
        MyNumber = Right("000" & MyNumber, 3)

        'Convert the hundreds place
        If Mid(MyNumber, 1, 1) <> "0" Then
            Dim T As String = GetDigit(Mid(MyNumber, 1, 1))
            If T = "æÇÍÏ" Then
                Result = " ãÆÉ "

            ElseIf T = "ÇËäíä" Then
                Result = " ãÆÊÇ "

            ElseIf T = "ËáÇËÉ" Then
                Result = " ËáÇËãÇÆÉ "

            ElseIf T = "ÇÑÈÚÉ" Then
                Result = " ÇÑÈÚãÇÆÉ "

            ElseIf T = "ÎãÓÉ" Then
                Result = " ÎãÓãÇÆÉ "

            ElseIf T = "ÓÊÉ" Then
                Result = " ÓÊãÇÆÉ "

            ElseIf T = "ÓÈÚÉ" Then
                Result = " ÓÈÚãÇÆÉ "

            ElseIf T = "ËãÇäíÉ" Then
                Result = " ËãÇäãÇÆÉ "

            ElseIf T = "ÊÓÚÉ" Then
                Result = " ÊÓÚãÇÆÉ "
            Else
                Result = T & " ãÆÉ "
            End If
        End If

        'Convert the tens æ æÇÍÏs place
        If Mid(MyNumber, 2, 1) <> "0" Then
            Dim T As String = GetTens(Mid(MyNumber, 2))
            If Result = "" Then
                Result = T
            Else
                Result = Result & "æ " & T
            End If
        ElseIf Mid(MyNumber, 3, 1) <> "0" Then
            Dim T As String = GetDigit(Mid(MyNumber, 3))
            If Result = "" Then
                Result = T
            Else
                Result = Result & "æ " & T
            End If
        End If

        GetHundreds = Result
    End Function

    'Converts a number from 10 to 99 into text
    Private Function GetTens(ByVal TensText As String) As String
        Dim Result As String

        Result = "" 'null out the temporary function value
        If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19
            Select Case Val(TensText)
                Case 10 : Result = "ÚÔÑÉ"
                Case 11 : Result = "ÇÍÏ ÚÔÑ"
                Case 12 : Result = "ÇËäÇ ÚÔÑ"
                Case 13 : Result = "ËáÇËÉ ÚÔÑ"
                Case 14 : Result = "ÃÑÈÚÉ ÚÔÑ"
                Case 15 : Result = "ÎãÓÉ ÚÔÑ"
                Case 16 : Result = "ÓÊÉ ÚÔÑ"
                Case 17 : Result = "ÓÈÚÉ ÚÔÑ"
                Case 18 : Result = "ËãÇäíÉ ÚÔÑ"
                Case 19 : Result = "ÊÓÚÉ ÚÔÑ"
                Case Else
                    Result = ""
            End Select
        Else ' If value between 20-99
            Select Case Val(Left(TensText, 1))
                Case 2 : Result = "ÚÔÑæä "
                Case 3 : Result = "ËáÇËæä "
                Case 4 : Result = "ÃÑÈÚæä "
                Case 5 : Result = "ÎãÓæä "
                Case 6 : Result = "ÓÊæä "
                Case 7 : Result = "ÓÈÚæä "
                Case 8 : Result = "ËãÇäæä "
                Case 9 : Result = "ÊÓÚæä "
                Case Else
            End Select
            If GetDigit(Right(TensText, 1)) = "" Then
                Result = GetDigit(Right(TensText, 1)) & Result
            Else
                Result = GetDigit(Right(TensText, 1)) & " æ " & Result

            End If
        End If
        GetTens = Result
    End Function

    'Converts a number from 1 to 9 into text
    Private Function GetDigit(ByVal Digit As String) As String
        Select Case Val(Digit)
            Case 1 : GetDigit = "æÇÍÏ"
            Case 2 : GetDigit = "ÇËäíä"
            Case 3 : GetDigit = "ËáÇËÉ"
            Case 4 : GetDigit = "ÇÑÈÚÉ"
            Case 5 : GetDigit = "ÎãÓÉ"
            Case 6 : GetDigit = "ÓÊÉ"
            Case 7 : GetDigit = "ÓÈÚÉ"
            Case 8 : GetDigit = "ËãÇäíÉ"
            Case 9 : GetDigit = "ÊÓÚÉ"
            Case Else : GetDigit = ""
        End Select
    End Function
#End Region

#Region "Get item Codes from marketing Documents"
    Public Function getItemCodesFromDocuments(ByVal aForm As SAPbouiCOM.Form) As String
        Dim strDocumentItemCodes As String = ""
        Dim strItem As String
        Dim oMatrix As SAPbouiCOM.Matrix
        oMatrix = aForm.Items.Item("38").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            strItem = getMatrixValues(oMatrix, "1", intRow)
            If strItem <> "" Then
                If strDocumentItemCodes = "" Then
                    strDocumentItemCodes = "'" & strItem & "'"
                Else
                    strDocumentItemCodes = strDocumentItemCodes & ",'" & strItem & "'"
                End If
            End If
        Next
        strDocumentItemCodes = "(" & strDocumentItemCodes & ")"
        strSPItemCode = strDocumentItemCodes
        Return strDocumentItemCodes
    End Function
#End Region

#Region "GetLoggedInUser"
    Public Function GetLoggedUserName() As String
        Dim strUser As String
        Dim ouserRs As SAPbobsCOM.Recordset
        ouserRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strUser = oApplication.Company.UserName
        '  ouserRs.DoQuery("Select * from OUSR where Userid='" & strUser & "'")
        Return strUser

    End Function
#End Region

#Region "Display Condition Records"

    Private Function CheckRecordExists(ByVal aCOn As String, ByVal aCon1 As String, ByVal aCon2 As String) As Boolean
        Dim otemprec As SAPbobsCOM.Recordset
        otemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemprec.DoQuery(aCOn)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        otemprec.DoQuery(aCon1)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        otemprec.DoQuery(aCon2)
        If otemprec.RecordCount > 0 Then
            Return True
        End If
        Return False
    End Function




#End Region


#Region "Update first level Item"
    Public Sub UpdateFirstlevelItems()
        Dim oFirstRs, oSecondRS As SAPbobsCOM.Recordset
        Dim strMainCode, strmastercode, strItemcode, strFirstsql, strsecondsql, strsql As String
        oFirstRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oSecondRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.SBO_Application.MessageBox("Do you want to set the first level item?", , "Yes", "No") = 2 Then
            Exit Sub
        End If
        strFirstsql = "Select isnull(U_MainCode,''),count(*) from OITM where isnull(U_Maincode,'')<>'' and isnull(U_Mastercode,'')='' group by isnull(U_MainCode,'')"
        oFirstRs.DoQuery(strFirstsql)
        For intRow As Integer = 0 To oFirstRs.RecordCount - 1
            oApplication.Utilities.Message("Processing....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strMainCode = oFirstRs.Fields.Item(0).Value
            strsecondsql = "select ItemCode from OITM where isnull(U_Mastercode,'')='" & strMainCode & "'"
            oSecondRS.DoQuery(strsecondsql)
            If oSecondRS.RecordCount <= 0 Then
                strsecondsql = "select ItemCode from OITM where isnull(U_Maincode,'')='" & strMainCode & "'"
                oSecondRS.DoQuery(strsecondsql)
                If oSecondRS.RecordCount > 0 Then
                    oSecondRS.MoveFirst()
                    strItemcode = oSecondRS.Fields.Item("ItemCode").Value
                    strsql = "Update OITM set U_Mastercode=U_Maincode,U_Maincode=Itemcode where Itemcode='" & strItemcode & "'"
                    oSecondRS.DoQuery(strsql)
                End If
            End If
            oFirstRs.MoveNext()
        Next
        oApplication.SBO_Application.MessageBox("Operation completed succesfully")
        oApplication.Utilities.Message("Operation completed successfully....", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
    End Sub
#End Region





#Region "Fill ComboBoxValues"
    Public Sub FillComboBoxColumn(ByVal aCombo As SAPbouiCOM.ComboBoxColumn, ByVal sql As String)
        Dim oComborec As SAPbobsCOM.Recordset
        oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oComborec.DoQuery(sql)
        For introw As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            aCombo.ValidValues.Remove(introw)
        Next
        aCombo.ValidValues.Add("", "")
        For introw As Integer = 0 To oComborec.RecordCount - 1
            aCombo.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            oComborec.MoveNext()
        Next


    End Sub

    Public Sub FillComboBox(ByVal aCombo As SAPbouiCOM.ComboBox, ByVal sql As String)
        Dim oComborec As SAPbobsCOM.Recordset
        oComborec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oComborec.DoQuery(sql)
        For introw As Integer = aCombo.ValidValues.Count - 1 To 0 Step -1
            Try
                aCombo.ValidValues.Remove(introw, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception
            End Try
        Next
        aCombo.ValidValues.Add("", "")
        For introw As Integer = 0 To oComborec.RecordCount - 1
            Try
                aCombo.ValidValues.Add(oComborec.Fields.Item(0).Value, oComborec.Fields.Item(1).Value)
            Catch ex As Exception
            End Try

            oComborec.MoveNext()
        Next
        aCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
    End Sub
#End Region

#Region "Copy Files"
    Public Sub CopyFilestoCustomers(ByVal aFileName As String, ByVal aLogPath As String)
        Dim otemp As SAPbobsCOM.Recordset
        Dim strFilePath, strDesgfilename, strMessage As String
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from OCRD where cardtype='C' and U_PharmInt = 'Y'")
        strFilePath = "C:\MYDATA"
        For intRow As Integer = 0 To otemp.RecordCount - 1
            strFilePath = strFilePath & "\" & otemp.Fields.Item("CardCode").Value
            If Directory.Exists(strFilePath) Then
            Else
                Directory.CreateDirectory(strFilePath)
            End If
            strDesgfilename = strFilePath & "\PROMFLQ.mfp"
            If File.Exists(strDesgfilename) Then
                File.Delete(strDesgfilename)
            End If
            File.Copy(aFileName, strDesgfilename)
            '  strFilePath = strExportFilePaty
            '    strMessage = "Exported :  File name : " & strDesgfilename
            ''WriteErrorlog(strMessage, aLogPath)
            otemp.MoveNext()
        Next


    End Sub
#End Region

#Region "Check the Company Settings"
    Public Sub CheckCompanySettings()
        Dim otempRs As SAPbobsCOM.Recordset
        otempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otempRs.DoQuery("Select isnull(U_MasExport,'N'),isnull(U_JEExport,'N') from OADM")
        If otempRs.Fields.Item(0).Value = "Y" Then
            blnMasterExport = True
        Else
            blnMasterExport = False
        End If
        If otempRs.Fields.Item(1).Value = "Y" Then
            blnFEExport = True
        Else
            blnFEExport = False
        End If
    End Sub
#End Region

#Region "Add Controls"

    '*****************************************************************
    'Type               : Procedure   
    'Name               : addControls
    'Parameter          : StrCode
    'Return Value       : string
    'Author             : Senthil Kumar B
    'Created Date       : 03-07-2009
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Create Controls in the SAP B1 Screens
    '*****************************************************************
    Public Sub AddControls(ByVal objForm As SAPbouiCOM.Form, ByVal ItemUID As String, ByVal SourceUID As String, ByVal ItemType As SAPbouiCOM.BoFormItemTypes, ByVal position As String, Optional ByVal fromPane As Integer = 1, Optional ByVal toPane As Integer = 1, Optional ByVal linkedUID As String = "", Optional ByVal strCaption As String = "", Optional ByVal aWidth As Double = 0)
        Dim objNewItem, objOldItem As SAPbouiCOM.Item
        Dim ostatic As SAPbouiCOM.StaticText
        Dim oButton As SAPbouiCOM.Button
        Dim oCheckbox As SAPbouiCOM.CheckBox
        objOldItem = objForm.Items.Item(SourceUID)
        objNewItem = objForm.Items.Add(ItemUID, ItemType)
        With objNewItem
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON Then
                .Left = objOldItem.Left - 15
                .Top = objOldItem.Top + 1
                .LinkTo = linkedUID
            Else
                If position.ToUpper = "RIGHT" Then
                    .Left = objOldItem.Left + objOldItem.Width + 5
                    .Top = objOldItem.Top
                    .Height = objOldItem.Height

                ElseIf position.ToUpper = "DOWN" Then
                    .Top = objOldItem.Top + objOldItem.Height + 1
                    .Left = objOldItem.Left
                ElseIf position.ToUpper = "TOP" Then
                    .Top = objOldItem.Top - objOldItem.Height - 3
                    .Left = objOldItem.Left
                End If
            End If
            .FromPane = fromPane
            .ToPane = toPane
            If ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
                .LinkTo = linkedUID
            End If
            ' .ForeColor = 255
        End With
        If (ItemType = SAPbouiCOM.BoFormItemTypes.it_EDIT Or ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC) Then
            objNewItem.Width = objOldItem.Width
        End If
        If ItemType = SAPbouiCOM.BoFormItemTypes.it_BUTTON Then
            If ItemUID = "btnDisplay" Then
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            Else
                objNewItem.Width = objOldItem.Width
                objNewItem.Width = objNewItem.Width + 60
            End If
            oButton = objNewItem.Specific
            oButton.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_STATIC Then
            ostatic = objNewItem.Specific
            ostatic.Caption = strCaption
        ElseIf ItemType = SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX Then
            oCheckbox = objNewItem.Specific
            oCheckbox.Caption = strCaption
            objNewItem.Width = objOldItem.Width + 10
        End If
        If aWidth <> 0 Then
            objNewItem.Width = aWidth
        End If
    End Sub
#End Region

#Region "validate onHandqty"
    Private Function validateOnhand(ByVal aItemCode As String, ByVal aWhs As String, ByVal dblqty As Double) As Boolean
        Dim oTempRs As SAPbobsCOM.Recordset
        Dim dblOnHand As Double
        oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRs.DoQuery("Select * from OITW where itemcode='" & aItemCode & "' and whscode='" & aWhs & "'")
        dblOnHand = 0
        If oTempRs.RecordCount > 0 Then
            dblOnHand = oTempRs.Fields.Item("OnHand").Value - oTempRs.Fields.Item("MinStock").Value
        End If
        If dblOnHand >= dblqty Then
            Return True
        Else
            Return False
        End If
    End Function
#End Region

#Region "Connect to Company"
    Public Sub Connect()
        Dim strCookie As String
        Dim strConnectionContext As String

        Try
            strCookie = oApplication.Company.GetContextCookie
            strConnectionContext = oApplication.SBO_Application.Company.GetConnectionContext(strCookie)

            If oApplication.Company.SetSboLoginContext(strConnectionContext) <> 0 Then
                Throw New Exception("Wrong login credentials.")
            End If

            'Open a connection to company
            If oApplication.Company.Connect <> 0 Then
                Throw New Exception("Cannot connect to company database. ")
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Genral Functions"

#Region "Get MaxCode"
    Public Function getMaxCode(ByVal sTable As String, ByVal sColumn As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(" & sColumn & " AS Numeric)) FROM [" & sTable & "]"
            ExecuteSQL(oRS, strSQL)

            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 1
            End If

            sCode = Format(MaxCode, "00000000")
            Return sCode
        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region "Write into ErrorLog File"
    Public Sub WriteErrorHeader(ByVal apath As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage, sPath As String
        sPath = apath
        aMessage = "FileName : " & apath
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(sPath, True)
        aSw.WriteLine(aMessage)
        aSw.WriteLine(Now.Date.ToShortDateString)
        aSw.Flush()
        aSw.Close()
    End Sub
    Public Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        Else

        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToLocalTime & "-->" & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region

    Public Function CheckSurcharge(ByVal aCardcode As String, ByVal aFieldName As String) As Boolean
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select isnull(" & aFieldName & ",'N') from OCRD where cardcode='" & aCardcode & "'")
        If otemp.Fields.Item(0).Value.ToString.ToUpper() = "YES" Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetLocalCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select ""Maincurncy"" from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function GetSystemCurrency() As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select ""SysCurrncy"" from OADM"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

    Public Function getBPCurrency(ByVal strCardcode As String) As String
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strSQL, strSegmentation As String
        strSQL = "Select ""Currency"" from OCRD where ""Cardcode""='" & strCardcode & "'"
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTemp.DoQuery(strSQL)
        strSegmentation = oTemp.Fields.Item(0).Value
        Return strSegmentation
    End Function

#Region "Status Message"
    Public Sub Message(ByVal sMessage As String, ByVal StatusType As SAPbouiCOM.BoStatusBarMessageType)
        oApplication.SBO_Application.StatusBar.SetText(sMessage, SAPbouiCOM.BoMessageTime.bmt_Short, StatusType)
    End Sub
#End Region

    Public Function GetDateTimeValue(ByVal DateString As String) As DateTime
        Dim objBridge As SAPbobsCOM.SBObob
        objBridge = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
        Return objBridge.Format_StringToDate(DateString).Fields.Item(0).Value
    End Function

#Region "SetDatabind"
    Public Sub setUserDatabind(ByVal aForm As SAPbouiCOM.Form, ByVal UID As String, ByVal strDBID As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aForm.Items.Item(UID).Specific
        objEdit.DataBind.SetBound(True, "", strDBID)
    End Sub
#End Region

#Region "Add Choose from List"
    Public Sub AddChooseFromList(ByVal FormUID As String, ByVal CFL_Text As String, ByVal CFL_Button As String, _
                                        ByVal ObjectType As SAPbouiCOM.BoLinkedObject, _
                                            Optional ByVal AliasName As String = "", Optional ByVal CondVal As String = "", _
                                                    Optional ByVal Operation As SAPbouiCOM.BoConditionOperation = SAPbouiCOM.BoConditionOperation.co_EQUAL)

        Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
        Dim oCons As SAPbouiCOM.Conditions
        Dim oCon As SAPbouiCOM.Condition
        Dim oCFL As SAPbouiCOM.ChooseFromList
        Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
        Try
            oCFLs = oApplication.SBO_Application.Forms.Item(FormUID).ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject( _
                                    SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            If ObjectType = SAPbouiCOM.BoLinkedObject.lf_Items Then
                oCFLCreationParams.MultiSelection = True
            Else
                oCFLCreationParams.MultiSelection = False
            End If

            oCFLCreationParams.ObjectType = ObjectType
            oCFLCreationParams.UniqueID = CFL_Text

            oCFL = oCFLs.Add(oCFLCreationParams)

            ' Adding Conditions to CFL1

            oCons = oCFL.GetConditions()

            If Not AliasName = "" Then
                oCon = oCons.Add()
                oCon.Alias = AliasName
                oCon.Operation = Operation
                oCon.CondVal = CondVal
                oCFL.SetConditions(oCons)
            End If

            oCFLCreationParams.UniqueID = CFL_Button
            oCFL = oCFLs.Add(oCFLCreationParams)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Linked Object Type"
    Public Function getLinkedObjectType(ByVal Type As SAPbouiCOM.BoLinkedObject) As String
        Return CType(Type, String)
    End Function

#End Region

#Region "Execute Query"
    Public Sub ExecuteSQL(ByRef oRecordSet As SAPbobsCOM.Recordset, ByVal SQL As String)
        Try
            If oRecordSet Is Nothing Then
                oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            End If

            oRecordSet.DoQuery(SQL)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Get Application path"
    Public Function getApplicationPath() As String

        Return Application.StartupPath.Trim

        'Return IO.Directory.GetParent(Application.StartupPath).ToString
    End Function
#End Region

#Region "Date Manipulation"

#Region "Convert SBO Date to System Date"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	ConvertStrToDate
    'Parameter          	:   ByVal oDate As String, ByVal strFormat As String
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	07/12/05
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To convert Date according to current culture info
    '********************************************************************
    Public Function ConvertStrToDate(ByVal strDate As String, ByVal strFormat As String) As DateTime
        Try
            Dim oDate As DateTime
            Dim ci As New System.Globalization.CultureInfo("en-GB", False)
            Dim newCi As System.Globalization.CultureInfo = CType(ci.Clone(), System.Globalization.CultureInfo)

            System.Threading.Thread.CurrentThread.CurrentCulture = newCi
            oDate = oDate.ParseExact(strDate, strFormat, ci.DateTimeFormat)

            Return oDate
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region

#Region " Get SBO Date Format in String (ddmmyyyy)"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	StrSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(ddmmyy value) as applicable to SBO
    '********************************************************************
    Public Function StrSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String, GetDateFormat As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yy"
                Case 1
                    GetDateFormat = "dd" & DateSep & "MM" & DateSep & "yyyy"
                Case 2
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yy"
                Case 3
                    GetDateFormat = "MM" & DateSep & "dd" & DateSep & "yyyy"
                Case 4
                    GetDateFormat = "yyyy" & DateSep & "dd" & DateSep & "MM"
                Case 5
                    GetDateFormat = "dd" & DateSep & "MMM" & DateSep & "yyyy"
            End Select
            Return GetDateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Get SBO date Format in Number"
    '********************************************************************
    'Type		            :   Public Procedure     
    'Name               	:	IntSBODateFormat
    'Parameter          	:   none
    'Return Value       	:	
    'Author             	:	Manu
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To get date Format(integer value) as applicable to SBO
    '********************************************************************
    Public Function NumSBODateFormat() As String
        Try
            Dim rsDate As SAPbobsCOM.Recordset
            Dim strsql As String
            Dim DateSep As Char

            strsql = "Select DateFormat,DateSep from OADM"
            rsDate = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsDate.DoQuery(strsql)
            DateSep = rsDate.Fields.Item(1).Value

            Select Case rsDate.Fields.Item(0).Value
                Case 0
                    NumSBODateFormat = 3
                Case 1
                    NumSBODateFormat = 103
                Case 2
                    NumSBODateFormat = 1
                Case 3
                    NumSBODateFormat = 120
                Case 4
                    NumSBODateFormat = 126
                Case 5
                    NumSBODateFormat = 130
            End Select
            Return NumSBODateFormat

        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#End Region

#Region "Get Rental Period"
    Public Function getRentalDays(ByVal Date1 As String, ByVal Date2 As String, ByVal IsWeekDaysBilling As Boolean) As Integer
        Dim TotalDays, TotalDaysincSat, TotalBillableDays As Integer
        Dim TotalWeekEnds As Integer
        Dim StartDate As Date
        Dim EndDate As Date
        Dim oRecordset As SAPbobsCOM.Recordset

        StartDate = CType(Date1.Insert(4, "/").Insert(7, "/"), Date)
        EndDate = CType(Date2.Insert(4, "/").Insert(7, "/"), Date)

        TotalDays = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekDaysBilling Then
            strSQL = " select dbo.WeekDays('" & Date1 & "','" & Date2 & "')"
            oApplication.Utilities.ExecuteSQL(oRecordset, strSQL)
            If oRecordset.RecordCount > 0 Then
                TotalBillableDays = oRecordset.Fields.Item(0).Value
            End If
            Return TotalBillableDays
        Else
            Return TotalDays + 1
        End If

    End Function

#Region "Assign Fright Expances Details"

    Public Function validateSurchargerequeired(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strCardCode, strDocType As String
        Dim oCombo As SAPbouiCOM.ComboBox
        oCombo = aForm.Items.Item("3").Specific
        strDocType = oCombo.Selected.Value
        If strDocType <> "I" Then
            Return False
        End If
        strCardCode = getEdittextvalue(aForm, "4")

        If CheckSurcharge(strCardCode, "U_Z_SHAK_FLAG") = True Then
            Return True
        ElseIf CheckSurcharge(strCardCode, "U_Z_APOT_FLAG") = True Then
            Return True
        Else
            Return False
        End If
        Return False

    End Function

    Public Sub CalculateSurcharges(ByVal aForm As SAPbouiCOM.Form)
        Dim oTempForm1, oTempForm2 As SAPbouiCOM.Form
        Dim Frtmtrx As SAPbouiCOM.Matrix
        Dim strCardCode, strCurrencyQuery, strCurrency, strFrightValue, strDate, strSurName, strRemarks, strDocBeftotal, strDiscount, strVatcode As String
        Dim dblVatAmount, dblFixedAmount, dblSurPer, dblVatPer, dbLDocDefTotal, dblDiscount As Double
        Dim dtDocdate As Date
        Dim otempRec, otemprs, oSurRecord As SAPbobsCOM.Recordset
        Dim w As Integer
        Dim oCombo As SAPbouiCOM.ComboBox

        If aForm.Type = frm_ARInvoice Or aForm.Type = frm_ARCreditNote Then
            If validateSurchargerequeired(aForm) = False Then
                Exit Sub
            End If
            aSourceForm = aForm
            oCombo = aSourceForm.Items.Item("70").Specific
            strCurrency = oCombo.Selected.Value
            strCardCode = getEdittextvalue(aForm, "4")
            strDocBeftotal = getEdittextvalue(aSourceForm, "22")
            strDiscount = getEdittextvalue(aSourceForm, "42")
            If strDocBeftotal = "" And strDiscount = "" Then
                Exit Sub
            End If
            otemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strCurrencyQuery = ""
            Select Case strCurrency
                Case "C"
                    strCurrencyQuery = "Select Currency from OCRD where Cardcode='" & strCardCode & "'"
                Case "L"
                    strCurrencyQuery = "Select MainCurncy from OADM"
                Case "S"
                    strCurrencyQuery = "Select SysCurncy from OADM"
            End Select
            If strCurrencyQuery <> "" Then
                otemprs.DoQuery(strCurrencyQuery)
                strCurrency = otemprs.Fields.Item(0).Value
            Else
                strCurrency = ""
            End If

            If strDocBeftotal.Length > 3 Then
                If strCurrency <> "##" Then
                    strDocBeftotal = strDocBeftotal.Replace(strCurrency, "")
                Else
                    strDocBeftotal = strDocBeftotal.Substring(3)
                End If
            End If

            If strDiscount.Length > 3 Then
                If strCurrency <> "##" Then
                    strDiscount = strDiscount.Replace(strCurrency, "")
                Else
                    strDiscount = strDiscount.Substring(3)
                End If
            End If
            If strDocBeftotal <> "" Then
                dbLDocDefTotal = CDbl(strDocBeftotal)
            Else
                dbLDocDefTotal = 0
            End If
            If strDiscount <> "" Then
                dblDiscount = CDbl(strDiscount)
            Else
                dblDiscount = 0
            End If


            strDate = getEdittextvalue(aForm, "10")
            dtDocdate = GetDateTimeValue(strDate)
            otempRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oSurRecord = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempForm1 = aForm
            'oTempForm1.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oTempForm2 = oApplication.SBO_Application.Forms.GetForm("3007", 1)  '// Freight Screen
            oTempForm2 = oApplication.SBO_Application.Forms.ActiveForm()
            If oTempForm2.Type <> 3007 Then
                Exit Sub
            End If
            Dim strFrightName As String
            Try
                oTempForm2.Freeze(True)
                Dim strSQL As String
                If CheckSurcharge(strCardCode, "U_Z_SHAK_FLAG") = True Then
                    strSQL = "Select * from [@Z_SURCHARGES] where U_Z_SUR_BPNM='U_Z_SHAK_FLAG' and '" & dtDocdate.ToString("yyyy-MM-dd") & " '  between U_Z_SUR_FRMDATE and isnull(U_Z_SUR_TODATE,dateadd(m,40,getdate())) order by U_Z_SUR_FRMDATE Desc, U_Z_SUR_TODATE Desc"
                    oSurRecord.DoQuery(strSQL)
                    dblFixedAmount = 0
                    dblVatAmount = 0
                    strRemarks = ""
                    dblVatPer = 0
                    dblSurPer = 0
                    If oSurRecord.RecordCount > 0 Then
                        Frtmtrx = oTempForm2.Items.Item("3").Specific
                        w = 1
                        strVatcode = oSurRecord.Fields.Item("U_Z_SUR_VAT").Value
                        dblVatPer = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value
                        dblSurPer = oSurRecord.Fields.Item("U_Z_SUR_PER").Value
                        strRemarks = oSurRecord.Fields.Item("U_Z_SUR_REM").Value
                        dblFixedAmount = ((dbLDocDefTotal - dblDiscount) * dblSurPer) / 100
                        dblVatAmount = dblFixedAmount * dblVatPer / 100
                        strSurName = oSurRecord.Fields.Item("U_Z_SUR_NAME").Value
                        strFrightName = ""
                        While w <= Frtmtrx.RowCount
                            Try
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description
                            Catch ex As Exception
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.value
                                otempRec.DoQuery("SELECT Expnscode,Expnsname  FROM OEXD T0 where ExpnsCode=" & strFrightName)
                                If otempRec.RecordCount > 0 Then
                                    strFrightName = otempRec.Fields.Item(1).Value
                                End If
                            End Try


                            If strFrightName.ToUpper = strSurName.ToUpper Then '//AD
                                Frtmtrx.Columns.Item("2").Cells.Item(w).Specific.value = strRemarks
                                Frtmtrx.Columns.Item("3").Cells.Item(w).Specific.value = dblFixedAmount
                                Try
                                    oCombo = Frtmtrx.Columns.Item("11").Cells.Item(w).Specific
                                    oCombo.Select(strVatcode, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    '    Frtmtrx.Columns.Item("12").Cells.Item(w).Specific.value = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value

                                    '    Frtmtrx.Columns.Item("17").Cells.Item(w).Specific.value = strVatcode
                                Catch ex As Exception

                                End Try

                                Exit While
                            End If
                            w = w + 1
                        End While
                    End If
                End If
                dblFixedAmount = 0
                dblVatAmount = 0
                strRemarks = ""
                dblVatPer = 0
                dblSurPer = 0
                If CheckSurcharge(strCardCode, "U_Z_APOT_FLAG") = True Then
                    strSQL = "Select * from [@Z_SURCHARGES] where U_Z_SUR_BPNM='U_Z_APOT_FLAG' and '" & dtDocdate.ToString("yyyy-MM-dd") & " '  between U_Z_SUR_FRMDATE and isnull(U_Z_SUR_TODATE,dateadd(m,40,getdate())) order by U_Z_SUR_FRMDATE Desc, U_Z_SUR_TODATE desc"
                    oSurRecord.DoQuery(strSQL)
                    If oSurRecord.RecordCount > 0 Then
                        Frtmtrx = oTempForm2.Items.Item("3").Specific
                        strVatcode = oSurRecord.Fields.Item("U_Z_SUR_VAT").Value
                        dblVatPer = oSurRecord.Fields.Item("U_Z_SUR_VATPER").Value
                        dblSurPer = oSurRecord.Fields.Item("U_Z_SUR_PER").Value
                        strRemarks = oSurRecord.Fields.Item("U_Z_SUR_REM").Value
                        dblFixedAmount = ((dbLDocDefTotal - dblDiscount) * dblSurPer) / 100
                        dblVatAmount = dblFixedAmount * dblVatPer / 100
                        w = 1
                        strSurName = oSurRecord.Fields.Item("U_Z_SUR_NAME").Value
                        strFrightName = ""
                        While w <= Frtmtrx.RowCount
                            '  If Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description = strSurName Then '//AD
                            Try
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.selected.description
                            Catch ex As Exception
                                strFrightName = Frtmtrx.Columns.Item("1").Cells.Item(w).Specific.value
                                otempRec.DoQuery("SELECT Expnscode,Expnsname  FROM OEXD T0 where ExpnsCode=" & strFrightName)
                                If otempRec.RecordCount > 0 Then
                                    strFrightName = otempRec.Fields.Item(1).Value
                                End If
                            End Try
                            If strFrightName.ToUpper = strSurName.ToUpper Then '//AD
                                Frtmtrx.Columns.Item("2").Cells.Item(w).Specific.value = strRemarks
                                Frtmtrx.Columns.Item("3").Cells.Item(w).Specific.value = dblFixedAmount
                                Try
                                    oCombo = Frtmtrx.Columns.Item("11").Cells.Item(w).Specific
                                    oCombo.Select(strVatcode, SAPbouiCOM.BoSearchKey.psk_ByValue)

                                    '   Frtmtrx.Columns.Item("17").Cells.Item(w).Specific.value = strVatcode
                                Catch ex As Exception

                                End Try

                                Exit While
                            End If
                            w = w + 1
                        End While
                    End If
                End If
                oTempForm2.Freeze(False)
                If blnDocumentItem = True Then
                    If oTempForm2.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    Else
                        oTempForm2.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    End If
                End If
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                oTempForm2.Freeze(False)
            End Try
            blnDocumentItem = False
        End If
    End Sub
#End Region


#Region "Check Condition Type"
    Public Function CheckConditionType(ByVal aCode As String) As String
        Dim oCheckRs As SAPbobsCOM.Recordset
        oCheckRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oCheckRs.DoQuery("Select isnull(U_Z_COND_TYPE,'') from [@Z_CONDITIONS] where U_Z_COND_NAME='" & aCode & "'")
        Return (oCheckRs.Fields.Item(0).Value)
    End Function
#End Region

#Region "calculate Discount"

#Region "Check COndition Status"
    Private Function CheckConditionStatus(ByVal aCode As String) As Boolean
        Dim Temprec As SAPbobsCOM.Recordset
        Temprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Temprec.DoQuery("Select * from [@Z_CONDITIONS] where U_Z_COND_CODE='" & aCode & "'")
        If Temprec.RecordCount > 0 Then
            If Temprec.Fields.Item("U_Z_COND_STATUS").Value = "Y" Then
                Return True
            Else
                Return False
            End If
        End If
        Return False
    End Function
#End Region

    Private Function getQuantityDiscount(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aCardCode As String, ByVal aItemGroup As String, ByVal aDate As Date, ByVal LineQty As Double) As Double
        Dim oQtyRS, oItemRS As SAPbobsCOM.Recordset
        Dim strQtyRS, strItemCode, strLineItemGroup, strMainItemGroup As String
        Dim dblDis, dblScalqty, dblLineQty, dblScaleToQty As Double
        oItemRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblLineQty = 0

        strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        If oQtyRS.RecordCount > 0 Then
            strMainItemGroup = oQtyRS.Fields.Item(0).Value
        Else
            Return dblLineQty
        End If
        For IntLoop As Integer = 1 To aMatrix.RowCount
            strItemCode = getMatrixValues(aMatrix, "1", IntLoop)
            If strItemCode <> "" Then
                'oItemRS.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
                strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & strItemCode & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
                oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'oQtyRS.DoQuery(strQtyRS)
                oItemRS.DoQuery(strQtyRS)
                If oItemRS.RecordCount > 0 Then
                    strLineItemGroup = oItemRS.Fields.Item(0).Value
                    If strMainItemGroup = strLineItemGroup Then
                        dblLineQty = dblLineQty + getDocumentQuantity(getMatrixValues(aMatrix, "11", IntLoop))
                    End If
                End If
            End If
        Next
        LineQty = dblLineQty
        strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET],isnull(T0.[U_Z_QTY_TO_SCALE],0) 'ScaleTo' FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        ' strQtyRS = "SELECT T0.[U_Z_DISC_LINK], T0.[U_Z_DISC_CCODE], T0.[U_Z_DISC_ICODE], T0.[U_Z_QTY_FROM_SCALE], T0.[U_Z_QTY_SCALE_DIS], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET],isnull(T0.[U_Z_QTY_TO_SCALE]," & LineQty & ") 'ScaleTo' FROM [dbo].[@Z_QTY_DISCOUNT]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_CCODE]='" & aCardCode & "' and T0.[U_Z_DISC_ICODE]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code) Desc"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        dblDis = 0
        If oQtyRS.RecordCount > 0 Then
            'oQtyRS.MoveLast()
            For intRow As Integer = 0 To oQtyRS.RecordCount - 1
                If CheckConditionStatus(oQtyRS.Fields.Item(0).Value) = True Then
                    dblScalqty = oQtyRS.Fields.Item("U_Z_QTY_FROM_SCALE").Value
                    '  dblScaleToQty = oQtyRS.Fields.Item("ScaleTo").Value
                    'If dblScaleToQty = 0 Then
                    '    dblScaleToQty = LineQty
                    'End If
                    If dblScalqty <= LineQty Then
                        ' If LineQty <= dblScaleToQty And LineQty >= dblScalqty Then
                        dblDis = oQtyRS.Fields.Item("U_Z_QTY_SCALE_DIS").Value
                        Exit For
                    End If
                End If
                oQtyRS.MoveNext()
            Next
        End If
        If dblDis = 0 Then
            oQtyRS.DoQuery(strQtyRS)
            If oQtyRS.RecordCount > 0 Then
                oQtyRS.MoveLast()
                'dblDis = oQtyRS.Fields.Item("U_Z_QTY_SCALE_DIS").Value
                dblDis = 0
            End If
        End If
        Return dblDis
    End Function

    Private Function getConditiongroup(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal aCardCode As String, ByVal aItemGroup As String, ByVal aDate As Date) As Double
        Dim oQtyRS, oItemRS As SAPbobsCOM.Recordset
        Dim strQtyRS, strItemCode, strLineItemGroup, strMainItemGroup As String
        Dim dblDis, dblScalqty, dblLineQty, dblScaleToQty As Double
        oItemRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        dblLineQty = 0
        oItemRS.DoQuery("Select * from OCRD where Cardcode='" & aCardCode & "'")
        aCardCode = oItemRS.Fields.Item("GroupCode").Value
        strQtyRS = "SELECT T0.[U_Z_DISC_LINK],  T0.[U_Z_DISC_PERC], T0.[U_Z_DISC_DATEF], T0.[U_Z_DISC_DATET] FROM [dbo].[@Z_DISC_BP_ITM_GROUP]  T0 "
        strQtyRS = strQtyRS & " where T0.[U_Z_DISC_BP_GROUP]='" & aCardCode & "' and T0.[U_Z_DISC_ITM_GROUP]='" & aItemGroup & "' and '" & aDate.ToString("yyyy-MM-dd") & "' between T0.[U_Z_DISC_DATEF] and T0.[U_Z_DISC_DATET] order by T0.[U_Z_DISC_DATEF] desc, T0.[U_Z_DISC_DATET] ,convert(numeric,Code)"
        oQtyRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oQtyRS.DoQuery(strQtyRS)
        dblDis = 0
        If oQtyRS.RecordCount > 0 Then
            'oQtyRS.MoveLast()
            For intRow As Integer = 0 To oQtyRS.RecordCount - 1
                If CheckConditionStatus(oQtyRS.Fields.Item(0).Value) = True Then
                    dblDis = dblDis + oQtyRS.Fields.Item("U_Z_DISC_PERC").Value
                End If
                oQtyRS.MoveNext()
            Next
        End If
        Return dblDis
    End Function
    Public Sub CalculateDiscount(ByVal aForm As SAPbouiCOM.Form)
        Dim oMatrix As SAPbouiCOM.Matrix
        Dim oCombo As SAPbouiCOM.ComboBox
        Dim strCardCode, strItemCode, strSQL, strTempQuery, strConditionCode, strConditionQuery, strPostingdate As String
        Dim dtPostingdate As Date
        Dim oItemRs, oConditionGroup, oConditionType, oTempRs As SAPbobsCOM.Recordset
        Try
            Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            aForm.Freeze(True)
            oItemRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oConditionGroup = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oConditionType = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oMatrix = aForm.Items.Item("38").Specific
            strCardCode = getEdittextvalue(aForm, "4")
            strPostingdate = getEdittextvalue(aForm, "10")
            If strPostingdate <> "" Then
                dtPostingdate = GetDateTimeValue(strPostingdate)
            End If
            oCombo = aForm.Items.Item("3").Specific
            strTempQuery = ""

            'Condition Type Discount for Price 

            strTempQuery = "SELECT *  FROM [dbo].[@Z_DISCOUNT_GROUP]  T0 inner join  [dbo].[@Z_CONDITIONS]  "
            ' strTempQuery = strTempQuery & " T1 on T0.U_Z_Disc_Link=T1.U_Z_COND_CODE and T1.U_Z_COND_STATUS='Y' where U_Z_DISC_CCODE='" & strCardCode & "'"
            strTempQuery = strTempQuery & " T1 on T0.U_Z_Disc_Link=T1.U_Z_COND_CODE   where U_Z_DISC_CCODE='" & strCardCode & "'"
            strTempQuery = strTempQuery & "  and ('" & dtPostingdate.ToString("yyyy-MM-dd") & "' between U_Z_DISC_DATEF and U_Z_DISC_DATET )"
            oConditionType.DoQuery(strTempQuery)
            Dim dblCumdiscount, dblDiscount As Double
            If oConditionType.RecordCount > 0 Then
                If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                    For intRow As Integer = 1 To oMatrix.RowCount
                        Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        strItemCode = getMatrixValues(oMatrix, "1", intRow)
                        dblCumdiscount = 0
                        dblDiscount = 0
                        If strItemCode <> "" Then
                            strSQL = ""
                            strSQL = strTempQuery & " and U_Z_DISC_ICODE='" & strItemCode & "'"
                            oTempRs.DoQuery(strSQL)
                            If oTempRs.RecordCount > 0 Then
                                If oTempRs.Fields.Item("U_Z_COND_STATUS").Value = "Y" Then
                                    For intLoop As Integer = 0 To oTempRs.RecordCount - 1
                                        dblDiscount = 0
                                        dblDiscount = oTempRs.Fields.Item("U_Z_DISC_PERC").Value
                                        dblCumdiscount = dblCumdiscount + dblDiscount
                                        oTempRs.MoveNext()
                                    Next
                                End If
                                Dim dbllinelineqty, dblScalediscount As Double
                                SetMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow, dblCumdiscount)
                                dblCumdiscount = dblCumdiscount + dblScalediscount
                                dblCumdiscount = dblCumdiscount
                                SetMatrixValues(oMatrix, "15", intRow, dblCumdiscount.ToString)
                            End If
                        End If
                    Next
                End If
            Else
                For intLoop As Integer = 1 To oMatrix.RowCount
                    strItemCode = getMatrixValues(oMatrix, "1", intLoop)
                    If strItemCode <> "" Then
                        dblCumdiscount = 0
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intLoop, dblCumdiscount)
                    End If
                Next
            End If

            'Condition Group Discount 
            If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = getMatrixValues(oMatrix, "1", intRow)
                    dblCumdiscount = 0
                    dblDiscount = 0
                    If strItemCode <> "" Then
                        strSQL = ""
                        oItemRs.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                        Dim dblScalediscount As Double
                        dblCumdiscount = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow))
                        dblScalediscount = getConditiongroup(oMatrix, strCardCode, oItemRs.Fields.Item("ItmsGrpCod").Value, dtPostingdate)
                        'dblScalediscount = dblCumdiscount + dblScalediscount
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_GROUP", intRow, dblScalediscount)
                    End If
                Next
            End If



            'Condition Type Discount for Scales
            If strCardCode <> "" And oCombo.Selected.Value = "I" And strPostingdate <> "" Then
                For intRow As Integer = 1 To oMatrix.RowCount
                    Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                    strItemCode = getMatrixValues(oMatrix, "1", intRow)
                    dblCumdiscount = 0
                    dblDiscount = 0
                    If strItemCode <> "" Then
                        strSQL = ""
                        oItemRs.DoQuery("Select * from OITM where Itemcode='" & strItemCode & "'")
                        Dim dbllinelineqty, dblScalediscount, dblGroupPrice As Double
                        dbllinelineqty = getDocumentQuantity(getMatrixValues(oMatrix, "11", intRow))
                        dblCumdiscount = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_PRICE", intRow))
                        dblGroupPrice = getDocumentQuantity(getMatrixValues(oMatrix, "U_Z_DISCOUNT_GROUP", intRow))
                        dblScalediscount = getQuantityDiscount(oMatrix, strCardCode, oItemRs.Fields.Item("ItemCode").Value, dtPostingdate, dbllinelineqty)
                        SetMatrixValues(oMatrix, "U_Z_DISCOUNT_SCALE", intRow, dblScalediscount)
                        dblCumdiscount = dblCumdiscount + dblScalediscount + dblGroupPrice
                        If dblCumdiscount <> 0 Then
                            SetMatrixValues(oMatrix, "15", intRow, dblCumdiscount.ToString)
                        Else
                            GetB1Price(strItemCode, strCardCode, oMatrix, intRow)
                        End If
                        'End If
                    End If
                Next
            End If
            aForm.Freeze(False)
            Message("", SAPbouiCOM.BoStatusBarMessageType.smt_None)
        Catch ex As Exception
            aForm.Freeze(False)
            Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End Try

    End Sub
#End Region
    Public Function WorkDays(ByVal dtBegin As Date, ByVal dtEnd As Date) As Long
        Try
            Dim dtFirstSunday As Date
            Dim dtLastSaturday As Date
            Dim lngWorkDays As Long

            ' get first sunday in range
            dtFirstSunday = dtBegin.AddDays((8 - Weekday(dtBegin)) Mod 7)

            ' get last saturday in range
            dtLastSaturday = dtEnd.AddDays(-(Weekday(dtEnd) Mod 7))

            ' get work days between first sunday and last saturday
            lngWorkDays = (((DateDiff(DateInterval.Day, dtFirstSunday, dtLastSaturday)) + 1) / 7) * 5

            ' if first sunday is not begin date
            If dtFirstSunday <> dtBegin Then

                ' assume first sunday is after begin date
                ' add workdays from begin date to first sunday
                lngWorkDays = lngWorkDays + (7 - Weekday(dtBegin))

            End If

            ' if last saturday is not end date
            If dtLastSaturday <> dtEnd Then

                ' assume last saturday is before end date
                ' add workdays from last saturday to end date
                lngWorkDays = lngWorkDays + (Weekday(dtEnd) - 1)

            End If

            WorkDays = lngWorkDays
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Function

#End Region

#Region "Get Price "
    Private Sub GetB1Price(ByVal StrItem As String, ByVal strBP As String, ByVal amatrix As SAPbouiCOM.Matrix, ByVal intRow As Integer)
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItems As SAPbobsCOM.Items
        Dim oRec, oREc1, oRecTemp, oRecDiscount As SAPbobsCOM.Recordset
        Dim strSQL, strSQL1, strDiscount, strBPCod As String
        Dim price, discount As Double
        Dim intFlag As Integer
        Dim intPriceList As Integer
        Dim blnDiscountflag As Boolean
        '  Dim oBP As SAPbobsCOM.BusinessPartners
        Dim objForm As SAPbouiCOM.Form
        ' Dim oRec As SAPbobsCOM.Recordset
        Dim oStatic As SAPbouiCOM.StaticText
        Dim oItem, oItem1 As SAPbouiCOM.Item
        price = 0
        discount = 0
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        oItems = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oREc1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecDiscount = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        intFlag = 0
        blnDiscountflag = False
        If oBP.GetByKey(strBP) Then
            'Find discount in Special Price Table
            strSQL = "SELECT T0.[ItemCode], T0.[CardCode], T0.[Discount], T0.[ListNum] FROM OSPP T0 where T0.Cardcode='" & strBP & "' and T0.ItemCode='" & StrItem & "'"
            oRec.DoQuery(strSQL)
            If oRec.RecordCount > 0 Then
                discount = oRec.Fields.Item(2).Value
                intFlag = 1
                blnDiscountflag = True
            End If
            ' Exit Sub

            'Find discount in Discount Group for given BP
            strSQL = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
            oRec.DoQuery(strSQL)
            If blnDiscountflag = False And oRec.RecordCount > 0 Then
                If Convert.ToDouble(oRec.Fields.Item(1).Value) = 52 Then 'Item Group
                    If oItems.GetByKey(StrItem) Then
                        strSQL1 = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "' and T0.ObjKey=" & oItems.ItemsGroupCode
                        oREc1.DoQuery(strSQL1)
                        If oREc1.RecordCount > 0 Then
                            discount = oREc1.Fields.Item(3).Value
                            intFlag = 1
                            blnDiscountflag = True
                        End If
                    End If
                ElseIf Convert.ToDouble(oRec.Fields.Item(1).Value) = 8 Then ' Item Property
                    Dim strProperty, strD As String
                    strBPCod = "Select dscntrel from ocrd where Cardcode='" & strBP & "'"
                    oRecTemp.DoQuery(strBPCod)
                    If oRecTemp.RecordCount > 0 Then
                        strDiscount = oRecTemp.Fields.Item(0).Value
                        Select Case strDiscount
                            Case "L" ' Lowest
                                strD = "Select Min(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "H" 'Highest
                                strD = "Select Max(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "A" 'Average
                                strD = "Select Avg(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                            Case "S" 'Discount Total
                                strD = "Select Sum(T0.[Discount]) FROM OSPG T0 where T0.Cardcode='" & strBP & "'"
                        End Select
                        oRecDiscount.DoQuery(strD)

                        For IntTemp As Integer = 0 To oRec.RecordCount - 1
                            strProperty = oRec.Fields.Item(2).Value
                            strProperty = "QryGroup" & strProperty
                            strSQL1 = "select " & strProperty & " from OITM where Itemcode='" & StrItem & "'"
                            oREc1.DoQuery(strSQL1)
                            If oREc1.RecordCount > 0 Then
                                If oREc1.Fields.Item(0).Value = "Y" Then
                                    discount = oRecDiscount.Fields.Item(0).Value
                                    intFlag = 1
                                    blnDiscountflag = True
                                    Exit For
                                End If
                            End If
                            oRec.MoveNext()
                        Next
                    End If

                ElseIf Convert.ToDouble(oRec.Fields.Item(1).Value) = 43 Then 'Manufacture
                    strSQL1 = "SELECT T0.[CardCode], T0.[ObjType], T0.[ObjKey], T0.[Discount] FROM OSPG T0 where T0.Cardcode='" & strBP & "' T0.ObjKey=" & oItems.Manufacturer
                    oREc1.DoQuery(strSQL1)
                    If oREc1.RecordCount > 0 Then
                        intFlag = 1
                        discount = oREc1.Fields.Item(3).Value
                        blnDiscountflag = True
                    End If
                End If
            End If

            'Find Discount in Hierarchies for given Item code
            strSQL = "SELECT T0.[ItemCode], T0.[CardCode], T0.[ListNum], T0.[Discount], T0.[FromDate], T0.[ToDate]  FROM SPP1 T0 where   T0.Itemcode='" & StrItem & "' and Getdate() between T0. Fromdate and T0.ToDate "
            oRec.DoQuery(strSQL)
            If blnDiscountflag = False And oRec.RecordCount > 0 Then
                discount = oRec.Fields.Item(3).Value
                intPriceList = Convert.ToInt64(oRec.Fields.Item(2).Value)
                strSQL1 = "SELECT T1.[ItemCode], T1.[PriceList], T1.[Price] FROM OPLN T0  INNER JOIN ITM1 T1 ON T0.ListNum = T1.PriceList where T1.Itemcode='" & StrItem & "' and T1.PriceList=" & intPriceList
                oREc1.DoQuery(strSQL)
                If oREc1.RecordCount > 0 Then
                    price = Convert.ToDouble(oREc1.Fields.Item(2).Value)
                    intFlag = 2
                End If
                blnDiscountflag = True
            End If

            If intFlag <> 2 Then 'Take to price for BP Price List
                strSQL = "SELECT T1.[ItemCode], T1.[PriceList], T1.[Price] FROM OPLN T0  INNER JOIN ITM1 T1 ON T0.ListNum = T1.PriceList where T1.Itemcode='" & StrItem & "' and T1.PriceList=" & oBP.PriceListNum
                oRec.DoQuery(strSQL)
                If oRec.RecordCount > 0 Then
                    price = Convert.ToDouble(oRec.Fields.Item(2).Value)
                End If
            End If
        End If
        ' amatrix.Columns.Item("14").Cells.Item(intRow).Specific.value = price
        amatrix.Columns.Item("15").Cells.Item(intRow).Specific.value = discount
        oBP = Nothing
        oItem = Nothing

    End Sub
#End Region

#Region "Get Item Price with Factor"
    Public Function getPrcWithFactor(ByVal CardCode As String, ByVal ItemCode As String, ByVal RntlDays As Integer, ByVal Qty As Double) As Double
        Dim oItem As SAPbobsCOM.Items
        Dim Price, Expressn As Double
        Dim oDataSet, oRecSet As SAPbobsCOM.Recordset

        oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oApplication.Utilities.ExecuteSQL(oDataSet, "Select U_RentFac, U_NumDys From [@REN_FACT] order by U_NumDys ")
        If oItem.GetByKey(ItemCode) And oDataSet.RecordCount > 0 Then

            oApplication.Utilities.ExecuteSQL(oRecSet, "Select ListNum from OCRD where CardCode = '" & CardCode & "'")
            oItem.PriceList.SetCurrentLine(oRecSet.Fields.Item(0).Value - 1)
            Price = oItem.PriceList.Price
            Expressn = 0
            oDataSet.MoveFirst()

            While RntlDays > 0

                If oDataSet.EoF Then
                    oDataSet.MoveLast()
                End If

                If RntlDays < oDataSet.Fields.Item(1).Value Then
                    Expressn += (oDataSet.Fields.Item(0).Value * RntlDays * Price * Qty)
                    RntlDays = 0
                    Exit While
                End If
                Expressn += (oDataSet.Fields.Item(0).Value * oDataSet.Fields.Item(1).Value * Price * Qty)
                RntlDays -= oDataSet.Fields.Item(1).Value
                oDataSet.MoveNext()

            End While

        End If
        If oItem.UserFields.Fields.Item("U_Rental").Value = "Y" Then
            Return CDbl(Expressn / Qty)
        Else
            Return Price
        End If


    End Function
#End Region

#Region "Get WareHouse List"
    Public Function getUsedWareHousesList(ByVal ItemCode As String, ByVal Quantity As Double) As DataTable
        Dim oDataTable As DataTable
        Dim oRow As DataRow
        Dim rswhs As SAPbobsCOM.Recordset
        Dim LeftQty As Double
        Try
            oDataTable = New DataTable
            oDataTable.Columns.Add(New System.Data.DataColumn("ItemCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("WhsCode"))
            oDataTable.Columns.Add(New System.Data.DataColumn("Quantity"))

            strSQL = "Select WhsCode, ItemCode, (OnHand + OnOrder - IsCommited) As Available From OITW Where ItemCode = '" & ItemCode & "' And " & _
                        "WhsCode Not In (Select Whscode From OWHS Where U_Reserved = 'Y' Or U_Rental = 'Y') Order By (OnHand + OnOrder - IsCommited) Desc "

            ExecuteSQL(rswhs, strSQL)
            LeftQty = Quantity

            While Not rswhs.EoF
                oRow = oDataTable.NewRow()

                oRow.Item("WhsCode") = rswhs.Fields.Item("WhsCode").Value
                oRow.Item("ItemCode") = rswhs.Fields.Item("ItemCode").Value

                LeftQty = LeftQty - CType(rswhs.Fields.Item("Available").Value, Double)

                If LeftQty <= 0 Then
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double) + LeftQty
                    oDataTable.Rows.Add(oRow)
                    Exit While
                Else
                    oRow.Item("Quantity") = CType(rswhs.Fields.Item("Available").Value, Double)
                End If

                oDataTable.Rows.Add(oRow)
                rswhs.MoveNext()
                oRow = Nothing
            End While

            'strSQL = ""
            'For count As Integer = 0 To oDataTable.Rows.Count - 1
            '    strSQL += oDataTable.Rows(count).Item("WhsCode") & " : " & oDataTable.Rows(count).Item("Quantity") & vbNewLine
            'Next
            'MessageBox.Show(strSQL)

            Return oDataTable

        Catch ex As Exception
            Throw ex
        Finally
            oRow = Nothing
        End Try
    End Function
#End Region

#Region "GetDocumentQuantity"
    Public Function getDocumentQuantity(ByVal strQuantity As String) As Double
        Dim dblQuant As Double
        Dim strTemp, strTemp1 As String
        strTemp1 = strQuantity
        strTemp = CompanyDecimalSeprator
        If strTemp1 = "" Then
            Return 0
        End If
        If CompanyDecimalSeprator <> "." Then
            If CompanyThousandSeprator <> strTemp Then
            End If
            strQuantity = strQuantity.Replace(".", CompanyDecimalSeprator)
        End If
        Try
            dblQuant = Convert.ToDouble(strQuantity)
        Catch ex As Exception
            dblQuant = Convert.ToDouble(strTemp1)
        End Try
        Return dblQuant
    End Function
#End Region


#Region "Set / Get Values from Matrix"
    Public Function getMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer) As String
        Return aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value
    End Function
    Public Sub SetMatrixValues(ByVal aMatrix As SAPbouiCOM.Matrix, ByVal coluid As String, ByVal intRow As Integer, ByVal strvalue As String)
        aMatrix.Columns.Item(coluid).Cells.Item(intRow).Specific.value = strvalue
    End Sub
#End Region

#Region "Get Edit Text"
    Public Function getEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String) As String
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Return objEdit.String
    End Function
    Public Sub setEdittextvalue(ByVal aform As SAPbouiCOM.Form, ByVal UID As String, ByVal newvalue As String)
        Dim objEdit As SAPbouiCOM.EditText
        objEdit = aform.Items.Item(UID).Specific
        Try
            objEdit.String = newvalue
        Catch ex As Exception
            objEdit.Value = newvalue
        End Try

    End Sub
#End Region

#End Region

#Region "Write to LogFile"
    Public Sub WriteToLogFile(ByVal strMsg As String)
        Dim dtdate As Date
        Dim strFileName As String
        Dim FS As FileStream
        Try
            ErrorLogFile = System.Windows.Forms.Application.StartupPath & "\Log.txt"
            strFileName = ErrorLogFile
            If File.Exists(strFileName) Then
                FS = New FileStream(strFileName, FileMode.Append)
            Else
                FS = New FileStream(strFileName, FileMode.Create, FileAccess.ReadWrite)
            End If
            Dim SW As New StreamWriter(FS)
            strMsg = strMsg
            SW.WriteLine(strMsg)
            SW.Close()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "Functions related to Load XML"

#Region "Add/Remove Menus "
    Public Sub AddRemoveMenus(ByVal sFileName As String)
        Dim oXMLDoc As New Xml.XmlDocument
        Dim sFilePath As String
        Try
            sFilePath = getApplicationPath() & "\XML Files\" & sFileName
            oXMLDoc.Load(sFilePath)
            oApplication.SBO_Application.LoadBatchActions(oXMLDoc.InnerXml)
        Catch ex As Exception
            Throw ex
        Finally
            oXMLDoc = Nothing
        End Try
    End Sub
#End Region

#Region "Load XML File "
    Private Function LoadXMLFiles(ByVal sFileName As String) As String
        Dim oXmlDoc As Xml.XmlDocument
        Dim oXNode As Xml.XmlNode
        Dim oAttr As Xml.XmlAttribute
        Dim sPath As String
        Dim FrmUID As String
        Try
            oXmlDoc = New Xml.XmlDocument

            sPath = getApplicationPath() & "\XML Files\" & sFileName

            oXmlDoc.Load(sPath)
            oXNode = oXmlDoc.GetElementsByTagName("form").Item(0)
            oAttr = oXNode.Attributes.GetNamedItem("uid")
            oAttr.Value = oAttr.Value & FormNum
            FormNum = FormNum + 1
            oApplication.SBO_Application.LoadBatchActions(oXmlDoc.InnerXml)
            FrmUID = oAttr.Value

            Return FrmUID

        Catch ex As Exception
            Throw ex
        Finally
            oXmlDoc = Nothing
        End Try
    End Function
#End Region

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String) As SAPbouiCOM.Form
        'Return LoadForm(XMLFile, FormType.ToString(), FormType & "_" & oApplication.SBO_Application.Forms.Count.ToString)
        LoadXMLFiles(XMLFile)
        Return Nothing
    End Function

    '*****************************************************************
    'Type               : Function   
    'Name               : LoadForm
    'Parameter          : XmlFile,FormType,FormUID
    'Return Value       : SBO Form
    'Author             : Senthil Kumar B Senthil Kumar B
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To Load XML file 
    '*****************************************************************

    Public Function LoadForm(ByVal XMLFile As String, ByVal FormType As String, ByVal FormUID As String) As SAPbouiCOM.Form

        Dim oXML As System.Xml.XmlDocument
        Dim objFormCreationParams As SAPbouiCOM.FormCreationParams
        Try
            oXML = New System.Xml.XmlDocument
            oXML.Load(XMLFile)
            objFormCreationParams = (oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams))
            objFormCreationParams.XmlData = oXML.InnerXml
            objFormCreationParams.FormType = FormType
            objFormCreationParams.UniqueID = FormUID
            Return oApplication.SBO_Application.Forms.AddEx(objFormCreationParams)
        Catch ex As Exception
            Throw ex

        End Try

    End Function



#Region "Load Forms"
    Public Sub LoadForm(ByRef oObject As Object, ByVal XmlFile As String)
        Try
            oObject.FrmUID = LoadXMLFiles(XmlFile)
            oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FrmUID)
            If Not oApplication.Collection.ContainsKey(oObject.FrmUID) Then
                oApplication.Collection.Add(oObject.FrmUID, oObject)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#End Region

#Region "Functions related to System Initilization"

#Region "Create Tables"
    Public Sub CreateTables()
        Dim oCreateTable As clsTable
        Try
            oCreateTable = New clsTable
            oCreateTable.CreateTables()
        Catch ex As Exception
            Throw ex
        Finally
            oCreateTable = Nothing
        End Try
    End Sub
#End Region

#Region "Notify Alert"
    Public Sub NotifyAlert()
        'Dim oAlert As clsPromptAlert

        'Try
        '    oAlert = New clsPromptAlert
        '    oAlert.AlertforEndingOrdr()
        'Catch ex As Exception
        '    Throw ex
        'Finally
        '    oAlert = Nothing
        'End Try

    End Sub
#End Region

#End Region

#Region "Function related to Quantities"

#Region "Get Available Quantity"
    Public Function getAvailableQty(ByVal ItemCode As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset

        strSQL = "Select SUM(T1.OnHand + T1.OnOrder - T1.IsCommited) From OITW T1 Left Outer Join OWHS T3 On T3.Whscode = T1.WhsCode " & _
                    "Where T1.ItemCode = '" & ItemCode & "'"
        Me.ExecuteSQL(rsQuantity, strSQL)

        If rsQuantity.Fields.Item(0) Is System.DBNull.Value Then
            Return 0
        Else
            Return CLng(rsQuantity.Fields.Item(0).Value)
        End If

    End Function
#End Region

#Region "Get Rented Quantity"
    Public Function getRentedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim RentedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_RDR1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_ORDR] Where U_Status = 'R') " & _
                    " and '" & StartDate & "' between [@REN_RDR1].U_ShipDt1 and [@REN_RDR1].U_ShipDt2 "
        '" and [@REN_RDR1].U_ShipDt1 between '" & StartDate & "' and '" & EndDate & "'"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            RentedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return RentedQty

    End Function
#End Region

#Region "Get Reserved Quantity"
    Public Function getReservedQty(ByVal ItemCode As String, ByVal StartDate As String, ByVal EndDate As String) As Long
        Dim rsQuantity As SAPbobsCOM.Recordset
        Dim ReservedQty As Long

        strSQL = " select Sum(U_ReqdQty) from [@REN_QUT1] Where U_ItemCode = '" & ItemCode & "' " & _
                    " And DocEntry IN " & _
                    " (Select DocEntry from [@REN_OQUT] Where U_Status = 'R' And Status = 'O') " & _
                    " and '" & StartDate & "' between [@REN_QUT1].U_ShipDt1 and [@REN_QUT1].U_ShipDt2"

        ExecuteSQL(rsQuantity, strSQL)
        If Not rsQuantity.Fields.Item(0).Value Is System.DBNull.Value Then
            ReservedQty = rsQuantity.Fields.Item(0).Value
        End If

        Return ReservedQty

    End Function
#End Region
    Public Function getAccountCode(ByVal aCode As String) As String
        Dim oRS As SAPbobsCOM.Recordset
        oRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRS.DoQuery("Select AcctCode from OACT where FormatCode='" & aCode & "'")
        If oRS.RecordCount > 0 Then
            '    MsgBox(oRS.Fields.Item(0).Value)
            Return oRS.Fields.Item(0).Value
        Else
            Return ""
        End If
    End Function

#End Region

#Region "Functions related to Tax"

#Region "Get Tax Codes"
    Public Sub getTaxCodes(ByRef oCombo As SAPbouiCOM.ComboBox)
        Dim rsTaxCodes As SAPbobsCOM.Recordset

        strSQL = "Select Code, Name From OVTG Where Category = 'O' Order By Name"
        Me.ExecuteSQL(rsTaxCodes, strSQL)

        oCombo.ValidValues.Add("", "")
        If rsTaxCodes.RecordCount > 0 Then
            While Not rsTaxCodes.EoF
                oCombo.ValidValues.Add(rsTaxCodes.Fields.Item(0).Value, rsTaxCodes.Fields.Item(1).Value)
                rsTaxCodes.MoveNext()
            End While
        End If
        oCombo.ValidValues.Add("Define New", "Define New")
        'oCombo.Select("")
    End Sub
#End Region

#Region "Get Applicable Code"

    Public Function getApplicableTaxCode1(ByVal CardCode As String, ByVal ItemCode As String, ByVal Shipto As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    strSQL = "select LicTradNum from CRD1 where Address ='" & Shipto & "' and CardCode ='" & CardCode & "'"
                    Me.ExecuteSQL(rsExempt, strSQL)
                    If rsExempt.RecordCount > 0 Then
                        rsExempt.MoveFirst()
                        TaxGroup = rsExempt.Fields.Item(0).Value
                    Else
                        TaxGroup = ""
                    End If
                    'TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If




        Return TaxGroup

    End Function


    Public Function getApplicableTaxCode(ByVal CardCode As String, ByVal ItemCode As String) As String
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim oItem As SAPbobsCOM.Items
        Dim rsExempt As SAPbobsCOM.Recordset
        Dim TaxGroup As String
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        If oBP.GetByKey(CardCode.Trim) Then
            If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
                If oBP.VatGroup.Trim <> "" Then
                    TaxGroup = oBP.VatGroup.Trim
                Else
                    TaxGroup = oBP.FederalTaxID
                End If
            ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
                strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
                Me.ExecuteSQL(rsExempt, strSQL)
                If rsExempt.RecordCount > 0 Then
                    rsExempt.MoveFirst()
                    TaxGroup = rsExempt.Fields.Item(0).Value
                Else
                    TaxGroup = ""
                End If
            End If
        End If

        'If oBP.GetByKey(CardCode.Trim) Then
        '    If oBP.VatLiable = SAPbobsCOM.BoVatStatus.vLiable Or oBP.VatLiable = SAPbobsCOM.BoVatStatus.vEC Then
        '        If oBP.VatGroup.Trim <> "" Then
        '            TaxGroup = oBP.VatGroup.Trim
        '        Else
        '            oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '            If oItem.GetByKey(ItemCode.Trim) Then
        '                TaxGroup = oItem.SalesVATGroup.Trim
        '            End If
        '        End If
        '    ElseIf oBP.VatLiable = SAPbobsCOM.BoVatStatus.vExempted Then
        '        strSQL = "Select Code From OVTG Where Rate = 0 And Category = 'O' Order By Code"
        '        Me.ExecuteSQL(rsExempt, strSQL)
        '        If rsExempt.RecordCount > 0 Then
        '            rsExempt.MoveFirst()
        '            TaxGroup = rsExempt.Fields.Item(0).Value
        '        Else
        '            TaxGroup = ""
        '        End If
        '    End If
        'End If
        Return TaxGroup

    End Function
#End Region

#End Region

#Region "Log Transaction"
    Public Sub LogTransaction(ByVal DocNum As Integer, ByVal ItemCode As String, _
                                    ByVal FromWhs As String, ByVal TransferedQty As Double, ByVal ProcessDate As Date)
        Dim sCode As String
        Dim sColumns As String
        Dim sValues As String
        Dim rsInsert As SAPbobsCOM.Recordset

        sCode = Me.getMaxCode("@REN_PORDR", "Code")

        sColumns = "Code, Name, U_DocNum, U_WhsCode, U_ItemCode, U_Quantity, U_RetQty, U_Date"
        sValues = "'" & sCode & "','" & sCode & "'," & DocNum & ",'" & FromWhs & "','" & ItemCode & "'," & TransferedQty & ", 0, Convert(DateTime,'" & ProcessDate.ToString("yyyyMMdd") & "')"

        strSQL = "Insert into [@REN_PORDR] (" & sColumns & ") Values (" & sValues & ")"
        oApplication.Utilities.ExecuteSQL(rsInsert, strSQL)

    End Sub

    Public Sub LogCreatedDocument(ByVal DocNum As Integer, ByVal CreatedDocType As SAPbouiCOM.BoLinkedObject, ByVal CreatedDocNum As String, ByVal sCreatedDate As String)
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim sCode As String
        Dim CreatedDate As DateTime
        Try
            oUserTable = oApplication.Company.UserTables.Item("REN_DORDR")

            sCode = Me.getMaxCode("@REN_DORDR", "Code")

            If Not oUserTable.GetByKey(sCode) Then
                oUserTable.Code = sCode
                oUserTable.Name = sCode

                With oUserTable.UserFields.Fields
                    .Item("U_DocNum").Value = DocNum
                    .Item("U_DocType").Value = CInt(CreatedDocType)
                    .Item("U_DocEntry").Value = CInt(CreatedDocNum)

                    If sCreatedDate <> "" Then
                        CreatedDate = CDate(sCreatedDate.Insert(4, "/").Insert(7, "/"))
                        .Item("U_Date").Value = CreatedDate
                    Else
                        .Item("U_Date").Value = CDate(Format(Now, "Long Date"))
                    End If

                End With

                If oUserTable.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserTable = Nothing
        End Try
    End Sub
#End Region

    Public Function FormatDataSourceValue(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If

            If Value.IndexOf(CompanyThousandSeprator) > -1 Then
                Value = Value.Replace(CompanyThousandSeprator, "")
            End If
        Else
            Value = "0"

        End If

        ' NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue


        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue
    End Function

    Public Function FormatScreenValues(ByVal Value As String) As Double
        Dim NewValue As Double

        If Value <> "" Then
            If Value.IndexOf(".") > -1 Then
                Value = Value.Replace(".", CompanyDecimalSeprator)
            End If
        Else
            Value = "0"
        End If

        'NewValue = CDbl(Value)
        NewValue = Val(Value)

        Return NewValue

        'Dim dblValue As Double
        'Value = Value.Replace(CompanyThousandSeprator, "")
        'Value = Value.Replace(CompanyDecimalSeprator, System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
        'dblValue = Val(Value)
        'Return dblValue

    End Function

    Public Function SetScreenValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function

    Public Function SetDBValues(ByVal Value As String) As String

        If Value.IndexOf(CompanyDecimalSeprator) > -1 Then
            Value = Value.Replace(CompanyDecimalSeprator, ".")
        End If

        Return Value

    End Function



#Region "Load Check Details"
    Public Sub LoadCheckDetails(ByVal aForm As SAPbouiCOM.Form)
        Dim oGrid As SAPbouiCOM.Grid
        Dim strCheck, strContrFrom, strContrTo As String
        strContrFrom = oApplication.Utilities.getEdittextvalue(aForm, "8")
        strContrTo = oApplication.Utilities.getEdittextvalue(aForm, "9")
        strCheck = ""
        If strContrFrom <> "" Then
            strCheck = strCheck & "   T0.[U_Z_CONTID] >=" & strContrFrom
        Else
            strCheck = strCheck & "  1=1"
        End If

        If strContrTo <> "" Then
            strCheck = strCheck & " and  T0.[U_Z_CONTID] <=" & strContrTo
        Else
            strCheck = strCheck & " and 1=1"
        End If
        oGrid = aForm.Items.Item("7").Specific
        '   Dim st As String = "SELECT T0.[U_Z_CONTID] 'Contract ID',  T0.[U_Z_CNTNUMBER] 'Contract Number', T0.[DocNum] 'Incoming Payment Number', T0.[DueDate], T0.[CheckSum], T0.[Currency], T0.[U_Z_FROMDATE] 'From Date', T0.[U_Z_TODATE] 'End Date' FROM RCT1 T0 inner Join ORCT T1 on T1.DocEntry=T0.DocNum "
        Dim st As String = "SELECT T0.[U_Z_CONTID] 'Contract ID',  T0.[U_Z_CNTNUMBER] 'Contract Number', T0.[DocNum] 'Incoming Payment Number', T0.[DueDate], T0.[CheckNum], T0.[BankCode], T0.[Branch], T0.[AcctNum],T0.[CheckSum], T0.[Currency], T0.[U_Z_FROMDATE] 'From Date', T0.[U_Z_TODATE] 'End Date' FROM RCT1 T0 inner Join ORCT T1 on T1.DocEntry=T0.DocNum"
        st = st & " where " & strCheck
        oGrid.DataTable.ExecuteQuery(st)
        Dim oEdittext As SAPbouiCOM.EditTextColumn
        oEdittext = oGrid.Columns.Item("Incoming Payment Number")
        oEdittext.LinkedObjectType = "24"
        oGrid.AutoResizeColumns()
    End Sub
#End Region

#Region "Approval Functionalities"
    Public Sub Resize(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            aForm.Items.Item("1").Height = (aForm.Height / 2) - 50
            aForm.Items.Item("1").Width = aForm.Width - 10
            aForm.Items.Item("4").Top = aForm.Items.Item("1").Top + aForm.Items.Item("1").Height + 1
            aForm.Items.Item("5").Top = aForm.Items.Item("4").Top
            aForm.Items.Item("3").Top = aForm.Items.Item("4").Top + aForm.Items.Item("4").Height + 5
            aForm.Items.Item("3").Width = (aForm.Width / 2)
            aForm.Items.Item("3").Height = (aForm.Height / 2) - 50
            aForm.Items.Item("5").Left = aForm.Items.Item("3").Left + aForm.Items.Item("3").Width + 50
            aForm.Items.Item("7").Left = aForm.Items.Item("5").Left
            aForm.Items.Item("9").Left = aForm.Items.Item("5").Left
            aForm.Items.Item("8").Left = aForm.Items.Item("7").Left + aForm.Items.Item("7").Width + 1
            aForm.Items.Item("10").Left = aForm.Items.Item("9").Left + aForm.Items.Item("9").Width + 1
            aForm.Items.Item("8").Top = aForm.Items.Item("3").Top
            aForm.Items.Item("7").Top = aForm.Items.Item("8").Top
            aForm.Items.Item("10").Top = aForm.Items.Item("8").Top + aForm.Items.Item("8").Height + 1
            aForm.Items.Item("9").Top = aForm.Items.Item("10").Top
            aForm.Freeze(False)
        Catch ex As Exception

        End Try
    End Sub
    Public Function ApprovalValidation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Try
            oCombo = aform.Items.Item("8").Specific
            oExEdit = aform.Items.Item("10").Specific
            If oCombo.Selected.Value = "R" Then
                If oExEdit.Value = "" Then
                    oApplication.Utilities.Message("Remarks is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
    End Function
    Public Sub LoadStatusRemarks(ByVal aForm As SAPbouiCOM.Form, ByVal intRow As Integer)
        Try
            aForm.Freeze(True)
            oGrid = aForm.Items.Item("3").Specific
            oEdit = aForm.Items.Item("6").Specific
            oCombo = aForm.Items.Item("8").Specific
            oExEdit = aForm.Items.Item("10").Specific
            oEdit.Value = oGrid.DataTable.GetValue("DocEntry", intRow)
            oCombo.Select("P", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oExEdit.Value = oGrid.DataTable.GetValue("U_Z_Remarks", intRow)

            If oGrid.DataTable.GetValue("U_Z_ApproveBy", intRow) <> oApplication.Company.UserName Then
                aForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                aForm.Items.Item("8").Enabled = False
                aForm.Items.Item("10").Enabled = False
            Else
                aForm.Items.Item("8").Enabled = True
                aForm.Items.Item("10").Enabled = True
            End If
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Function GetTemplateID(ByVal aForm As SAPbouiCOM.Form, ByVal DocType As String) As String
        Try
            Dim strQuery As String = ""
            Dim Status As String = ""
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select * from ""@Z_OAPPT"" T0 left join ""@Z_APPT2"" T1 on T0.""DocEntry""=T1.""DocEntry"" where isnull(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType & "' "
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                Status = oRecordSet.Fields.Item("DocEntry").Value
            Else
                Status = "0"
            End If
            Return Status
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End Try
    End Function
    Public Sub UpdateApprovalRequired(ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal ReqValue As String, ByVal AppTempId As String)
        Try
            Dim strQuery As String
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Update [" & strTable & "] set U_Z_AppRequired='" & ReqValue & "',U_Z_AppReqDate=getdate()  where " & sColumn & "='" & StrCode & "'"
            oRecordSet.DoQuery(strQuery)
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
        End Try
    End Sub


    Public Sub UpdateApprovalRequired1(ByVal strTable As String, ByVal sColumn As String, ByVal StrCode As String, ByVal ReqValue As String, ByVal AppTempId As String)
        Try
            Dim strQuery As String
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Update [" & strTable & "] set U_Z_AppRequired1='" & ReqValue & "',U_Z_AppReqDate1=getdate()  where " & sColumn & "='" & StrCode & "'"
            oRecordSet.DoQuery(strQuery)
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
        End Try
    End Sub
    Public Sub assignMatrixLineno(ByVal aGrid As SAPbouiCOM.Grid, ByVal aform As SAPbouiCOM.Form)
        aform.Freeze(True)
        Try
            For intNo As Integer = 0 To aGrid.DataTable.Rows.Count - 1
                aGrid.RowHeaders.SetText(intNo, intNo + 1)
            Next
        Catch ex As Exception
        End Try
        aGrid.Columns.Item("RowsHeader").TitleObject.Caption = "#"
        aform.Freeze(False)
    End Sub
    Public Sub InitializationApproval(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String)
        Try
            Dim oTempDt As SAPbouiCOM.DataTable
            Select Case enDocType
                Case "TEA"
                    sQuery = " select T0.DocEntry,U_Z_UNITCODE,U_Z_DESC,U_Z_STARTDATE,U_Z_ENDDATE,U_Z_TENCODE,U_Z_TENNAME,U_Z_ANNUALRENT,U_Z_DEPOSIT,U_Z_MONTHLY,U_Z_APPSTATUS,U_Z_TERSTATUS,U_Z_CURAPPROVER,U_Z_NXTAPPROVER, "
                    sQuery += " Case U_Z_AppRequired when 'Y' then 'Yes' else 'No' End as  'Approval Required',U_Z_AppReqDate 'Requested Date'"
                    sQuery += " From [@Z_CONTRACT]  T0 Left JOIN [@Z_APPT2] T2 ON (T0.U_Z_CURAPPROVER  = T2.U_Z_AUSER or T0.U_Z_NXTAPPROVER=T2.U_Z_AUSER) and T0.""U_Z_AppStatus""='N'  "
                    sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                    sQuery += " And (T0.U_Z_CurApprover = '" + oApplication.Company.UserName + "' OR T0.U_Z_NxtApprover = '" + oApplication.Company.UserName + "')"
                    sQuery += " And isnull(T2.U_Z_AMan,'N')='Y' AND isnull(T3.U_Z_Active,'N')='Y' and  isnull(T0.U_Z_AppRequired,'N')='Y' and  T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType + "'"
                    sQuery += "  Order by T0.DocEntry Desc"
                Case "TER"
                    sQuery = " select T0.DocEntry,U_Z_CNTNO,U_Z_UNITCODE,U_Z_DESC,U_Z_STARTDATE,U_Z_ENDDATE,U_Z_TENCODE,U_Z_TENNAME,U_Z_ANNUALRENT,U_Z_DEPOSIT,U_Z_MONTHLY,Case U_Z_STATUS when 'A' then 'Approved' when 'R' then 'Rejected' else 'Open' end as U_Z_STATUS,U_Z_CURAPPROVER,U_Z_NXTAPPROVER,U_Z_DOCNO, "
                    sQuery += " Case U_Z_AppRequired when 'Y' then 'Yes' else 'No' End as  'Approval Required',U_Z_AppReqDate 'Requested Date'"
                    sQuery += " From [@Z_TCONTRACT]  T0 Left JOIN [@Z_APPT2] T2 ON (T0.U_Z_CURAPPROVER  = T2.U_Z_AUSER or T0.U_Z_NXTAPPROVER=T2.U_Z_AUSER )   "
                    sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                    sQuery += " And (T0.U_Z_CurApprover = '" + oApplication.Company.UserName + "' OR T0.U_Z_NxtApprover = '" + oApplication.Company.UserName + "')"
                    sQuery += " And isnull(T2.U_Z_AMan,'N')='Y' AND isnull(T3.U_Z_Active,'N')='Y' and  isnull(T0.U_Z_AppRequired,'N')='Y' and  T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType + "'"
                    sQuery += " And  U_Z_STATUS='O'  Order by T0.DocEntry Desc"
            End Select
            oGrid = aForm.Items.Item("1").Specific
            oTempDt = aForm.DataSources.DataTables.Item("dtDocumentList")
            oTempDt.ExecuteQuery(sQuery)
            oGrid.DataTable.ExecuteQuery(sQuery)
            formatDocument(aForm, enDocType)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub formatDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String)
        Try
            aForm.Freeze(True)
            Dim strQuery As String
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            Dim oRecSet As SAPbobsCOM.Recordset
            Dim oGECol As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("1").Specific
            Select Case enDocType
                Case "TEA"
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Request No."
                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                    oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                    oGrid.Columns.Item("U_Z_UNITCODE").TitleObject.Caption = "Unit Code"
                    oGrid.Columns.Item("U_Z_DESC").TitleObject.Caption = "Unit Description"
                    oGrid.Columns.Item("U_Z_STARTDATE").TitleObject.Caption = "Start Date"
                    oGrid.Columns.Item("U_Z_ENDDATE").TitleObject.Caption = "End Date"
                    oGrid.Columns.Item("U_Z_TENCODE").TitleObject.Caption = "Tenant Code"
                    oGrid.Columns.Item("U_Z_TENNAME").TitleObject.Caption = "Tenant Name"
                    oGrid.Columns.Item("U_Z_ANNUALRENT").TitleObject.Caption = "Annual Rent"
                    oGrid.Columns.Item("U_Z_DEPOSIT").TitleObject.Caption = "Security Deposit"
                    oGrid.Columns.Item("U_Z_MONTHLY").TitleObject.Caption = "Monthly Rent"
                    oGrid.Columns.Item("U_Z_APPSTATUS").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_CURAPPROVER").TitleObject.Caption = "Current Approver"
                    oGrid.Columns.Item("U_Z_NXTAPPROVER").TitleObject.Caption = "Next Approver"
                    oGrid.Columns.Item("U_Z_TERSTATUS").Visible = False
                    oGrid.AutoResizeColumns()
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                Case "TER"
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Request No."
                    oGrid.Columns.Item("U_Z_CNTNO").TitleObject.Caption = "Contract Number"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_CNTNO")
                    oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                    oGrid.Columns.Item("U_Z_UNITCODE").TitleObject.Caption = "Unit Code"
                    oGrid.Columns.Item("U_Z_DESC").TitleObject.Caption = "Unit Description"
                    oGrid.Columns.Item("U_Z_STARTDATE").TitleObject.Caption = "Start Date"
                    oGrid.Columns.Item("U_Z_ENDDATE").TitleObject.Caption = "End Date"
                    oGrid.Columns.Item("U_Z_TENCODE").TitleObject.Caption = "Tenant Code"
                    oGrid.Columns.Item("U_Z_TENNAME").TitleObject.Caption = "Tenant Name"
                    oGrid.Columns.Item("U_Z_ANNUALRENT").TitleObject.Caption = "Annual Rent"
                    oGrid.Columns.Item("U_Z_DEPOSIT").TitleObject.Caption = "Security Deposit"
                    oGrid.Columns.Item("U_Z_MONTHLY").TitleObject.Caption = "Monthly Rent"
                    oGrid.Columns.Item("U_Z_STATUS").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_CURAPPROVER").TitleObject.Caption = "Current Approver"
                    oGrid.Columns.Item("U_Z_NXTAPPROVER").TitleObject.Caption = "Next Approver"
                    oGrid.Columns.Item("U_Z_DOCNO").Visible = False
                    oGrid.AutoResizeColumns()
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub ApprovalSummary(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String)
        Try
            Dim oTempDt As SAPbouiCOM.DataTable
            Select Case enDocType
                Case "TEA"
                    sQuery = " select T0.DocEntry,U_Z_UNITCODE,U_Z_DESC,U_Z_STARTDATE,U_Z_ENDDATE,U_Z_TENCODE,U_Z_TENNAME,U_Z_ANNUALRENT,U_Z_DEPOSIT,U_Z_MONTHLY,U_Z_APPSTATUS,U_Z_TERSTATUS,U_Z_CURAPPROVER,U_Z_NXTAPPROVER, "
                    sQuery += " Case U_Z_AppRequired when 'Y' then 'Yes' else 'No' End as  'Approval Required',U_Z_AppReqDate 'Requested Date'"
                    sQuery += " From [@Z_CONTRACT]  T0 Left JOIN [@Z_APPT2] T2 ON (T0.U_Z_CURAPPROVER  = T2.U_Z_AUSER or T0.U_Z_NXTAPPROVER=T2.U_Z_AUSER ) "
                    sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                    sQuery += " And (T0.U_Z_CurApprover = '" + oApplication.Company.UserName + "' OR T0.U_Z_NxtApprover = '" + oApplication.Company.UserName + "')"
                    sQuery += " And isnull(T2.U_Z_AMan,'N')='Y' AND isnull(T3.U_Z_Active,'N')='Y' and  isnull(T0.U_Z_AppRequired,'N')='Y' and  T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType + "'"
                    sQuery += "  Order by T0.DocEntry Desc"
                Case "TER"
                    sQuery = " select T0.DocEntry,U_Z_CNTNO,U_Z_UNITCODE,U_Z_DESC,U_Z_STARTDATE,U_Z_ENDDATE,U_Z_TENCODE,U_Z_TENNAME,U_Z_ANNUALRENT,U_Z_DEPOSIT,U_Z_MONTHLY,Case U_Z_STATUS when 'A' then 'Approved' when 'R' then 'Rejected' else 'Open' end as U_Z_STATUS,U_Z_CURAPPROVER,U_Z_NXTAPPROVER,U_Z_DOCNO, "
                    sQuery += " Case U_Z_AppRequired when 'Y' then 'Yes' else 'No' End as  'Approval Required',U_Z_AppReqDate 'Requested Date'"
                    sQuery += " From [@Z_TCONTRACT]  T0 Left JOIN [@Z_APPT2] T2 ON (T0.U_Z_CURAPPROVER  = T2.U_Z_AUSER or T0.U_Z_NXTAPPROVER=T2.U_Z_AUSER )   "
                    sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                    sQuery += " And (T0.U_Z_CurApprover = '" + oApplication.Company.UserName + "' OR T0.U_Z_NxtApprover = '" + oApplication.Company.UserName + "')"
                    sQuery += " And isnull(T2.U_Z_AMan,'N')='Y' AND isnull(T3.U_Z_Active,'N')='Y' and  isnull(T0.U_Z_AppRequired,'N')='Y' and  T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType + "'"
                    sQuery += "  Order by T0.DocEntry Desc"
            End Select
            oGrid = aForm.Items.Item("19").Specific
            oTempDt = aForm.DataSources.DataTables.Item("dtDocumentList")
            oTempDt.ExecuteQuery(sQuery)
            oGrid.DataTable.ExecuteQuery(sQuery)
            SummaryDocument(aForm, enDocType)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            oGrid.Columns.Item("RowsHeader").Click(0, False, False)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub SummaryDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String)
        Try
            aForm.Freeze(True)
            Dim strQuery As String
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            Dim oRecSet As SAPbobsCOM.Recordset
            Dim oGECol As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("19").Specific
            Select Case enDocType
                Case "TEA"
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Request No."
                    oEditTextColumn = oGrid.Columns.Item("DocEntry")
                    oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                    oGrid.Columns.Item("U_Z_UNITCODE").TitleObject.Caption = "Unit Code"
                    oGrid.Columns.Item("U_Z_DESC").TitleObject.Caption = "Unit Description"
                    oGrid.Columns.Item("U_Z_STARTDATE").TitleObject.Caption = "Start Date"
                    oGrid.Columns.Item("U_Z_ENDDATE").TitleObject.Caption = "End Date"
                    oGrid.Columns.Item("U_Z_TENCODE").TitleObject.Caption = "Tenant Code"
                    oGrid.Columns.Item("U_Z_TENNAME").TitleObject.Caption = "Tenant Name"
                    oGrid.Columns.Item("U_Z_ANNUALRENT").TitleObject.Caption = "Annual Rent"
                    oGrid.Columns.Item("U_Z_DEPOSIT").TitleObject.Caption = "Security Deposit"
                    oGrid.Columns.Item("U_Z_MONTHLY").TitleObject.Caption = "Monthly Rent"
                    oGrid.Columns.Item("U_Z_APPSTATUS").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_CURAPPROVER").TitleObject.Caption = "Current Approver"
                    oGrid.Columns.Item("U_Z_NXTAPPROVER").TitleObject.Caption = "Next Approver"
                    oGrid.Columns.Item("U_Z_TERSTATUS").Visible = False
                    oGrid.AutoResizeColumns()
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
                Case "TER"
                    oGrid.Columns.Item("DocEntry").TitleObject.Caption = "Request No."
                    oGrid.Columns.Item("U_Z_CNTNO").TitleObject.Caption = "Contract Number"
                    oEditTextColumn = oGrid.Columns.Item("U_Z_CNTNO")
                    oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                    oGrid.Columns.Item("U_Z_UNITCODE").TitleObject.Caption = "Unit Code"
                    oGrid.Columns.Item("U_Z_DESC").TitleObject.Caption = "Unit Description"
                    oGrid.Columns.Item("U_Z_STARTDATE").TitleObject.Caption = "Start Date"
                    oGrid.Columns.Item("U_Z_ENDDATE").TitleObject.Caption = "End Date"
                    oGrid.Columns.Item("U_Z_TENCODE").TitleObject.Caption = "Tenant Code"
                    oGrid.Columns.Item("U_Z_TENNAME").TitleObject.Caption = "Tenant Name"
                    oGrid.Columns.Item("U_Z_ANNUALRENT").TitleObject.Caption = "Annual Rent"
                    oGrid.Columns.Item("U_Z_DEPOSIT").TitleObject.Caption = "Security Deposit"
                    oGrid.Columns.Item("U_Z_MONTHLY").TitleObject.Caption = "Monthly Rent"
                    oGrid.Columns.Item("U_Z_STATUS").TitleObject.Caption = "Approval Status"
                    oGrid.Columns.Item("U_Z_CURAPPROVER").TitleObject.Caption = "Current Approver"
                    oGrid.Columns.Item("U_Z_NXTAPPROVER").TitleObject.Caption = "Next Approver"
                    oGrid.Columns.Item("U_Z_DOCNO").Visible = False
                    oGrid.AutoResizeColumns()
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            End Select
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub LoadHistory(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String, ByVal strDocEntry As String)
        Try
            aForm.Freeze(True)
            Dim oTempDt As SAPbouiCOM.DataTable
            oGrid = aForm.Items.Item("3").Specific
            sQuery = " Select DocEntry,U_Z_DocEntry,U_Z_DocType,U_Z_EmpId,U_Z_EmpName,U_Z_ApproveBy,CreateDate ,CreateTime,UpdateDate,UpdateTime,U_Z_AppStatus,U_Z_Remarks From [@Z_APHIS] "
            sQuery += " Where U_Z_DocType = '" + enDocType.ToString() + "'"
            sQuery += " And U_Z_DocEntry = '" + strDocEntry + "'"
            oTempDt = aForm.DataSources.DataTables.Item("dtHistoryList")
            oTempDt.ExecuteQuery(sQuery)
            oGrid.DataTable = oTempDt
            formatHistory(aForm)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub formatHistory(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oComboBox, oComboBox1, oComboBox2 As SAPbouiCOM.ComboBox
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("3").Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocEntry").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DocType").Visible = False
            oGrid.Columns.Item("U_Z_EmpId").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EmpId")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EmpName").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("U_Z_ApproveBy").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_AppStatus").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_AppStatus").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_AppStatus")
            oGridCombo.ValidValues.Add("P", "Pending")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_Remarks").TitleObject.Caption = "Remarks"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

    Public Sub InitialMessage(ByVal strReqType As String, ByVal strReqNo As String, ByVal strAppStatus As String _
          , ByVal strTemplateNo As String, ByVal strOrginator As String, ByVal enDocType As String)
        Try
            Dim strQuery As String
            Dim strMessageUser As String
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select Top 1 U_Z_AUser From [@Z_APPT2] Where DocEntry = '" + strTemplateNo + "'  and isnull(U_Z_AMan,'')='Y' Order By LineId Asc "
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                strMessageUser = oRecordSet.Fields.Item(0).Value
                oMessage.Subject = strReqType + ":" + "Need Your Approval "
                Dim strMessage As String = ""
                Select Case enDocType
                    Case "TEA"
                        strQuery = "Select * from  [@Z_CONTRACT] where DocEntry='" & strReqNo & "'"
                        oTemp.DoQuery(strQuery)
                        strMessage = " Requested by  :" & oTemp.Fields.Item("U_Z_TENNAME").Value & ": Unit Name : " & oTemp.Fields.Item("U_Z_DESC").Value
                        strOrginator = strMessage
                    Case "TER"
                        strQuery = "Select * from  [@Z_TCONTRACT] where DocEntry='" & strReqNo & "'"
                        oTemp.DoQuery(strQuery)
                        strMessage = " Requested by  :" & oTemp.Fields.Item("U_Z_TENNAME").Value & ": Unit Name : " & oTemp.Fields.Item("U_Z_DESC").Value
                        strOrginator = strMessage
                End Select
                oMessage.Text = strReqType + "  " + strReqNo + " " + strOrginator + " Needs Your Approval "
                oRecipientCollection = oMessage.RecipientCollection

                oRecipientCollection.Add()
                oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                oRecipientCollection.Item(0).UserCode = strMessageUser
                pMessageDataColumns = oMessage.MessageDataColumns

                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "Request No"
                oLines = pMessageDataColumn.MessageDataLines()
                oLine = oLines.Add()
                oLine.Value = strReqNo
                oMessageService.SendMessage(oMessage)
                Dim strEmailMessage As String
                strEmailMessage = strReqType + "  " + strReqNo + " " + strOrginator + " Needs Your Approval "
                ' SendMail_Approval(strEmailMessage, strMessageUser, strMessageUser)
                Select Case enDocType
                    Case "TEA"
                        strQuery = "Update [@Z_CONTRACT] set U_Z_CurApprover='" & strMessageUser & "',U_Z_NxtApprover='" & strMessageUser & "' where DocEntry='" & strReqNo & "'"
                    Case "TER"
                        strQuery = "Update [@Z_TCONTRACT] set U_Z_CurApprover='" & strMessageUser & "',U_Z_NxtApprover='" & strMessageUser & "' where DocEntry='" & strReqNo & "'"
                End Select
                oTemp.DoQuery(strQuery)
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub SendMail_Approval(ByVal aMessage As String, ByVal aMail As String, ByVal aUser As String)
        oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecordSet.DoQuery("Select U_Z_SMTPSERV,U_Z_SMTPPORT,U_Z_SMTPUSER,U_Z_SMTPPWD,U_Z_SSL From [@Z_OMAIL]")
        If Not oRecordSet.EoF Then
            mailServer = oRecordSet.Fields.Item("U_Z_SMTPSERV").Value
            mailPort = oRecordSet.Fields.Item("U_Z_SMTPPORT").Value
            mailId = oRecordSet.Fields.Item("U_Z_SMTPUSER").Value
            mailPwd = oRecordSet.Fields.Item("U_Z_SMTPPWD").Value
            mailSSL = oRecordSet.Fields.Item("U_Z_SSL").Value
            If mailServer <> "" And mailId <> "" And mailPwd <> "" Then
                oRecordSet.DoQuery("Select * from OUSR where USER_CODE='" & aUser & "'")
                aMail = oRecordSet.Fields.Item("E_Mail").Value
                If aMail <> "" Then
                    SendMailforApproval(mailServer, mailPort, mailId, mailPwd, mailSSL, aMail, aMail, "Approval", aMessage)
                End If
            Else
                oApplication.Utilities.Message("Mail Server Details Not Configured...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            End If

        End If
    End Sub
    Private Sub SendMailforApproval(ByVal mailServer As String, ByVal mailPort As String, ByVal mailId As String, ByVal mailpwd As String, ByVal mailSSL As String, ByVal toId As String, ByVal ccId As String, ByVal mType As String, ByVal Message As String)
        Try
            SmtpServer.Credentials = New Net.NetworkCredential(mailId, mailpwd)
            SmtpServer.Port = mailPort
            SmtpServer.EnableSsl = mailSSL
            SmtpServer.Host = mailServer
            mail = New Net.Mail.MailMessage()
            mail.From = New Net.Mail.MailAddress(mailId, "Property Management")
            mail.To.Add(toId)
            mail.IsBodyHtml = True
            mail.Priority = Net.Mail.MailPriority.High
            mail.Subject = Message
            mail.Body = Message
            SmtpServer.Send(mail)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
            mail.Dispose()
        End Try
    End Sub

    Public Function DocApproval(ByVal aForm As SAPbouiCOM.Form, ByVal DocType As String) As String
        Try
            Dim strQuery As String = ""
            Dim Status As String = ""
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select * from ""@Z_OAPPT"" T0 left join ""@Z_APPT2"" T1 on T0.""DocEntry""=T1.""DocEntry"" where isnull(T0.""U_Z_Active"",'N')='Y' and T0.""U_Z_DocType""='" & DocType & "' "
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                Status = "P"
            Else
                Status = "A"
            End If
            Return Status
        Catch ex As Exception
            MsgBox(oApplication.Company.GetLastErrorDescription)
            Return False
        End Try
    End Function
    Public Sub addUpdateDocument(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String)
        Dim oGeneralService As SAPbobsCOM.GeneralService
        Dim oGeneralData As SAPbobsCOM.GeneralData
        Dim oGeneralParams As SAPbobsCOM.GeneralDataParams
        Dim oCompanyService As SAPbobsCOM.CompanyService
        Dim oChildren As SAPbobsCOM.GeneralDataCollection
        oCompanyService = oApplication.Company.GetCompanyService()
        Dim otestRs As SAPbobsCOM.Recordset
        Dim oChild As SAPbobsCOM.GeneralData
        Dim strCode, strQuery As String
        Dim blnRecordExists As Boolean = False
        Dim HeadDocEntry, UserLineId As Integer
        Dim oRecordSet As SAPbobsCOM.Recordset
        Dim oComboBox1, oCombobox2 As SAPbouiCOM.ComboBox
        Try
            If oApplication.SBO_Application.MessageBox("Documents once approved can not be changed. Do you want Continue?", , "Contine", "Cancel") = 2 Then
                Exit Sub
            End If
            oGeneralService = oCompanyService.GetGeneralService("Z_APHIS")
            oGeneralData = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData)
            oGeneralParams = oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otestRs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oGrid = aForm.Items.Item("1").Specific
            oEdit = aForm.Items.Item("6").Specific
            oCombo = aForm.Items.Item("8").Specific
            oExEdit = aForm.Items.Item("10").Specific
            Dim strDocEntry As String = ""
            Dim strDocType1 As String
            Dim strHeader As String = enDocType
            Dim strEmpID As String = ""
            Dim strUnitType As String = ""
            Select Case enDocType
                Case "TEA"
                    strDocType1 = "Contract Request"
                    For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.Rows.IsSelected(index) Then
                            strDocEntry = oGrid.DataTable.GetValue("DocEntry", index)
                            strEmpID = oGrid.DataTable.GetValue("U_Z_TENNAME", index)
                            strUnitType = oGrid.DataTable.GetValue("U_Z_DESC", index)
                            Exit For
                        End If
                    Next
                Case "TER"
                    strDocType1 = "Termination Request"
                    For index As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                        If oGrid.Rows.IsSelected(index) Then
                            strDocEntry = oGrid.DataTable.GetValue("U_Z_DOCNO", index)
                            strEmpID = oGrid.DataTable.GetValue("U_Z_TENNAME", index)
                            strUnitType = oGrid.DataTable.GetValue("U_Z_DESC", index)
                            Exit For
                        End If
                    Next
            End Select
            strQuery = "select T0.DocEntry,T1.LineId from [@Z_OAPPT] T0 JOIN [@Z_APPT2] T1 on T0.DocEntry=T1.DocEntry"
            strQuery += " where T0.U_Z_DocType='" & enDocType.ToString() & "' AND T1.U_Z_AUser='" & oApplication.Company.UserName & "'"
            
            otestRs.DoQuery(strQuery)
            If otestRs.RecordCount > 0 Then
                HeadDocEntry = otestRs.Fields.Item(0).Value
                UserLineId = otestRs.Fields.Item(1).Value
            End If
            Dim strEmpName As String = ""
            strQuery = "Select * from [@Z_APHIS] where U_Z_DocEntry='" & strDocEntry & "' and U_Z_DocType='" & enDocType.ToString() & "' and U_Z_ApproveBy='" & oApplication.Company.UserName & "'"
            oRecordSet.DoQuery(strQuery)
            If oRecordSet.RecordCount > 0 Then
                oGeneralParams.SetProperty("DocEntry", oRecordSet.Fields.Item("DocEntry").Value)
                oGeneralData = oGeneralService.GetByParams(oGeneralParams)
                oGeneralData.SetProperty("U_Z_APPSTATUS", oCombo.Selected.Value)
                oGeneralData.SetProperty("U_Z_REMARKS", oExEdit.Value)
                oGeneralData.SetProperty("U_Z_ADOCENTRY", HeadDocEntry)
                oGeneralData.SetProperty("U_Z_ALINEID", UserLineId)
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp.DoQuery("Select * ,isnull(""firstName"",'') +  ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature)
                If oTemp.RecordCount > 0 Then
                    oGeneralData.SetProperty("U_Z_EMPID", oTemp.Fields.Item("empID").Value.ToString())
                    oGeneralData.SetProperty("U_Z_EMPNAME", oTemp.Fields.Item("EmpName").Value)
                    strEmpName = oTemp.Fields.Item("EmpName").Value
                Else
                    oGeneralData.SetProperty("U_Z_EMPID", "")
                    oGeneralData.SetProperty("U_Z_EMPNAME", "")
                End If
                oGeneralService.Update(oGeneralData)
            ElseIf (strDocEntry <> "" And strDocEntry <> "0") Then
                Dim oTemp As SAPbobsCOM.Recordset
                oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strQuery = "Select * ,isnull(""firstName"",'') + ' ' + isnull(""middleName"",'') +  ' ' + isnull(""lastName"",'') 'EmpName' from OHEM where ""userid""=" & oApplication.Company.UserSignature
                oTemp.DoQuery(strQuery)
                If oTemp.RecordCount > 0 Then
                    oGeneralData.SetProperty("U_Z_EMPID", oTemp.Fields.Item("empID").Value.ToString())
                    oGeneralData.SetProperty("U_Z_EMPNAME", oTemp.Fields.Item("EmpName").Value)
                    strEmpName = oTemp.Fields.Item("EmpName").Value
                Else
                    oGeneralData.SetProperty("U_Z_EMPID", "")
                    oGeneralData.SetProperty("U_Z_EMPNAME", "")
                End If
                oGeneralData.SetProperty("U_Z_DOCENTRY", strDocEntry.ToString())
                oGeneralData.SetProperty("U_Z_DOCTYPE", enDocType.ToString())
                oGeneralData.SetProperty("U_Z_APPSTATUS", oCombo.Selected.Value)
                oGeneralData.SetProperty("U_Z_REMARKS", oExEdit.Value)
                oGeneralData.SetProperty("U_Z_APPROVEBY", oApplication.Company.UserName)
                oGeneralData.SetProperty("U_Z_APPROVEDT", System.DateTime.Now)
                oGeneralData.SetProperty("U_Z_ADOCENTRY", HeadDocEntry)
                oGeneralData.SetProperty("U_Z_ALINEID", UserLineId)
                oGeneralService.Add(oGeneralData)
            End If
            updateFinalStatus(aForm, enDocType, strDocEntry)
            If oCombo.Selected.Value = "A" And oCombo.Selected.Value <> "-" Then
                SendMessage(strDocType1, strDocEntry, oCombo.Selected.Value, HeadDocEntry, strEmpName, oApplication.Company.UserName, enDocType)
            End If
            LoadHistory(aForm, enDocType, strDocEntry)
            InitializationApproval(aForm, enDocType)
            ' ApprovalSummary(aForm, HeadDoc, enDocType)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub updateFinalStatus(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String, ByVal strDocEntry As String)
        Try
            oCombo = aForm.Items.Item("8").Specific
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oComboBox1, oComboBox2 As SAPbouiCOM.ComboBox
            oExEdit = aForm.Items.Item("10").Specific
            If oCombo.Selected.Value = "A" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@Z_APPT2] T2 "
                sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " Where  U_Z_AFinal = 'Y'"
                sQuery += " And T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType.ToString() + "'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    Select Case enDocType
                        Case "TEA"
                            sQuery = "Update ""@Z_CONTRACT"" Set U_Z_AppStatus = 'Y',U_Z_Status='APP',U_Z_ConAppStatus='A' Where DocEntry = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                        Case "TER"
                            sQuery = "Update [@Z_CONTRACT] Set U_Z_TERStatus = 'Y',U_Z_Status='TER',U_Z_TerAppStatus='A' Where DocEntry = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                            sQuery = "Update [@Z_TCONTRACT] Set U_Z_Status='A' Where U_Z_DocNo = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                    End Select
                End If
            ElseIf oCombo.Selected.Value = "R" Then
                sQuery = " Select T2.DocEntry "
                sQuery += " From [@Z_APPT2] T2 "
                sQuery += " JOIN [@Z_OAPPT] T3 ON T2.DocEntry = T3.DocEntry  "
                sQuery += " And T2.U_Z_AUser = '" + oApplication.Company.UserName + "' And T3.U_Z_DocType = '" + enDocType.ToString() + "'"
                oRecordSet.DoQuery(sQuery)
                If Not oRecordSet.EoF Then
                    Select Case enDocType
                        Case "TEA"
                            sQuery = "Update ""@Z_CONTRACT"" Set U_Z_AppStatus = 'Y',U_Z_Status='CAN',U_Z_ConAppStatus='R' Where DocEntry = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                        Case "TER"
                            sQuery = "Update [@Z_TCONTRACT] Set U_Z_Status='R' Where U_Z_DocNo = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                            sQuery = "Update [@Z_CONTRACT] Set U_Z_TERStatus = 'N',U_Z_TerAppStatus='R' Where DocEntry = '" + strDocEntry + "'"
                            oRecordSet.DoQuery(sQuery)
                    End Select
                End If
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub SendMessage(ByVal strReqType As String, ByVal strReqNo As String, ByVal strAppStatus As String _
      , ByVal strTemplateNo As String, ByVal strOrginator As String, ByVal strAuthorizer As String, ByVal enDocType As String)
        Try
            Dim strQuery As String
            Dim strMessageUser As String
            Dim intLineID As Integer
            Dim oRecordSet, oTemp As SAPbobsCOM.Recordset
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim oMessageService As SAPbobsCOM.MessagesService
            Dim oMessage As SAPbobsCOM.Message
            Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns
            Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn
            Dim oLines As SAPbobsCOM.MessageDataLines
            Dim oLine As SAPbobsCOM.MessageDataLine
            Dim oRecipientCollection As SAPbobsCOM.RecipientCollection
            oCmpSrv = oApplication.Company.GetCompanyService()
            oMessageService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService)
            oMessage = oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            strQuery = "Select LineId From [@Z_APPT2] Where DocEntry = '" & strTemplateNo & "' And U_Z_AUser = '" & strAuthorizer & "'"
            oRecordSet.DoQuery(strQuery)
            If Not oRecordSet.EoF Then
                intLineID = CInt(oRecordSet.Fields.Item(0).Value)
                strQuery = "Select Top 1 U_Z_AUser From [@Z_APPT2] Where  DocEntry = '" & strTemplateNo & "' And LineId > '" & intLineID.ToString() & "' and isnull(U_Z_AMan,'')='Y'  Order By LineId Asc "
                oRecordSet.DoQuery(strQuery)

                If Not oRecordSet.EoF Then
                    strMessageUser = oRecordSet.Fields.Item(0).Value
                    oMessage.Subject = strReqType & ":" & " Need Your Approval "
                    Dim strMessage As String = ""
                    strQuery = "Select * from  [@Z_CONTRACT] where DocEntry='" & strReqNo & "'"
                    oTemp.DoQuery(strQuery)
                    strMessage = " Requested by  :" & oTemp.Fields.Item("U_Z_TENNAME").Value & ": Unit Type : " & oTemp.Fields.Item("U_Z_DESC").Value
                    strOrginator = strMessage


                    oMessage.Text = strReqType & " " & strReqNo & strOrginator & " Needs Your Approval "
                    oRecipientCollection = oMessage.RecipientCollection
                    oRecipientCollection.Add()
                    oRecipientCollection.Item(0).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(0).UserCode = strMessageUser
                    pMessageDataColumns = oMessage.MessageDataColumns
                    pMessageDataColumn = pMessageDataColumns.Add()
                    pMessageDataColumn.ColumnName = "Request No"
                    oLines = pMessageDataColumn.MessageDataLines()
                    oLine = oLines.Add()
                    oLine.Value = strReqNo
                    oMessageService.SendMessage(oMessage)
                    Dim strEmailMessage As String = strReqType + "  " + strReqNo + " " + strOrginator + " Needs Your Approval "
                    ' SendMail_Approval(strEmailMessage, strMessageUser, strMessageUser)
                    Select Case enDocType
                        Case "TEA"
                            strQuery = "Update [@Z_CONTRACT] set U_Z_CurApprover='" & oApplication.Company.UserName & "',U_Z_NxtApprover='" & strMessageUser & "' where DocEntry='" & strReqNo & "'"
                        Case "TER"
                            strQuery = "Update [@Z_TCONTRACT] set U_Z_CurApprover='" & oApplication.Company.UserName & "',U_Z_NxtApprover='" & strMessageUser & "' where DocEntry='" & strReqNo & "'"
                    End Select
                    oTemp.DoQuery(strQuery)
                End If
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub SummaryHistory(ByVal aForm As SAPbouiCOM.Form, ByVal enDocType As String, ByVal strDocEntry As String)
        Try
            aForm.Freeze(True)
            Dim oTempDt As SAPbouiCOM.DataTable
            oGrid = aForm.Items.Item("20").Specific
            sQuery = " Select DocEntry,U_Z_DOCENTRY,U_Z_DOCTYPE,U_Z_EMPID,U_Z_EMPNAME,U_Z_APPROVEBY,CreateDate ,CreateTime,UpdateDate,UpdateTime,U_Z_APPSTATUS,U_Z_REMARKS From [@Z_APHIS] "
            sQuery += " Where U_Z_DOCTYPE = '" + enDocType.ToString() + "'"
            sQuery += " And U_Z_DOCENTRY = '" + strDocEntry + "'"
            oTempDt = aForm.DataSources.DataTables.Item("dtHistoryList")
            oTempDt.ExecuteQuery(sQuery)
            oGrid.DataTable = oTempDt
            SummaryformatHistory(aForm)
            oApplication.Utilities.assignMatrixLineno(oGrid, aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
    Private Sub SummaryformatHistory(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oComboBox, oComboBox1, oComboBox2 As SAPbouiCOM.ComboBox
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("20").Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DOCENTRY").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DOCENTRY").Visible = False
            oGrid.Columns.Item("U_Z_DOCTYPE").Visible = False
            oGrid.Columns.Item("U_Z_EMPID").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("U_Z_APPROVEBY").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_APPSTATUS").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_APPSTATUS").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_APPSTATUS")
            oGridCombo.ValidValues.Add("P", "Pending")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub
#End Region
End Class
