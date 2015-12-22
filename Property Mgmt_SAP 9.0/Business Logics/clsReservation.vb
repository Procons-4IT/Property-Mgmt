Public Class clsReservation
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBoxColumn
    Private oCombo As SAPbouiCOM.ComboBox
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False
    Dim count, RowtoDelete, MatrixId As Integer
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Reservation) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Reservation, frm_Reservation)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "4"

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        AddChooseFromList(oForm)
        databind(oForm)
        ' oForm.Items.Item("38").DisplayDesc = True
        oForm.Freeze(False)
    End Sub


#Region "Add Choose From List"
    Private Sub AddChooseFromList(ByVal objForm As SAPbouiCOM.Form)
        Try

            Dim oCFLs As SAPbouiCOM.ChooseFromListCollection
            Dim oCons As SAPbouiCOM.Conditions
            Dim oCon As SAPbouiCOM.Condition


            oCFLs = objForm.ChooseFromLists
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oCFLCreationParams As SAPbouiCOM.ChooseFromListCreationParams
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            ' Adding 2 CFL, one for the button and one for the edit text.
            oCFL = oCFLs.Item("CFL_2")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "S"
            oCFL.SetConditions(oCons)
           

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#End Region

#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oCombo = aForm.Items.Item("33").Specific
            For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try
            Next
            oCombo.ValidValues.Add("", "")
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("SELECT T0.[GroupNum], T0.[PymntGroup] FROM OCTG T0")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub
    Private Sub FillCombo(ByVal aForm As SAPbouiCOM.Form)
        Dim strPropertyCode, strstring As String
        oCombo = aForm.Items.Item("34").Specific
        strPropertyCode = oApplication.Utilities.getEdittextvalue(aForm, "6")
        strstring = "Select U_Z_ProItemCode,U_Z_Desc from [@Z_PROPUNIT] where U_Z_PropCode='" & strPropertyCode & "' order by DocEntry"
        aForm.Freeze(True)
        oApplication.Utilities.FillComboBox(oCombo, strstring)
        aForm.Freeze(False)
    End Sub

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("12").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
        End Select
        Try
            aForm.Freeze(True)

            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If
            oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
            Try
                If oEditText.Value <> "" Then
                    oMatrix.AddRow()
                    Select Case aForm.PaneLevel
                        Case "1"
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                    End Select
                End If

            Catch ex As Exception
                aForm.Freeze(False)
                oMatrix.AddRow()
            End Try

            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)

        End Try
    End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("25").Specific
                'Case "2"
                '    oMatrix = aForm.Items.Item("13").Specific
        End Select
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
            End If
        Next

    End Sub

    Private Function ValidateDeletion(ByVal aForm As SAPbouiCOM.Form) As Boolean
        If intSelectedMatrixrow <= 0 Then
            Return True
        End If
        oMatrix = frmSourceMatrix

        Return True


    End Function
    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry As String

        If validation(aForm) = False Then
            Return False
        End If
        'AssignLineNo(aForm)
        Return True
    End Function
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If Me.MatrixId.ToString = "25" Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
            'Else
            '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        End If
        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
        If intSelectedMatrixrow <= 0 Then
            Exit Sub

        End If
        Me.RowtoDelete = intSelectedMatrixrow
        oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        oMatrix = frmSourceMatrix
        oMatrix.FlushToDataSource()
        For count = 1 To oDataSrc_Line.Size - 1
            oDataSrc_Line.SetValue("LineId", count - 1, count)
        Next
        oMatrix.LoadFromDataSource()
        If oMatrix.RowCount > 0 Then
            oMatrix.DeleteRow(oMatrix.RowCount)
        End If
    End Sub
#End Region
#End Region

#End Region


    Private Function CreateDownPaymentInvoice(ByVal acode As Integer) As Boolean
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strcustCode, stritemcode, strqty, strprice, strtotal, strdriver, strcardname As String
        Dim oDoc As SAPbobsCOM.Documents
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim dblQty, dblPrice As Double
        Dim blnRecordExits As Boolean = False
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        'otemp1.DoQuery("Select isnull(U_Z_DPDocentry,0) from [@Z_CONTRACT] where DocEntry=" & acode)
        'If otemp1.Fields.Item(0).Value > 0 Then
        '    oApplication.Utilities.Message("Delivery already crated for this booking. You can not create downpayment document", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return True
        'End If

        otemp1.DoQuery("SELECT sum(DocTotal)-sum(VatSum)  FROM ODPI T0 WHERE T0.[U_Z_DPType]='A' and T0.[U_Z_ContID] =" & acode)
        Dim dbltotalDownpayment, dbldownpayment As Double
        If otemp1.RecordCount > 0 Then
            dbltotalDownpayment = otemp1.Fields.Item(0).Value
            blnRecordExits = True
        Else
            blnRecordExits = False
            dbltotalDownpayment = 0
        End If
        If dbltotalDownpayment = 0 Then
            blnRecordExits = False
        End If
        Dim strsql, strType, strOwner As String
        otemp.DoQuery("select isnull(U_Z_DownAmount,0), * from [@Z_RESER] where DocEntry=" & acode)
        dbldownpayment = otemp.Fields.Item(0).Value
        If blnRecordExits = True Then
            dbldownpayment = dbldownpayment - dbltotalDownpayment
        End If
        If dbldownpayment > 0 Then
            ' otemp.DoQuery("select isnull(U_Z_Deposit,0) + isnull(U_Z_Salik,0)+isnull(U_Z_DPAmount,0), * from [@Z_ORDR] where DocEntry=" & acode)
            otemp.DoQuery("select isnull(U_Z_DownAmount,0),* from [@Z_RESER] where DocEntry=" & acode)
            If otemp.Fields.Item(0).Value > 0 Then
                strcustCode = otemp.Fields.Item("Type").Value
                If strcustCode = "O" Then
                    strOwner = otemp.Fields.Item("U_Z_TenCode").Value
                Else
                    strOwner = otemp.Fields.Item("U_Z_OwnerCode").Value
                End If
                strcustCode = otemp.Fields.Item("U_Z_TenCode").Value
                stritemcode = otemp.Fields.Item("U_Z_UnitCode").Value
                strprice = otemp.Fields.Item("U_Z_Annualrent").Value
                strtotal = otemp.Fields.Item("U_Z_Annualrent").Value
                dblPrice = dbldownpayment ' otemp.Fields.Item("U_Z_Annualrent").Value
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                oDoc.CardCode = strcustCode
                oDoc.TaxDate = Now.Date
                oDoc.DocDate = Now.Date
                oDoc.DocDueDate = Now.Date
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = acode
                oDoc.UserFields.Fields.Item("U_Z_STARTDATE").Value = otemp.Fields.Item("U_Z_StartDate").Value
                oDoc.UserFields.Fields.Item("U_Z_ENDDATE").Value = otemp.Fields.Item("U_Z_EndDate").Value
                oDoc.UserFields.Fields.Item("U_Z_DPTYPE").Value = "A"
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                Dim oTemp4 As SAPbobsCOM.Recordset
                Dim strProject, strCostCenter As String
                oTemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & otemp.Fields.Item("U_Z_UnitCode").Value & "'")
                oDoc.Project = oTemp4.Fields.Item(0).Value
                strProject = oTemp4.Fields.Item(0).Value
                strCostCenter = oTemp4.Fields.Item(1).Value
                If oBP.GetByKey(strOwner) Then
                    oDoc.Lines.AccountCode = oBP.DownPaymentClearAct
                    'oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(otemp.Fields.Item("U_Z_AcctCode").Value)
                    oDoc.Lines.ItemDescription = "Advance Annual Rent for UntiCode : " & otemp.Fields.Item("U_Z_UnitCode").Value
                    oDoc.Lines.TaxCode = oBP.VatGroup
                    If strCostCenter <> "" Then
                        oDoc.Lines.CostingCode = strCostCenter
                    End If
                    If strProject <> "" Then
                        oDoc.Lines.ProjectCode = strProject
                    End If
                    oDoc.Lines.LineTotal = dblPrice ' otemp.Fields.Item("U_MonthRent").Value
                End If

                If oDoc.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                Else
                    Dim strdocnum As String
                    oApplication.Company.GetNewObjectCode(strdocnum)
                    otemp.DoQuery("Select DocEntry,Docnum from ODPI where DocEntry=" & strdocnum)
                    otemp1.DoQuery("SELECT sum(DocTotal)-sum(VatSum)  FROM ODPI T0 WHERE T0.[U_Z_ContID] =" & acode)
                    If otemp1.RecordCount > 0 Then
                        dbltotalDownpayment = otemp1.Fields.Item(0).Value
                        blnRecordExits = True
                    Else
                        blnRecordExits = False
                        dbltotalDownpayment = 0
                    End If
                    otemp1.DoQuery("Update [@Z_CONTRACT] set U_Z_DPEntry=" & otemp.Fields.Item(0).Value & ",U_Z_DPNumber='" & otemp.Fields.Item(1).Value & "' where docentry=" & acode)
                    oApplication.Utilities.Message("Down Payment Created sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                End If
            End If
        End If
        'CreateIncomingPayment(acode)
    End Function

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strReservedate, strStartDate, strEndDate, strIssueDate, strInvoiceamt As String
        Dim dtreserverdate, dtstartdate, dtenddate, dtissuedate As String
        Dim dblInvoiceamt As Double
        Dim dtStDate, dtEDate As Date
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then


            If oApplication.Utilities.getEdittextvalue(aform, "6") = "" Then
                oApplication.Utilities.Message("Property code can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If oApplication.Utilities.getEdittextvalue(aform, "34") = "" Then
                oApplication.Utilities.Message("Unit Code can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aform, "10") = "" Then
                oApplication.Utilities.Message("Customer Code can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "14") = "" Then
            oApplication.Utilities.Message("Start date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strStartDate = oApplication.Utilities.getEdittextvalue(aform, "14")
            dtstartdate = oApplication.Utilities.GetDateTimeValue(strStartDate)
            dtStDate = oApplication.Utilities.GetDateTimeValue(strStartDate)
        End If

        If oApplication.Utilities.getEdittextvalue(aform, "16") = "" Then
            oApplication.Utilities.Message("End date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strEndDate = oApplication.Utilities.getEdittextvalue(aform, "16")
            dtenddate = oApplication.Utilities.GetDateTimeValue(strEndDate)
            dtEDate = oApplication.Utilities.GetDateTimeValue(strEndDate)
        End If

        If oApplication.Utilities.getEdittextvalue(aform, "24") = "" Then
            oApplication.Utilities.Message("Issue date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strIssueDate = oApplication.Utilities.getEdittextvalue(aform, "24")
            dtissuedate = oApplication.Utilities.GetDateTimeValue(strIssueDate)
        End If

        strReservedate = oApplication.Utilities.getEdittextvalue(aform, "44")
        dtreserverdate = oApplication.Utilities.GetDateTimeValue(strReservedate)

        Dim intDifferenance As Integer
        intDifferenance = dateDifference(dtstartdate, dtreserverdate)

        If intDifferenance < 0 Then
            oApplication.Utilities.Message("Start date should be greater than or equal to Reservation Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        intDifferenance = dateDifference(dtenddate, dtstartdate)

        If intDifferenance < 0 Then
            oApplication.Utilities.Message("End date should be greater than Start Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        intDifferenance = dateDifference(dtissuedate, dtstartdate)

        If intDifferenance < 0 Then

            oApplication.Utilities.Message("Issue date should be greater than Start Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If

        oCombo = aform.Items.Item("42").Specific
        If oCombo.Selected.Value = "Y" Then
            Dim strAmoutn As String
            strAmoutn = oApplication.Utilities.getEdittextvalue(aform, "47")
            If strAmoutn = "" Then
                oApplication.Utilities.Message("Down payment amount should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                If CDbl(strAmoutn) <= 0 Then
                    oApplication.Utilities.Message("Down payment amount should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            'If oApplication.Utilities.getEdittextvalue(aform, "53") = "" Then
            '    oApplication.Utilities.Message("Receiable account is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
        End If
        Dim strSQL1 As String
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            strSQL = "Select * from [@Z_RESER] where ('" & dtStDate.ToString("yyyy-MM-dd") & "' between ""U_Z_StartDate"" and ""U_Z_EndDate"") and isnull(""U_Z_Status"",'P')='P'"
            strSQL1 = "Select * from [@Z_RESER] where ('" & dtStDate.ToString("yyyy-MM-dd") & "' between ""U_Z_StartDate"" and ""U_Z_EndDate"") and isnull(""U_Z_Status"",'P')='C'"
            oRecset.DoQuery(strSQL)
            If oRecset.RecordCount > 0 Then
                If oApplication.SBO_Application.MessageBox("Reserveration already exists for this selected Property / Unit for the period. Do you want to continue ? ", , "Continue", "Cancel") = 2 Then
                    Return False
                Else
                    Return True
                End If
            End If
            oRecset.DoQuery(strSQL1)
            If oRecset.RecordCount > 0 Then
                 oApplication.Utilities.Message("Reserveration already Confirmed for this selected Property / Unit for the period", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        ElseIf aform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
            strSQL = "Select * from ""@Z_RESER"" where DocEntry<>" & oApplication.Utilities.getEdittextvalue(aform, "4") & " and  ('" & dtStDate.ToString("yyyy-MM-dd") & "' between ""U_Z_StartDate"" and ""U_Z_EndDate"") and isnull(""U_Z_Status"",'P')='P'"
            strSQL1 = "Select * from ""@Z_RESER"" where DocEntry<>" & oApplication.Utilities.getEdittextvalue(aform, "4") & " and  ('" & dtStDate.ToString("yyyy-MM-dd") & "' between ""U_Z_StartDate"" and ""U_Z_EndDate"") and isnull(""U_Z_Status"",'P')='C'"
            oRecset.DoQuery(strSQL)
            If oRecset.RecordCount > 0 Then
                If oApplication.SBO_Application.MessageBox("Reserveration already exists for this selected Property / Unit for the period. Do you want to continue ? ", , "Continue", "Cancel") = 2 Then
                    Return False
                Else
                    Return True
                End If
            End If
            oRecset.DoQuery(strSQL1)
            If oRecset.RecordCount > 0 Then
                oApplication.Utilities.Message("Reserveration already Confirmed for this selected Property / Unit for the period", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Return True
    End Function
#End Region

    Private Function dateDifference(ByVal aDate1 As Date, ByVal adate2 As Date) As Double
        Dim intDifferenance As Integer
        intDifferenance = DateDiff(DateInterval.Day, adate2, aDate1)
        Return intDifferenance
    End Function
    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("25").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
            oMatrix.FlushToDataSource()
            For count = 1 To oDataSrc_Line.Size
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix.LoadFromDataSource()

            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub


#Region "Update / Add Item Details in SAP"
    Private Function AddtoItemMaster(ByVal strProjectCode As String, ByVal strChoice As String) As Boolean
        Dim oTempRS, otemp4 As SAPbobsCOM.Recordset
        Dim intDocEntry As Integer
        Dim obp As SAPbobsCOM.BusinessPartners
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        Dim strProject, strProjectName, strItemCode, strItemname, strPropertyCode, strCostCenter As String
        oApplication.Utilities.ExecuteSQL(oTempRS, "Select * from [@Z_RESER] where isnull(U_Z_DownAmount,0) >0 and  U_Z_DownPay='Y' and convert(varchar,isnull(U_Z_DownPayRef,0))='0' and  DocEntry=" & strProjectCode)
        intDocEntry = 0
        strProject = ""
        strProjectName = ""
        strItemCode = ""
        If oTempRS.RecordCount > 0 Then
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strProjectCode = oTempRS.Fields.Item("U_Z_PropCode").Value
            strProject = oTempRS.Fields.Item("U_Z_UnitCode").Value
            strProjectName = oTempRS.Fields.Item("U_Z_AcctCode").Value
            otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp4.DoQuery("Select isnull(U_Z_ACTCODE,'') from [@Z_PROP] where U_Z_COde='" & strProjectCode & "'")
            strProjectName = otemp4.Fields.Item(0).Value
            strProjectName = oApplication.Utilities.getAccountCode(strProjectName)
            strItemCode = strProject
            Dim oDoc As SAPbobsCOM.Documents
            oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
            oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
            otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & strProject & "'")
            strCostCenter = otemp4.Fields.Item(1).Value

            oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

            oDoc.CardCode = oTempRS.Fields.Item("U_Z_CardCode").Value
            oDoc.Project = oTempRS.Fields.Item("U_Z_PropCode").Value
            oDoc.DocDate = oTempRS.Fields.Item("CreateDate").Value
            oDoc.Comments = "Down payment for Property Unit : " & strProject
            oDoc.Lines.SetCurrentLine(0)
            'oDoc.Lines.TaxCode = "0"

            If obp.GetByKey(oTempRS.Fields.Item("U_Z_CardCode").Value) Then
                oDoc.Lines.AccountCode = obp.DownPaymentClearAct
                oDoc.Lines.TaxCode = obp.VatGroup
            End If
            'oDoc.Lines.AccountCode = strProjectName
            If strCostCenter <> "" Then
                oDoc.Lines.CostingCode = strCostCenter
            End If
            oDoc.Lines.ProjectCode = oTempRS.Fields.Item("U_Z_PropCode").Value
            oDoc.Lines.ItemDescription = "Downpayment Amount for Unit Code: " & strProject
            oDoc.Lines.LineTotal = oTempRS.Fields.Item("U_Z_DownAmount").Value
            If oDoc.Add <> 0 Then
                oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                Dim strDocNum As String
                oApplication.Company.GetNewObjectCode(strDocNum)
                oTempRS.DoQuery("Select * from ODPI where DocEntry=" & strDocNum)
                strProject = oTempRS.Fields.Item("DocEntry").Value
                strProjectName = oTempRS.Fields.Item("DocNum").Value
                oTempRS.DoQuery("Update [@Z_RESER] set U_Z_DownPayRef=" & strProject & ",U_Z_DownNumber=" & strProjectName & " where docentry=" & intDocEntry)
            End If
        Else
            Exit Function
        End If
        Return True
    End Function
#End Region
#End Region

#Region "Events"
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Reservation
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "4", oApplication.Utilities.getMaxCode("@Z_RESER", "DocEntry"))
                        oForm.Items.Item("44").Enabled = True
                        oApplication.Utilities.setEdittextvalue(oForm, "44", "t")
                        oApplication.SBO_Application.SendKeys("{TAB}")
                        oForm.Items.Item("44").Enabled = False
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("34").Enabled = True
                        oForm.Items.Item("10").Enabled = True
                    End If

                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oForm.Items.Item("6").Enabled = True
                    oForm.Items.Item("34").Enabled = True
                    oForm.Items.Item("10").Enabled = True


            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region


    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            Try
                If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                    Dim strDocNum, strDocType As String
                    Dim objedittext As SAPbouiCOM.EditText
                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    objedittext = oForm.Items.Item("4").Specific
                    strDocNum = objedittext.String
                    AddtoItemMaster(strDocNum, "Add")
                End If

                If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("34").Enabled = False
                        oForm.Items.Item("10").Enabled = False
                    Else
                        oForm.Items.Item("6").Enabled = True
                        oForm.Items.Item("34").Enabled = True
                        oForm.Items.Item("10").Enabled = True
                    End If
                End If

                If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                    Dim strDocNum, strDocType As String
                    Dim objedittext As SAPbouiCOM.EditText
                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    objedittext = oForm.Items.Item("4").Specific
                    strDocNum = objedittext.String
                    AddtoItemMaster(strDocNum, "Update")
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Reservation Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "47" And pVal.CharPressed <> 9 Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "40") <> "" Then
                                        oApplication.Utilities.Message("Down payment invoice already created", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "42" Then
                                    If oApplication.Utilities.getEdittextvalue(oForm, "40") <> "" Then
                                        oApplication.Utilities.Message("Down payment invoice already created", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "54" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                    Dim strDocNum As String
                                    Dim objedittext As SAPbouiCOM.EditText
                                    objedittext = oForm.Items.Item("4").Specific
                                    strDocNum = objedittext.String
                                    AddtoItemMaster(strDocNum, "Update")
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "34" Then
                                    oCombo = oForm.Items.Item("34").Specific
                                    oApplication.Utilities.setEdittextvalue(oForm, "32", oCombo.Selected.Description)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "34" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "34")
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_PROPUNIT] where U_Z_PropCode='" & oApplication.Utilities.getEdittextvalue(oForm, "6") & "' and U_Z_ProItemCode='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "34", "")
                                        strIns = ""
                                    Else
                                        Exit Sub
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    clsChooseFromList.ItemUID = pVal.ItemUID
                                    clsChooseFromList.SourceFormUID = FormUID
                                    clsChooseFromList.SourceLabel = 0
                                    clsChooseFromList.CFLChoice = "ProUnit" 'oCombo.Selected.Value
                                    clsChooseFromList.choice = "Bin"
                                    'clsChooseFromList.Documentchoice = oApplication.Utilities.getEdittextvalue(oForm, "6") 'TenCode
                                    clsChooseFromList.ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "6") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsChooseFromList.sourceColumID = "32"
                                    oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val, val2 As String
                                Dim intChoice As Integer
                                Dim codebar As String
                                Try
                                    oCFLEvento = pVal
                                    sCHFL_ID = oCFLEvento.ChooseFromListUID
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                    oCFL = oForm.ChooseFromLists.Item(sCHFL_ID)
                                    If (oCFLEvento.BeforeAction = False) Then
                                        Dim oDataTable As SAPbouiCOM.DataTable
                                        oDataTable = oCFLEvento.SelectedObjects
                                        intChoice = 0
                                        'oForm.Freeze(True)
                                        If pVal.ItemUID = "10" Then
                                            oForm.Freeze(True)
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val2 = oDataTable.GetValue("CardName", 0)
                                            val1 = oDataTable.GetValue("GroupNum", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "12", val2)
                                            oCombo = oForm.Items.Item("33").Specific
                                            oCombo.Select(val1, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            oForm.Freeze(False)
                                        End If
                                        If pVal.ItemUID = "20" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception

                                            End Try

                                            oForm.Items.Item("49").Enabled = True
                                            oApplication.Utilities.setEdittextvalue(oForm, "49", val1)
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            oForm.Items.Item("22").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("49").Enabled = False

                                        End If
                                        If pVal.ItemUID = "22" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val1 = oDataTable.GetValue("firstName", 0) + " " + oDataTable.GetValue("lastName", 0)


                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception

                                            End Try
                                            oForm.Items.Item("51").Enabled = True
                                            oApplication.Utilities.setEdittextvalue(oForm, "51", val1)
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            oForm.Items.Item("24").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oForm.Items.Item("51").Enabled = False
                                        End If
                                        If pVal.ItemUID = "6" Then
                                            oForm.Freeze(True)
                                            val = oDataTable.GetValue("U_Z_CODE", 0)
                                            val1 = oDataTable.GetValue("U_Z_DESC", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "30", val1)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                            End Try
                                            oForm.Freeze(False)
                                        ElseIf pVal.ItemUID = "53" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If

                                        'If pVal.ItemUID = "34" Then
                                        '    val = oDataTable.GetValue("U_Z_PROITEMCODE", 0)
                                        '    Try
                                        '        oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        '    Catch ex As Exception
                                        '    End Try
                                        'End If

                                    End If
                                Catch ex As Exception
                                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                    End If

                                    oForm.Freeze(False)
                                End Try
                        End Select
                End Select
            End If

        Catch ex As Exception
            oForm = oApplication.SBO_Application.Forms.Item(FormUID)
            oForm.Freeze(False)
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#End Region
End Class
