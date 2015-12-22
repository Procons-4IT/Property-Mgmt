Imports System.IO
Public Class clsBillGeneration
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox, oCombobox1 As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oRecset As SAPbobsCOM.Recordset
    Private oGrid As SAPbouiCOM.Grid
    Private dtTemp As SAPbouiCOM.DataTable
    Private dtResult As SAPbouiCOM.DataTable
    Private oMode As SAPbouiCOM.BoFormMode
    Private oItem As SAPbobsCOM.Items
    Private oInvoice As SAPbobsCOM.Documents
    Private InvBase As DocumentType
    Private InvBaseDocNo, sPath, sFailureLog As String
    Private InvForConsumedItems, intpane As Integer
    Private blnFlag As Boolean = False
#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_BillGeneration) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_BillGeneration, frm_BillGeneration)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.PaneLevel = 0
        oForm.DataSources.UserDataSources.Add("toProp", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oForm.DataSources.UserDataSources.Add("toUnit", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oEditText = oForm.Items.Item("24").Specific
        oEditText.DataBind.SetBound(True, "", "toProp")
        oEditText.ChooseFromListUID = "CFL_4"
        oEditText.ChooseFromListAlias = "U_Z_CODE"

        oEditText = oForm.Items.Item("26").Specific
        oEditText.DataBind.SetBound(True, "", "toUnit")
        oEditText.ChooseFromListUID = "CFL_5"
        oEditText.ChooseFromListAlias = "U_Z_PROITEMCODE"
        AddChooseFromList(oForm)
        DataBind(oForm)
        '  PropertyBind(oForm)
        ' PropertyUnitBind(oForm)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub

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
            oCFL = oCFLs.Item("PROP")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Canceled"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "N"
            oCFL.SetConditions(oCons)

            'oCFL = oCFLs.Item("UNIT")
            '' Adding Conditions to CFL1
            'oCons = oCFL.GetConditions()
            'oCon = oCons.Add()
            'oCon.Alias = "Canceled"
            'oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            'oCon.CondVal = "N"
            'oCFL.SetConditions(oCons)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub PropertyBind(ByVal aForm As SAPbouiCOM.Form)
        Try
            oCombobox = aForm.Items.Item("20").Specific
            For intRow As Integer = oCombobox.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombobox.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception
                End Try
            Next
            oCombobox.ValidValues.Add("", "")
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_Z_CODE,U_Z_DESC from [@Z_PROP] order by DocEntry")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombobox.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            aForm.Items.Item("20").DisplayDesc = True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub PropertyUnitBind(ByVal aForm As SAPbouiCOM.Form)
        Try
            oCombobox1 = aForm.Items.Item("22").Specific
            For intRow As Integer = oCombobox1.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombobox1.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception
                End Try
            Next
            oCombobox1.ValidValues.Add("", "")
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select U_Z_PROITEMCODE,U_Z_DESC from [@Z_PROPUNIT] order by DocEntry")
            For introw As Integer = 0 To otemp.RecordCount - 1
                Try
                    oCombobox1.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                Catch ex As Exception
                End Try
                otemp.MoveNext()
            Next
            aForm.Items.Item("22").DisplayDesc = True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
    Private Sub DataBind(ByVal aForm As SAPbouiCOM.Form)

        aForm.DataSources.UserDataSources.Add("intYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aForm.DataSources.UserDataSources.Add("intMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        oCombobox = aForm.Items.Item("7").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 2010 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombobox.DataBind.SetBound(True, "", "intYear")
        aForm.Items.Item("7").DisplayDesc = True

        oCombobox = aForm.Items.Item("9").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 1 To 12
            oCombobox.ValidValues.Add(intRow, MonthName(intRow))
        Next

        oCombobox.DataBind.SetBound(True, "", "intMonth")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("9").DisplayDesc = True

        oGrid = aForm.Items.Item("11").Specific
        dtTemp = oGrid.DataTable
        Dim strsql, strsql1 As String

        If intpane = 0 Then
            strsql = "SELECT Year(t0.U_Z_StartDate),Month(t0.U_Z_StartDate),T0.[DocEntry], T0.[U_Z_UNITCODE], T0.[U_Z_TENCODE], T0.[U_Z_ANNUALRENT], T0.[U_Z_PAYTRMS], T0.[U_Z_CHGMONTH], T1.[U_Z_SPACE], T0.[U_Z_STARTDATE], T0.[U_Z_ENDDATE], T0.[U_Z_STATUS] FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
            strsql = strsql & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where (9 between Month(T0.U_Z_Startdate) and month(T0.U_Z_EndDate))"
            strsql = strsql & " and (2012 between Year(T0.U_Z_Startdate) and Year(T0.U_Z_EndDate))"
            strsql1 = "select U_Z_CODE,U_Z_NAME,U_Z_GLACC,U_Z_TYPE,U_Z_RATE from [@Z_OEXP]"
            'dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code  where EmpId=10000000")
        Else
            strsql = "SELECT * from [@Z_OBILL] where U_Month=3333"
            '            dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code ")
        End If
        dtTemp.ExecuteQuery(strsql)
        oGrid.DataTable = dtTemp
        oGrid = aForm.Items.Item("12").Specific
        oGrid.DataTable.ExecuteQuery(strsql1)
        Formatgrid(oForm, "Load")
    End Sub
#End Region


#Region "Generate Billing"

    Private Function GenerateBilling(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim intMonth, intYear As Integer
        Dim strSQL, strRefCode, strCode, strCode1, strSpace As String
        Dim dblSpace, dblAmount As Double
        Dim blnLineExists As Boolean = False
        Dim blnRecordExists As Boolean = False
        Dim oTemp, oTemp1, otemp2 As SAPbobsCOM.Recordset
        Dim oHeaderGrid, oLineGird, oExpGrid As SAPbouiCOM.Grid
        Dim oUserTable, oUsertable1 As SAPbobsCOM.UserTable
        Dim strECode, strESocial, strEname, strETax, strGLAcc, strStartDate, strEndDate, strPostDate As String
        Dim dtPostDate As Date
        oCombobox = aForm.Items.Item("7").Specific
        Try
            aForm.Freeze(True)
            Try
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intYear = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try
            Try
                oCombobox = aForm.Items.Item("9").Specific
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intMonth = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try
            strPostDate = oApplication.Utilities.getEdittextvalue(aForm, "edPost")
            If strPostDate = "" Then
                oApplication.Utilities.Message("Posting date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            Else
                dtPostDate = oApplication.Utilities.GetDateTimeValue(strPostDate)
            End If

            Dim strProperty, strPropertyUnit As String
            If oApplication.Utilities.getEdittextvalue(aForm, "20") = "" Then
                strProperty = " (1=1 "
            Else
                strProperty = "(T1.U_Z_PROPCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "20") & "'"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "24") = "" Then
                strProperty = strProperty & " and 1=1)"
            Else
                strProperty = strProperty & " and T1.U_Z_PROPCODE<='" & oApplication.Utilities.getEdittextvalue(aForm, "24") & "')"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
                strPropertyUnit = "(1=1 "
            Else
                strPropertyUnit = "(T1.U_Z_PROITEMCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "22") & "'"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "26") = "" Then
                strPropertyUnit = strPropertyUnit & " and 1=1)"
            Else
                strPropertyUnit = strPropertyUnit & " and T1.U_Z_PROITEMCODE <='" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "')"
            End If

            strProperty = "(" & strProperty & ")"
            strPropertyUnit = "(" & strPropertyUnit & ")"

            strStartDate = intYear.ToString("0000") & "-" & intMonth.ToString("00") & "-01"
            Dim strEndday As String
            Select Case intMonth
                Case 1, 3, 5, 7, 8, 10, 12
                    strEndday = "-31"
                Case 4, 6, 9, 11
                    strEndday = "-30"
                Case 2
                    strEndday = "-28"
            End Select
            'strEndDate = intYear.ToString("0000") & "-" & intMonth.ToString("00") & "-15"
            strEndDate = intYear.ToString("0000") & "-" & intMonth.ToString("00") & strEndday
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If aForm.PaneLevel = 3 Then 'Update Expenses allocation
                oExpGrid = aForm.Items.Item("15").Specific
                oUserTable = oApplication.Company.UserTables.Item("Z_OBILL2")
                Dim dblamt, dblSq As Double
                Dim strExpCode As String
                For intLoop1 As Integer = 0 To oExpGrid.DataTable.Rows.Count - 1
                    strCode = oExpGrid.DataTable.GetValue("Code", intLoop1)
                    If oUserTable.GetByKey(strCode) Then
                        'strCode = oApplication.Utilities.getMaxCode("@Z_OBILL2", "Code")
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode & "N"
                        oUserTable.UserFields.Fields.Item("U_Year").Value = intYear
                        oUserTable.UserFields.Fields.Item("U_Month").Value = intMonth
                        strExpCode = oExpGrid.DataTable.GetValue("U_Z_CODE", intLoop1)
                        oUserTable.UserFields.Fields.Item("U_Z_TotalSq").Value = oExpGrid.DataTable.GetValue("U_Z_TOTALSQ", intLoop1)
                        oUserTable.UserFields.Fields.Item("U_Z_AMOUNT").Value = oExpGrid.DataTable.GetValue("U_Z_AMOUNT", intLoop1)
                        dblAmount = oExpGrid.DataTable.GetValue("U_Z_AMOUNT", intLoop1)
                        dblSq = oExpGrid.DataTable.GetValue("U_Z_TOTALSQ", intLoop1)
                        dblAmount = dblAmount / dblSq
                        oUserTable.UserFields.Fields.Item("U_Z_Rate").Value = dblAmount
                        If oUserTable.Update <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        Else
                            oTemp.DoQuery("Update [@Z_OBILL1] set U_Z_RATE=" & dblAmount & " where U_Z_CODE='" & strExpCode & "' and U_Year=" & intYear & " and U_Month=" & intMonth)
                        End If
                    End If
                Next

            ElseIf aForm.PaneLevel = 2 Then 'add expenses allocation
                oTemp1.DoQuery("Select * from [@Z_OBILL]  where U_Invoiced='Y' and  U_Month=" & intMonth & " and U_Year=" & intYear)
                blnRecordExists = False
                If oTemp1.RecordCount > 0 Then
                    blnRecordExists = True
                    If oApplication.SBO_Application.MessageBox("Bill Generation already completed for this Selected year and month. Do you want to continue to generate unbilled Contracts?", , "Yes", "No") = 2 Then
                        aForm.Freeze(False)
                        Return False
                    End If
                End If
                oTemp1.DoQuery("Select * from [@Z_OBILL2]  where U_Month=" & intMonth & " and U_Year=" & intYear)
                If oTemp1.RecordCount <= 0 Then
                    strSQL = "SELECT isnull(sum( T1.[U_Z_SPACE]),0) FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where isnull(T0.U_Z_STATUS,'PEN')='AGR' and  (" & strProperty & " and " & strPropertyUnit & ") and ('" & strStartDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"
                    strSQL = strSQL & " or ('" & strEndDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"
                    oTemp1.DoQuery(strSQL)
                    Dim dblTotalSqlMeter As Double
                    dblTotalSqlMeter = oTemp1.Fields.Item(0).Value
                    If dblTotalSqlMeter <= 0 Then
                        dblTotalSqlMeter = 1
                        'oApplication.Utilities.Message("No Contract available for selected month and year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        'aForm.Freeze(False)
                        'Return False
                    End If
                    strSQL = "Select * from [@Z_OEXP] where U_Z_TYPE='S' "
                    oTemp1.DoQuery(strSQL)
                    For intLoop As Integer = 0 To oTemp1.RecordCount - 1
                        oUserTable = oApplication.Company.UserTables.Item("Z_OBILL2")
                        strCode = oApplication.Utilities.getMaxCode("@Z_OBILL2", "Code")
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode & "N"
                        oUserTable.UserFields.Fields.Item("U_Year").Value = intYear
                        oUserTable.UserFields.Fields.Item("U_Month").Value = intMonth
                        oUserTable.UserFields.Fields.Item("U_Z_CODE").Value = oTemp1.Fields.Item("U_Z_CODE").Value
                        oUserTable.UserFields.Fields.Item("U_Z_NAME").Value = oTemp1.Fields.Item("U_Z_NAME").Value
                        oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = oTemp1.Fields.Item("U_Z_GLACC").Value
                        oUserTable.UserFields.Fields.Item("U_Z_TotalSq").Value = dblTotalSqlMeter
                        oUserTable.UserFields.Fields.Item("U_Z_RATE").Value = oTemp1.Fields.Item("U_Z_RATE").Value
                        oUserTable.UserFields.Fields.Item("U_Z_AMOUNT").Value = oTemp1.Fields.Item("U_Z_RATE").Value
                        If oUserTable.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If
                        oTemp1.MoveNext()
                    Next
                End If
                oTemp1.DoQuery("Update [@Z_OBILL2] set U_Z_RATE=isnull(U_Z_AMOUNT,1)/isnull(U_Z_TOTALSQ,1) where U_Month=" & intMonth & " and U_Year=" & intYear)
                oHeaderGrid = aForm.Items.Item("11").Specific
                oLineGird = aForm.Items.Item("12").Specific
                oExpGrid = aForm.Items.Item("15").Specific
                oTemp1.DoQuery("Update [@Z_OBILL2] set U_Z_RATE=isnull(U_Z_AMOUNT,1)/isnull(U_Z_TOTALSQ,1) where U_Month=" & intMonth & " and U_Year=" & intYear)
                oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
                oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
                oExpGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL2]  where  U_Month=" & intMonth & " and U_Year=" & intYear)
                Formatgrid(aForm, "Payroll")
                If blnRecordExists = True Then
                    oExpGrid.Columns.Item("U_Z_AMOUNT").Editable = False
                Else
                    oExpGrid.Columns.Item("U_Z_AMOUNT").Editable = True
                End If

            End If

            oHeaderGrid = aForm.Items.Item("11").Specific
            oLineGird = aForm.Items.Item("12").Specific
            oExpGrid = aForm.Items.Item("15").Specific
            ' If oTemp.RecordCount > 0 Then
            oTemp.DoQuery("Select * from [@Z_OBILL] where U_Month=" & intMonth & " and U_Year=" & intYear)
            If oTemp.RecordCount > 0 And aForm.PaneLevel = 3 Then
                If oTemp.Fields.Item("U_Invoiced").Value = "Y" Then
                    If oApplication.SBO_Application.MessageBox("Bill Generation already completed for this Selected year and month. Do you want to continue for new contracts?", , "Yes", "No") = 2 Then
                        aForm.Freeze(False)
                        Return False
                    End If
                Else
                    oTemp1.DoQuery("Delete from [@Z_OBILL1] where U_Month=" & intMonth & " and U_Year=" & intYear)
                    oTemp1.DoQuery("Delete from [@Z_OBILL] where U_Invoiced<>'Y' and  U_Month=" & intMonth & " and U_Year=" & intYear)
                End If
            End If
            oTemp.DoQuery("Select * from [@Z_OBILL] where U_Month=" & intMonth & " and U_Year=" & intYear)
            If aForm.PaneLevel = 3 Then
                If 1 = 2 Then ' oTemp.RecordCount > 0 Then
                    oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
                    oExpGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL2]  where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    Formatgrid(aForm, "Payroll")
                    aForm.Items.Item("5").Enabled = False
                Else
                    Dim strExistingBills As String
                    Dim strunitCondition1 As String = "(Select T1.U_Z_PROITEMCODE from [@Z_PROPUNIT] T1 where " & strProperty & " and " & strPropertyUnit & ")"

                    Dim strunitCondition2 = "(Select T1.U_Z_ExtProNo from [@Z_PROPUNIT] T1 where " & strProperty & " and " & strPropertyUnit & ")"

                    strExistingBills = "Select Code from [@Z_OBILL] where U_INVOICED<>'Y' and  U_Month=" & intMonth & " and U_Year=" & intYear & " and U_UNITCODE IN " & strunitCondition1
                    oTemp1.DoQuery(strExistingBills)
                    If oTemp1.RecordCount > 0 Then
                        oTemp1.DoQuery("Delete from [@Z_OBILL1] where U_Month=" & intMonth & " and U_Year=" & intYear & " and U_Z_REFNO in (" & strExistingBills & ")")
                    Else
                        oTemp1.DoQuery("Delete from [@Z_OBILL1] where U_Month=" & intMonth & " and U_Year=" & intYear)
                    End If
                    oTemp1.DoQuery("Delete from [@Z_OBILL] where U_INVOICED<>'Y' and U_Month=" & intMonth & " and U_Year=" & intYear & " and U_UNITCODE IN " & strunitCondition1)

                    aForm.Items.Item("5").Enabled = True

                    strExistingBills = "Select U_ContractID from [@Z_OBILL] where U_INVOICED='Y' and  U_Month=" & intMonth & " and U_Year=" & intYear & " and U_UNITCODE IN " & strunitCondition1

                    strSQL = "SELECT T0.[DocEntry], T0.[U_Z_UNITCODE], T1.[U_Z_SPACE], T0.[U_Z_TENCODE], T0.[U_Z_ANNUALRENT], T0.[U_Z_PAYTRMS], T0.[U_Z_CHGMONTH],T0.[U_Z_ANNUALRENT]/T0.[U_Z_CHGMONTH],CASE T0.U_Z_TYPE  when 'T' then T0.U_Z_ACCTCODE1 else T0.U_Z_ACCTCODE1  end  ,'N', 0.000,0.000 ,T0.[U_Z_ConNo] 'ContrctID' ,T0.U_Z_ProType 'PropertyType',T0.U_Z_CommAc 'CommissionAccount',T0.U_Z_Comm 'ComPer', T0.U_Z_OwnerCode 'Owner' ,T0.U_Z_SeqNo 'Seq' ,T0.U_Z_CntNo ,T1.U_Z_ExtProNo,isnull(T0.U_Z_IsCommission,'N') 'IsComReq',T1.U_Z_PROPCODE, 'Property' FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE]  inner Join OCRD T2 on T2.Cardcode=T0.U_Z_TENCODE where isnull(T0.U_Z_STATUS,'PEN')='AGR'  and  (" & strProperty & " and " & strPropertyUnit & ")  and ( ('" & strStartDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"
                    strSQL = strSQL & ") and T0.DocEntry not in (" & strExistingBills & ")"



                    strSQL = strSQL & " Union  SELECT T0.[DocEntry], T0.[U_Z_UNITCODE], T1.[U_Z_SPACE], T0.[U_Z_TENCODE], T0.[U_Z_ANNUALRENT], T0.[U_Z_PAYTRMS], T0.[U_Z_CHGMONTH],T0.[U_Z_ANNUALRENT]/T0.[U_Z_CHGMONTH],CASE T0.U_Z_TYPE  when 'T' then T0.U_Z_ACCTCODE1 else T0.U_Z_ACCTCODE1  end  ,'N', 0.000,0.000 ,T0.[U_Z_ConNo] 'ContrctID' ,T0.U_Z_ProType 'PropertyType',T0.U_Z_CommAc 'CommissionAccount',T0.U_Z_Comm 'ComPer', T0.U_Z_OwnerCode 'Owner' ,T0.U_Z_SeqNo 'Seq' ,T0.U_Z_CntNo,T1.U_Z_ExtProNo ,isnull(T0.U_Z_IsCommission,'N') 'IsComReq',T1.U_Z_PROPCODE, 'Property' FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE]  inner Join OCRD T2 on T2.Cardcode=T0.U_Z_TENCODE where isnull(T0.U_Z_STATUS,'PEN')='AGR'  and  (" & strProperty & " and " & strPropertyUnit & ")  and "
                    strSQL = strSQL & "  ('" & strEndDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate)) and T0.DocEntry not in (" & strExistingBills & ")"


                    oTemp1.DoQuery(strSQL)
                    Dim intDocEntry As Integer
                    For intRow1 As Integer = 0 To oTemp1.RecordCount - 1
                        oUserTable = oApplication.Company.UserTables.Item("Z_OBILL")
                        strCode = oApplication.Utilities.getMaxCode("@Z_OBILL", "Code")
                        intDocEntry = oTemp1.Fields.Item("DocEntry").Value
                        Try
                            aForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        Catch ex As Exception

                        End Try
                        oUserTable.Code = strCode
                        oUserTable.Name = strCode & "N"
                        strECode = oTemp1.Fields.Item(0).Value
                        strSpace = oTemp1.Fields.Item(2).Value
                        If strSpace <> "" Then
                            Try
                                dblSpace = CDbl(strSpace)
                            Catch ex As Exception
                                dblSpace = 0
                            End Try
                            If dblSpace <= 0 Then
                                dblSpace = 1
                            End If
                        End If

                        oUserTable.UserFields.Fields.Item("U_DocDate").Value = dtPostDate
                        oUserTable.UserFields.Fields.Item("U_Month").Value = intMonth
                        oUserTable.UserFields.Fields.Item("U_Year").Value = intYear
                        oUserTable.UserFields.Fields.Item("U_CntNo").Value = oTemp1.Fields.Item("ContrctID").Value
                        oUserTable.UserFields.Fields.Item("U_ContractNumber").Value = oTemp1.Fields.Item("U_Z_CntNo").Value
                        oUserTable.UserFields.Fields.Item("U_ContractID").Value = oTemp1.Fields.Item(0).Value
                        oUserTable.UserFields.Fields.Item("U_Seq").Value = oTemp1.Fields.Item("Seq").Value
                        oUserTable.UserFields.Fields.Item("U_UnitCode").Value = oTemp1.Fields.Item(1).Value
                        oUserTable.UserFields.Fields.Item("U_ExtUnitCode").Value = oTemp1.Fields.Item("U_Z_ExtProNo").Value

                        'If oTemp1.Fields.Item("U_Z_ExtProNo").Value = "" Then
                        '    oUserTable.UserFields.Fields.Item("U_UnitCode").Value = oTemp1.Fields.Item(1).Value
                        'Else
                        '    oUserTable.UserFields.Fields.Item("U_UnitCode").Value = oTemp1.Fields.Item("U_Z_ExtProNo").Value
                        'End If

                        oUserTable.UserFields.Fields.Item("U_Space").Value = oTemp1.Fields.Item(2).Value
                        oUserTable.UserFields.Fields.Item("U_Annualrent").Value = oTemp1.Fields.Item(4).Value
                        oUserTable.UserFields.Fields.Item("U_PayTrms").Value = oTemp1.Fields.Item(5).Value
                        oUserTable.UserFields.Fields.Item("U_ChgMonth").Value = oTemp1.Fields.Item(6).Value
                        Dim dblMonthRent, dblCommPer, dblCommissionAmount As Double
                        Dim oRec1 As SAPbobsCOM.Recordset
                        oRec1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRec1.DoQuery("Select * from [@Z_CONINS] where U_Z_ConID=" & intDocEntry & " and U_Z_Month=" & intMonth & " and U_Z_Year=" & intYear)
                        If oRec1.RecordCount > 0 Then
                            dblMonthRent = oRec1.Fields.Item("U_Z_Amount").Value
                        Else
                            dblMonthRent = oTemp1.Fields.Item(7).Value
                        End If
                        dblCommPer = oTemp1.Fields.Item("ComPer").Value
                        dblCommissionAmount = dblMonthRent * dblCommPer / 100
                        If oTemp1.Fields.Item("IsComReq").Value = "N" Then
                            dblCommissionAmount = 0
                        End If
                        'T0.Z_ProType 'PropertyType',T0.U_Z_CommAc 'CommissionAccount',T0.U_Z_ConNo 'ComPer', T0.U_Z_OwnerCode 'Owner' 
                        If oTemp1.Fields.Item("PropertyType").Value = "A" Then
                            oUserTable.UserFields.Fields.Item("U_CardCode").Value = oTemp1.Fields.Item(3).Value
                            oUserTable.UserFields.Fields.Item("U_MonthRent").Value = dblMonthRent ' oTemp1.Fields.Item(7).Value
                            oUserTable.UserFields.Fields.Item("U_RentGL").Value = oTemp1.Fields.Item(8).Value
                            oUserTable.UserFields.Fields.Item("U_Remarks").Value = "Monthly Rental  :"
                            oUserTable.UserFields.Fields.Item("U_Commission").Value = dblCommissionAmount
                            oUserTable.UserFields.Fields.Item("U_ComPer").Value = dblCommPer
                            oUserTable.UserFields.Fields.Item("U_CommGL").Value = oTemp1.Fields.Item("CommissionAccount").Value
                            oUserTable.UserFields.Fields.Item("U_OwnerCode").Value = oTemp1.Fields.Item("Owner").Value
                            oUserTable.UserFields.Fields.Item("U_Z_ProType").Value = "A"
                        Else
                            oUserTable.UserFields.Fields.Item("U_CardCode").Value = oTemp1.Fields.Item(3).Value
                            oUserTable.UserFields.Fields.Item("U_MonthRent").Value = dblMonthRent ' oTemp1.Fields.Item(7).Value
                            oUserTable.UserFields.Fields.Item("U_Commission").Value = dblCommissionAmount
                            oUserTable.UserFields.Fields.Item("U_CommGL").Value = oTemp1.Fields.Item("CommissionAccount").Value
                            oUserTable.UserFields.Fields.Item("U_RentGL").Value = oTemp1.Fields.Item(8).Value
                            oUserTable.UserFields.Fields.Item("U_Remarks").Value = "Monthly Rental : " '"Commission Amount  :"
                            oUserTable.UserFields.Fields.Item("U_ComPer").Value = dblCommPer
                            oUserTable.UserFields.Fields.Item("U_OwnerCode").Value = oTemp1.Fields.Item("Owner").Value
                            oUserTable.UserFields.Fields.Item("U_Z_ProType").Value = "T"
                        End If

                        oUserTable.UserFields.Fields.Item("U_Expenses").Value = 0
                        oUserTable.UserFields.Fields.Item("U_Invoiced").Value = "N" 'oTemp.Fields.Item().Value
                        If oUserTable.Add() <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Committrans("Cancel")
                            Return False
                        Else
                            blnLineExists = True
                            strRefCode = strCode
                            Dim strMonths As String
                            otemp2.DoQuery("select U_Z_CODE,U_Z_NAME,U_Z_GLACC,U_Z_TYPE,U_Z_RATE,U_Z_Months,U_Z_Frequency from [@Z_CONTRACT1] where DocEntry=" & CInt(strECode))
                            For intRow2 As Integer = 0 To otemp2.RecordCount - 1
                                oUsertable1 = oApplication.Company.UserTables.Item("Z_OBILL1")
                                strCode = oApplication.Utilities.getMaxCode("@Z_OBILL1", "Code")
                                strMonths = otemp2.Fields.Item(5).Value
                                If strMonths.Contains(MonthName(intMonth)) Then
                                    oUsertable1.Code = strCode
                                    oUsertable1.Name = strCode & "N"
                                    strECode = oTemp1.Fields.Item(0).Value
                                    oUsertable1.UserFields.Fields.Item("U_Z_RefNo").Value = strRefCode
                                    oUsertable1.UserFields.Fields.Item("U_Month").Value = intMonth
                                    oUsertable1.UserFields.Fields.Item("U_Year").Value = intYear
                                    oUsertable1.UserFields.Fields.Item("U_Z_CODE").Value = otemp2.Fields.Item(0).Value
                                    oUsertable1.UserFields.Fields.Item("U_Z_NAME").Value = otemp2.Fields.Item(1).Value
                                    oUsertable1.UserFields.Fields.Item("U_Z_GLACC").Value = otemp2.Fields.Item(2).Value
                                    oUsertable1.UserFields.Fields.Item("U_Z_TYPE").Value = otemp2.Fields.Item(3).Value
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    Dim dblRate As Double
                                    If otemp2.Fields.Item(3).Value = "S" Then
                                        otest.DoQuery("Select isnull(U_Z_RATE,0) from [@Z_OBILL2] where U_Month=" & intMonth & " and U_Year=" & intYear & " and U_Z_CODE='" & otemp2.Fields.Item(0).Value & "'")
                                        dblRate = oApplication.Utilities.getDocumentQuantity(otest.Fields.Item(0).Value)
                                        oUsertable1.UserFields.Fields.Item("U_Z_RATE").Value = otest.Fields.Item(0).Value
                                        dblAmount = dblSpace * dblRate
                                    Else
                                        dblRate = oApplication.Utilities.getDocumentQuantity(oTemp1.Fields.Item(4).Value)
                                        oUsertable1.UserFields.Fields.Item("U_Z_RATE").Value = oTemp1.Fields.Item(4).Value
                                        dblAmount = CDbl(otemp2.Fields.Item(4).Value)
                                    End If
                                    oUsertable1.UserFields.Fields.Item("U_Z_AMOUNT").Value = dblAmount
                                    If oUsertable1.Add <> 0 Then
                                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                        Committrans("Cancel")
                                        Return False
                                    End If
                                End If
                                otemp2.MoveNext()
                            Next
                            otemp2.DoQuery("Select sum(U_Z_AMOUNT) from [@Z_OBILL1] where U_Z_RefNo='" & strRefCode & "'")
                            otemp2.DoQuery("UPdate [@Z_OBILL] set  U_Expenses='" & otemp2.Fields.Item(0).Value & "' where Code='" & strRefCode & "'")
                            otemp2.DoQuery("UPdate [@Z_OBILL] set U_Total=isnull(U_MonthRent,0)+ isnull(U_Expenses,0)  where Code='" & strRefCode & "'")
                        End If
                        oTemp1.MoveNext()
                    Next

                    oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear & " ORDER BY U_ContractID,U_INVOICED")
                    Dim strsql12 As String
                    Dim strunitCondition As String = "(Select T1.U_Z_PROITEMCODE from [@Z_PROPUNIT] T1 where " & strProperty & " and " & strPropertyUnit & ")"
                    Dim strunitCondition11 As String = "(Select T1.U_Z_ExtProNo from [@Z_PROPUNIT] T1 where " & strProperty & " and " & strPropertyUnit & ")"

                    strsql12 = "SELECT T0.[Code], T0.[Name], T0.[U_YEAR], T0.[U_MONTH], T0.[U_CONTRACTID], T0.[U_CONTRACTNUMBER], T0.[U_SEQ], T0.[U_UNITCODE],T0.U_EXTUNITCODE , T0.[U_SPACE], T0.[U_CARDCODE], T0.[U_PAYTRMS], T0.[U_ANNUALRENT], T0.[U_CHGMONTH], T0.[U_MONTHRENT], T0.[U_RENTGL], T0.[U_EXPENSES], T0.[U_TOTAL], T0.[U_INVOICED], T0.[U_INVENTRY], T0.[U_INVNUMBER], T0.[U_Z_PROTYPE], T0.[U_REMARKS], T0.[U_COMPER], T0.[U_COMMISSION], T0.[U_COMMGL], T0.[U_OWNERCODE] FROM [dbo].[@Z_OBILL]  T0"
                    strsql12 = strsql12 & "  where  U_Month=" & intMonth & " and U_Year=" & intYear & " and (U_UNITCODE IN  " & strunitCondition & ")" '   or ( U_ExtUnitCode in " & strunitCondition11 & ")   ORDER BY U_ContractID"


                    strsql12 = strsql12 & " Union  SELECT T0.[Code], T0.[Name], T0.[U_YEAR], T0.[U_MONTH], T0.[U_CONTRACTID], T0.[U_CONTRACTNUMBER], T0.[U_SEQ], T0.[U_UNITCODE],T0.U_EXTUNITCODE , T0.[U_SPACE], T0.[U_CARDCODE], T0.[U_PAYTRMS], T0.[U_ANNUALRENT], T0.[U_CHGMONTH], T0.[U_MONTHRENT], T0.[U_RENTGL], T0.[U_EXPENSES], T0.[U_TOTAL], T0.[U_INVOICED], T0.[U_INVENTRY], T0.[U_INVNUMBER], T0.[U_Z_PROTYPE], T0.[U_REMARKS], T0.[U_COMPER], T0.[U_COMMISSION], T0.[U_COMMGL], T0.[U_OWNERCODE] FROM [dbo].[@Z_OBILL]  T0"
                    strsql12 = strsql12 & "  where  U_Month=" & intMonth & " and U_Year=" & intYear & " and  ( U_ExtUnitCode in " & strunitCondition11 & ")   ORDER BY U_ContractID"


                    oHeaderGrid.DataTable.ExecuteQuery(strsql12)



                    oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
                    oExpGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL2]  where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    Formatgrid(aForm, "Payroll")
                    aForm.Items.Item("5").Enabled = True
                    If blnLineExists = False Then
                        oApplication.Utilities.Message("No New Contract available for selected month and year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Items.Item("5").Enabled = False
                        aForm.Freeze(False)
                        '                        Return False
                    End If

                End If
            End If

            aForm.Freeze(False)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
            Return False
        End Try
    End Function

    Private Function GenerateInvoice(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim intMonth, intYear As Integer
        Dim strSQL, strRefCode, strCode, strCode1, strSpace As String
        Dim dblSpace, dblAmount As Double
        Dim blnLineExists As Boolean = False
        Dim oTemp, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        Dim oHeaderGrid, oLineGird As SAPbouiCOM.Grid
        Dim oUserTable, oUsertable1 As SAPbobsCOM.UserTable
        Dim strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim dtPostDate As Date
        oCombobox = aForm.Items.Item("7").Specific
        Try
            aForm.Freeze(True)
            Try
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intYear = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try
            Try
                oCombobox = aForm.Items.Item("9").Specific
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intMonth = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try

            'Dim strProperty, strPropertyUnit As String
            ''If oApplication.Utilities.getEdittextvalue(aForm, "20") = "" Then
            ''    strProperty = "(1=1)"
            ''Else
            ''    strProperty = "(T1.U_Z_PROPCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "20") & "')"
            ''End If

            ''If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
            ''    strPropertyUnit = "(1=1)"
            ''Else
            ''    strPropertyUnit = "(T1.U_Z_PROITEMCODE='" & oApplication.Utilities.getEdittextvalue(aForm, "22") & "')"
            ''End If

            'If oApplication.Utilities.getEdittextvalue(aForm, "20") = "" Then
            '    strProperty = "(1=1)"
            'Else
            '    strProperty = "(T1.U_Z_PROPCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "20") & "')"
            'End If

            'If oApplication.Utilities.getEdittextvalue(aForm, "24") = "" Then
            '    strProperty = strProperty & " and (1=1)"
            'Else
            '    strProperty = strProperty & " and (T1.U_Z_PROPCODE<='" & oApplication.Utilities.getEdittextvalue(aForm, "24") & "')"
            'End If

            'If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
            '    strPropertyUnit = "(1=1)"
            'Else
            '    strPropertyUnit = "(T1.U_Z_PROITEMCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "22") & "')"
            'End If

            'If oApplication.Utilities.getEdittextvalue(aForm, "26") = "" Then
            '    strPropertyUnit = strProperty & " and (1=1)"
            'Else
            '    strPropertyUnit = strProperty & " and (T1.U_Z_PROITEMCODE <='" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "')"
            'End If


            Dim strProperty, strPropertyUnit As String
            If oApplication.Utilities.getEdittextvalue(aForm, "20") = "" Then
                strProperty = " (1=1 "
            Else
                strProperty = "(T1.U_Z_PROPCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "20") & "'"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "24") = "" Then
                strProperty = strProperty & " and 1=1)"
            Else
                strProperty = strProperty & " and T1.U_Z_PROPCODE<='" & oApplication.Utilities.getEdittextvalue(aForm, "24") & "')"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "22") = "" Then
                strPropertyUnit = "(1=1 "
            Else
                strPropertyUnit = "(T1.U_Z_PROITEMCODE >='" & oApplication.Utilities.getEdittextvalue(aForm, "22") & "'"
            End If

            If oApplication.Utilities.getEdittextvalue(aForm, "26") = "" Then
                strPropertyUnit = strPropertyUnit & " and 1=1)"
            Else
                strPropertyUnit = strPropertyUnit & " and T1.U_Z_PROITEMCODE <='" & oApplication.Utilities.getEdittextvalue(aForm, "26") & "')"
            End If

            strProperty = "(" & strProperty & ")"
            strPropertyUnit = "(" & strPropertyUnit & ")"

            Dim blnErrorflag As Boolean = False
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery("Select * from [@Z_LogDetail]")
            If oTemp.RecordCount > 0 Then
                If oTemp.Fields.Item("U_Z_LOG_PATH").Value <> "" Then
                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value
                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\ImportLog_Invoice.txt"
                    sFailureLog = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\FailureLog.txt"
                Else
                    sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                    sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                End If
            Else
                sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
            End If
            'sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
            'sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
            '  sFailureLog = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice1.txt"

            If File.Exists(sPath) Then
                File.Delete(sPath)
            End If
            If File.Exists(sFailureLog) Then
                File.Delete(sFailureLog)
            End If
            Dim blnDPAvailable As Boolean = False
            blnErrorflag = False
            WriteErrorlog("Processing Invoice Posting : Month  : " & MonthName(intMonth) & "  - Year : " & intYear.ToString("0000"), sPath)
            WriteErrorlog("Processing Invoice Posting : Month  : " & MonthName(intMonth) & "  - Year : " & intYear.ToString("0000"), sFailureLog)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Company.StartTransaction()

            Dim strunitCondition As String = "(Select T1.U_Z_PROITEMCODE from [@Z_PROPUNIT] T1 where " & strProperty & " and " & strPropertyUnit & ")"



            oTemp.DoQuery("Select * from [@Z_OBILL] where isnull(U_Invoiced,'N')='N' and U_Month=" & intMonth & " and U_Year=" & intYear & " and U_UNITCODE IN " & strunitCondition)
            Dim oDoc, oDoc2 As SAPbobsCOM.Documents
            Dim oDoc11, oDoc12 As SAPbobsCOM.Documents
            Dim strCostCenter, strProject, strCardCode As String
            Dim oBP As SAPbobsCOM.BusinessPartners
            oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            For intRow As Integer = 0 To oTemp.RecordCount - 1
                Try
                    aForm.Items.Item("18").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                Catch ex As Exception

                End Try
                oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDoc2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                If oTemp.Fields.Item("U_Z_ProType").Value = "A" Then
                    strCardCode = oTemp.Fields.Item("U_CARDCODE").Value
                Else
                    'strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                    strCardCode = oTemp.Fields.Item("U_CARDCODE").Value
                End If
                dtPostDate = oTemp.Fields.Item("U_DocDate").Value
                WriteErrorlog("Processing Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sFailureLog)
                WriteErrorlog("Processing Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sPath)
                oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                oDoc.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value
                oDoc.DocDate = dtPostDate ' Now.Date
                Dim st As String = "Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'"
                otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where (U_Z_ExtProNo='" & oTemp.Fields.Item("U_UNITCODE").Value & "' or  U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "')")
                oDoc.Project = otemp4.Fields.Item(0).Value
                strProject = otemp4.Fields.Item(0).Value
                strCostCenter = otemp4.Fields.Item(1).Value
                If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    'oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If



                Dim oDPI, oDOC1 As SAPbobsCOM.Documents
                Dim intCount As Integer = 0
                Dim aCode As Integer
                Dim DblRental, dblDownPayment As Double
                Dim oDPRec As SAPbobsCOM.Recordset
                oDPRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                DblRental = oTemp.Fields.Item("U_MonthRent").Value
                aCode = oTemp.Fields.Item("U_ContractID").Value
                Dim blnSPlitInvice As Boolean = False
                otemp2.DoQuery("Select isnull(U_Z_SPLIT,'N') 'U_Z_SPLIT' from [@Z_CONTRACT] where Docentry=" & aCode)
                If otemp2.Fields.Item("U_Z_SPLIT").Value = "Y" Then
                    blnSPlitInvice = True
                End If
                dblDownPayment = 0
                If 1 = 2 Then ' oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                Else
                    'otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus='C' and CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                    otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus<>'C' and CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                    If otemp2.RecordCount > 0 Then
                        blnDPAvailable = True
                        blnErrorflag = True
                        WriteErrorlog("Error : --> There are  open Downpayment invoices for the  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sFailureLog)
                    Else
                        blnDPAvailable = False
                    End If
                    If blnDPAvailable = False Then
                        otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0),PaidSum,paidtodate from ODPI where docstatus='C' and  CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                        For intLoop As Integer = 0 To otemp2.RecordCount - 1
                            oDPI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                            If oDPI.GetByKey(otemp2.Fields.Item("DocEntry").Value) Then
                                If DblRental > 0 Then
                                    If 1 = 1 Then ' oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Paid Then
                                        dblDownPayment = otemp2.Fields.Item(2).Value '.DownPaymentAmount
                                        oDPRec.DoQuery("Select isnull(DpmAppl,0) from ODPI where DocEntry=" & oDPI.DocEntry)
                                        dblDownPayment = dblDownPayment - oDPRec.Fields.Item(0).Value
                                        If dblDownPayment >= DblRental Then
                                            dblDownPayment = DblRental
                                            DblRental = DblRental - dblDownPayment
                                        Else
                                            dblDownPayment = dblDownPayment
                                            DblRental = DblRental - dblDownPayment
                                        End If
                                        Dim intd As Integer = oDPI.DocEntry
                                        If oDPI.GetByKey(intd) Then
                                            If dblDownPayment > 0 Then
                                                oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                                                oDoc.DownPaymentsToDraw.Add()
                                                oDoc.DownPaymentsToDraw.SetCurrentLine(intCount)
                                                oDoc.DownPaymentsToDraw.DocEntry = oDPI.DocEntry
                                                oDoc.DownPaymentsToDraw.AmountToDraw = Math.Round(dblDownPayment, 2) ' oDPI.DownPaymentAmount
                                                intCount = intCount + 1
                                                blnDPAvailable = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            otemp2.MoveNext()
                        Next

                        Dim intInvLines As Integer = 0
                        If oTemp.Fields.Item("U_Z_ProType").Value = "A" Then
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                        Else
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                        End If

                        oDoc.Lines.ItemDescription = oTemp.Fields.Item("U_Remarks").Value & ":" & oTemp.Fields.Item("U_UNITCODE").Value
                        If oBP.GetByKey(strCardCode) Then
                            If oBP.VatGroup <> "" Then
                                oDoc.Lines.TaxCode = oBP.VatGroup
                            Else
                            End If
                        End If

                        If strCostCenter <> "" Then
                            oDoc.Lines.CostingCode = strCostCenter
                        End If
                        If strProject <> "" Then
                            oDoc.Lines.ProjectCode = strProject
                        End If

                        If blnSPlitInvice = True Then 'Split Expenses to separte Invoice
                            If oDoc.Add <> 0 Then
                                blnErrorflag = True
                                WriteErrorlog("Error --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                WriteErrorlog("Error  --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                            Else
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                                WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & " Invoice Number " & otemp2.Fields.Item("DocNum").Value, sPath)
                                WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & " Invoice Number " & otemp2.Fields.Item("DocNum").Value, sFailureLog)
                                'Commission Invoice posting for Tenant
                                If oTemp.Fields.Item("U_Z_ProType").Value = "T" And oTemp.Fields.Item("U_Commission").Value > 0 Then
                                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                                    strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                                    oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                                    oDoc.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value
                                    oDoc.DocDate = dtPostDate ' Now.Date
                                    otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                                    oDoc.Project = otemp4.Fields.Item(0).Value
                                    strProject = otemp4.Fields.Item(0).Value
                                    strCostCenter = otemp4.Fields.Item(1).Value
                                    oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_CommGL").Value)
                                    oDoc.Lines.LineTotal = oTemp.Fields.Item("U_Commission").Value
                                    oDoc.Lines.ItemDescription = "Commisstion Amount for property Unit :" & oTemp.Fields.Item("U_UNITCODE").Value
                                    If oBP.GetByKey(strCardCode) Then
                                        If oBP.VatGroup <> "" Then
                                            oDoc.Lines.TaxCode = oBP.VatGroup
                                        End If
                                    End If
                                    If strCostCenter <> "" Then
                                        oDoc.Lines.CostingCode = strCostCenter
                                    End If
                                    If strProject <> "" Then
                                        oDoc.Lines.ProjectCode = strProject
                                    End If
                                    oDPRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    DblRental = oTemp.Fields.Item("U_MonthRent").Value
                                    aCode = oTemp.Fields.Item("U_ContractID").Value
                                    dblDownPayment = 0
                                    If oDoc.Add <> 0 Then
                                        WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                        WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                                        blnErrorflag = True
                                    Else
                                        oApplication.Company.GetNewObjectCode(strDocNum)
                                        otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                        ' otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                                        Dim str As String = "Commission Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Invoice Number : " & otemp2.Fields.Item("DocNum").Value
                                        WriteErrorlog(str, sPath)
                                    End If
                                End If
                            End If

                            'split expense into separte Invoice
                            oDoc2.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                            oDoc2.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                            oDoc2.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                            oDoc2.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                            oDoc2.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                            oDoc2.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value
                            oDoc2.DocDate = dtPostDate ' Now.Date
                            Dim st1 As String = "Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'"
                            otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where (U_Z_ExtProNo='" & oTemp.Fields.Item("U_UNITCODE").Value & "' or  U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "')")
                            oDoc2.Project = otemp4.Fields.Item(0).Value
                            strProject = otemp4.Fields.Item(0).Value
                            strCostCenter = otemp4.Fields.Item(1).Value
                            If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                                'oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                                oDoc2.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                            End If
                            intInvLines = 0
                            oTemp1.DoQuery("Select * from [@Z_OBILL1] where U_Z_RefNo='" & oTemp.Fields.Item("Code").Value & "'")
                            For intLoop As Integer = 0 To oTemp1.RecordCount - 1
                                If intInvLines > 0 Then
                                    oDoc2.Lines.Add()
                                End If
                                oDoc2.Lines.SetCurrentLine(intInvLines)
                                oDoc2.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp1.Fields.Item("U_Z_GLACC").Value)
                                oDoc2.Lines.ItemDescription = "Monthly Expenses  : " & oTemp1.Fields.Item("U_Z_NAME").Value
                                oDoc2.Lines.LineTotal = oTemp1.Fields.Item("U_Z_AMOUNT").Value
                                If strCostCenter <> "" Then
                                    oDoc2.Lines.CostingCode = strCostCenter
                                End If
                                If strProject <> "" Then
                                    oDoc2.Lines.ProjectCode = strProject
                                End If
                                If oBP.GetByKey(strCardCode) Then
                                    If oBP.VatGroup <> "" Then
                                        oDoc2.Lines.TaxCode = oBP.VatGroup
                                    Else

                                    End If
                                End If
                                intInvLines = intInvLines + 1
                                oTemp1.MoveNext()
                            Next
                            If intInvLines > 0 Then
                                If oDoc2.Add() <> 0 Then
                                    blnErrorflag = True
                                    WriteErrorlog("Error --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                    WriteErrorlog("Error  --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                                Else
                                    Dim strDocNum As String
                                    oApplication.Company.GetNewObjectCode(strDocNum)
                                    otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                End If
                            End If

                        Else
                            intInvLines = intInvLines + 1
                            oTemp1.DoQuery("Select * from [@Z_OBILL1] where U_Z_RefNo='" & oTemp.Fields.Item("Code").Value & "'")
                            For intLoop As Integer = 0 To oTemp1.RecordCount - 1
                                If intInvLines > 0 Then
                                    oDoc.Lines.Add()
                                End If
                                oDoc.Lines.SetCurrentLine(intInvLines)
                                oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp1.Fields.Item("U_Z_GLACC").Value)
                                oDoc.Lines.ItemDescription = "Monthly Expenses  : " & oTemp1.Fields.Item("U_Z_NAME").Value
                                oDoc.Lines.LineTotal = oTemp1.Fields.Item("U_Z_AMOUNT").Value
                                If strCostCenter <> "" Then
                                    oDoc.Lines.CostingCode = strCostCenter
                                End If
                                If strProject <> "" Then
                                    oDoc.Lines.ProjectCode = strProject
                                End If
                                If oBP.GetByKey(strCardCode) Then
                                    If oBP.VatGroup <> "" Then
                                        oDoc.Lines.TaxCode = oBP.VatGroup
                                    Else

                                    End If
                                End If
                                intInvLines = intInvLines + 1
                                oTemp1.MoveNext()
                            Next
                            If oDoc.Add <> 0 Then
                                blnErrorflag = True
                                WriteErrorlog("Error --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                WriteErrorlog("Error  --> Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                            Else
                                Dim strDocNum As String
                                oApplication.Company.GetNewObjectCode(strDocNum)
                                otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                                WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & " Invoice Number " & otemp2.Fields.Item("DocNum").Value, sPath)
                                WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & " Invoice Number " & otemp2.Fields.Item("DocNum").Value, sFailureLog)
                                'Commission Invoice posting for Tenant
                                If oTemp.Fields.Item("U_Z_ProType").Value = "T" And oTemp.Fields.Item("U_Commission").Value > 0 Then
                                    oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                    oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                                    strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                                    oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                                    oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                                    oDoc.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value
                                    oDoc.DocDate = dtPostDate ' Now.Date
                                    otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                                    oDoc.Project = otemp4.Fields.Item(0).Value
                                    strProject = otemp4.Fields.Item(0).Value
                                    strCostCenter = otemp4.Fields.Item(1).Value
                                    oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_CommGL").Value)
                                    oDoc.Lines.LineTotal = oTemp.Fields.Item("U_Commission").Value

                                    oDoc.Lines.ItemDescription = "Commisstion Amount for property Unit :" & oTemp.Fields.Item("U_UNITCODE").Value
                                    If oBP.GetByKey(strCardCode) Then
                                        If oBP.VatGroup <> "" Then
                                            oDoc.Lines.TaxCode = oBP.VatGroup
                                        End If
                                    End If
                                    If strCostCenter <> "" Then
                                        oDoc.Lines.CostingCode = strCostCenter
                                    End If
                                    If strProject <> "" Then
                                        oDoc.Lines.ProjectCode = strProject
                                    End If
                                    oDPRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    DblRental = oTemp.Fields.Item("U_MonthRent").Value
                                    aCode = oTemp.Fields.Item("U_ContractID").Value
                                    dblDownPayment = 0
                                    If oDoc.Add <> 0 Then
                                        WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                        WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                                        blnErrorflag = True
                                    Else
                                        oApplication.Company.GetNewObjectCode(strDocNum)
                                        otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                        ' otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                                        Dim str As String = "Commission Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Invoice Number : " & otemp2.Fields.Item("DocNum").Value
                                        WriteErrorlog(str, sPath)
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
                oTemp.MoveNext()
            Next
           
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            Committrans("Add")
            oHeaderGrid = aForm.Items.Item("11").Specific
            oLineGird = aForm.Items.Item("12").Specific
            oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear & " and U_UNITCODE IN " & strunitCondition)
            oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
            Formatgrid(aForm, "Payroll")
            aForm.Freeze(False)
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            If blnErrorflag = True Then
                If oApplication.Company.InTransaction() Then
                    'oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
                WriteErrorlog("Process completed with errors, Some Contracts are not posted.check the log file ", sFailureLog)
                WriteErrorlog("Please correct the errors and Re-run the Bill Generation Wizard. ", sFailureLog)
                aForm.Freeze(False)
                aForm.Items.Item("5").Enabled = True
                Return False
            End If
            aForm.Items.Item("5").Enabled = False
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            WriteErrorlog(" Error : " & oApplication.Company.GetLastErrorDescription, sPath)
            WriteErrorlog(" Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            aForm.Freeze(False)
            Return False
        End Try
    End Function

    Private Function GenerateInvoice_old(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim intMonth, intYear As Integer
        Dim strSQL, strRefCode, strCode, strCode1, strSpace As String
        Dim dblSpace, dblAmount As Double
        Dim blnLineExists As Boolean = False
        Dim oTemp, oTemp1, otemp2, otemp3, otemp4 As SAPbobsCOM.Recordset
        Dim oHeaderGrid, oLineGird As SAPbouiCOM.Grid
        Dim oUserTable, oUsertable1 As SAPbobsCOM.UserTable
        Dim strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim dtPostDate As Date
        oCombobox = aForm.Items.Item("7").Specific
        Try
            aForm.Freeze(True)
            Try
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intYear = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try
            Try
                oCombobox = aForm.Items.Item("9").Specific
                If oCombobox.Selected.Description = "" Then
                    oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Return False
                Else
                    intMonth = CInt(oCombobox.Selected.Value)
                End If
            Catch ex As Exception
                oApplication.Utilities.Message("Select the Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Return False
            End Try

            Dim blnErrorflag As Boolean = False

            sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
            sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
            '  sFailureLog = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice1.txt"

            If File.Exists(sPath) Then
                File.Delete(sPath)
            End If
            If File.Exists(sFailureLog) Then
                File.Delete(sFailureLog)
            End If
            Dim blnDPAvailable As Boolean = False
            blnErrorflag = False
            WriteErrorlog("Processing Invoice Posting : Month  : " & MonthName(intMonth) & "  - Year : " & intYear.ToString("0000"), sPath)
            WriteErrorlog("Processing Invoice Posting : Month  : " & MonthName(intMonth) & "  - Year : " & intYear.ToString("0000"), sFailureLog)
            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp3 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            oApplication.Company.StartTransaction()
            oTemp.DoQuery("Select * from [@Z_OBILL] where isnull(U_Invoiced,'N')='N' and U_Month=" & intMonth & " and U_Year=" & intYear)
            Dim oDoc, oDoc2 As SAPbobsCOM.Documents
            Dim strCostCenter, strProject, strCardCode As String
            Dim oBP As SAPbobsCOM.BusinessPartners
            oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
            For intRow As Integer = 0 To oTemp.RecordCount - 1
                oApplication.Utilities.Message("Processing...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDoc2 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                If oTemp.Fields.Item("U_Z_ProType").Value = "A" Then
                    strCardCode = oTemp.Fields.Item("U_CARDCODE").Value
                Else
                    'strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                    strCardCode = oTemp.Fields.Item("U_CARDCODE").Value
                End If
                dtPostDate = oTemp.Fields.Item("U_DocDate").Value
                WriteErrorlog("Processing Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sFailureLog)
                WriteErrorlog("Processing Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sPath)


                oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                oDoc.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value
                oDoc.DocDate = dtPostDate ' Now.Date
                otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                oDoc.Project = otemp4.Fields.Item(0).Value
                strProject = otemp4.Fields.Item(0).Value
                strCostCenter = otemp4.Fields.Item(1).Value
                If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    'oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If



                Dim oDPI, oDOC1 As SAPbobsCOM.Documents
                Dim intCount As Integer = 0
                Dim aCode As Integer
                Dim DblRental, dblDownPayment As Double
                Dim oDPRec As SAPbobsCOM.Recordset
                oDPRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                DblRental = oTemp.Fields.Item("U_MonthRent").Value
                aCode = oTemp.Fields.Item("U_ContractID").Value
                dblDownPayment = 0
                If 1 = 2 Then ' oTemp.Fields.Item("U_Z_ProType").Value = "T" Then

                Else
                    'otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus='C' and CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                    otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus<>'C' and CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                    If otemp2.RecordCount > 0 Then
                        blnDPAvailable = True
                        blnErrorflag = True
                        WriteErrorlog("There are some open Downpayment invoices for the  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                    Else
                        blnDPAvailable = False
                    End If
                    If blnDPAvailable = False Then
                        otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0),PaidSum,paidtodate from ODPI where docstatus='C' and  CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                        For intLoop As Integer = 0 To otemp2.RecordCount - 1
                            oDPI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                            If oDPI.GetByKey(otemp2.Fields.Item("DocEntry").Value) Then
                                If DblRental > 0 Then
                                    If 1 = 1 Then ' oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Paid Then
                                        dblDownPayment = otemp2.Fields.Item(2).Value '.DownPaymentAmount
                                        oDPRec.DoQuery("Select isnull(DpmAppl,0) from ODPI where DocEntry=" & oDPI.DocEntry)
                                        dblDownPayment = dblDownPayment - oDPRec.Fields.Item(0).Value
                                        If dblDownPayment >= DblRental Then
                                            dblDownPayment = DblRental
                                            DblRental = DblRental - dblDownPayment
                                        Else
                                            dblDownPayment = dblDownPayment
                                            DblRental = DblRental - dblDownPayment
                                        End If
                                        Dim intd As Integer = oDPI.DocEntry
                                        If oDPI.GetByKey(intd) Then
                                            If dblDownPayment > 0 Then
                                                oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                                                oDoc.DownPaymentsToDraw.Add()
                                                oDoc.DownPaymentsToDraw.SetCurrentLine(intCount)
                                                oDoc.DownPaymentsToDraw.DocEntry = oDPI.DocEntry
                                                oDoc.DownPaymentsToDraw.AmountToDraw = Math.Round(dblDownPayment, 2) ' oDPI.DownPaymentAmount
                                                intCount = intCount + 1
                                                blnDPAvailable = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                            otemp2.MoveNext()
                        Next

                        Dim intInvLines As Integer = 0

                        If oTemp.Fields.Item("U_Z_ProType").Value = "A" Then
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                        Else
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                        End If

                        oDoc.Lines.ItemDescription = oTemp.Fields.Item("U_Remarks").Value & ":" & oTemp.Fields.Item("U_UNITCODE").Value
                        If oBP.GetByKey(strCardCode) Then
                            If oBP.VatGroup <> "" Then
                                oDoc.Lines.TaxCode = oBP.VatGroup
                            Else
                                'Try
                                '    oDoc.Lines.TaxCode = "Exempt"
                                'Catch ex As Exception
                                '    oDoc.Lines.TaxCode = "Exempt"
                                'End Try
                            End If

                        End If
                        If strCostCenter <> "" Then
                            oDoc.Lines.CostingCode = strCostCenter
                        End If
                        If strProject <> "" Then
                            oDoc.Lines.ProjectCode = strProject
                        End If
                        intInvLines = intInvLines + 1


                        oTemp1.DoQuery("Select * from [@Z_OBILL1] where U_Z_RefNo='" & oTemp.Fields.Item("Code").Value & "'")
                        For intLoop As Integer = 0 To oTemp1.RecordCount - 1
                            If intInvLines > 0 Then
                                oDoc.Lines.Add()
                            End If
                            oDoc.Lines.SetCurrentLine(intLoop) ' + 1)
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp1.Fields.Item("U_Z_GLACC").Value)
                            oDoc.Lines.ItemDescription = "Monthly Expenses  : " & oTemp1.Fields.Item("U_Z_NAME").Value
                            oDoc.Lines.LineTotal = oTemp1.Fields.Item("U_Z_AMOUNT").Value
                            If strCostCenter <> "" Then
                                oDoc.Lines.CostingCode = strCostCenter
                            End If
                            If strProject <> "" Then
                                oDoc.Lines.ProjectCode = strProject
                            End If
                            If oBP.GetByKey(strCardCode) Then
                                If oBP.VatGroup <> "" Then
                                    oDoc.Lines.TaxCode = oBP.VatGroup
                                Else

                                End If
                            End If
                            oTemp1.MoveNext()
                        Next
                        If oDoc.Add <> 0 Then
                            blnErrorflag = True
                            WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                            WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                        Else
                            Dim strDocNum As String
                            oApplication.Company.GetNewObjectCode(strDocNum)
                            otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                            otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                            WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sPath)
                            WriteErrorlog("Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value, sFailureLog)
                            'Commission Invoice posting for Tenant
                            If oTemp.Fields.Item("U_Z_ProType").Value = "T" And oTemp.Fields.Item("U_Commission").Value > 0 Then
                                oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service

                                strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                                oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                                oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_CntNo").Value
                                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                                oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                                oDoc.UserFields.Fields.Item("U_SEQ").Value = oTemp.Fields.Item("U_Seq").Value

                                oDoc.DocDate = Now.Date
                                otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                                oDoc.Project = otemp4.Fields.Item(0).Value
                                strProject = otemp4.Fields.Item(0).Value
                                strCostCenter = otemp4.Fields.Item(1).Value
                                oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_CommGL").Value)
                                oDoc.Lines.LineTotal = oTemp.Fields.Item("U_Commission").Value

                                oDoc.Lines.ItemDescription = "Commisstion Amount for property Unit :" & oTemp.Fields.Item("U_UNITCODE").Value
                                If oBP.GetByKey(strCardCode) Then
                                    If oBP.VatGroup <> "" Then
                                        oDoc.Lines.TaxCode = oBP.VatGroup
                                    End If
                                End If
                                If strCostCenter <> "" Then
                                    oDoc.Lines.CostingCode = strCostCenter
                                End If
                                If strProject <> "" Then
                                    oDoc.Lines.ProjectCode = strProject
                                End If
                                oDPRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                DblRental = oTemp.Fields.Item("U_MonthRent").Value
                                aCode = oTemp.Fields.Item("U_ContractID").Value
                                dblDownPayment = 0
                                'If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                                '    otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus='C' and CardCode='" & strCardCode & "' and  [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                                '    For intLoop As Integer = 0 To otemp2.RecordCount - 1
                                '        oDPI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                                '        If oDPI.GetByKey(otemp2.Fields.Item("DocEntry").Value) Then
                                '            If DblRental > 0 Then
                                '                If oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Close Or oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Paid Then
                                '                    dblDownPayment = oDPI.DownPaymentAmount
                                '                    oDPRec.DoQuery("Select isnull(DpmAppl,0) from ODPI where DocEntry=" & oDPI.DocEntry)
                                '                    dblDownPayment = dblDownPayment - oDPRec.Fields.Item(0).Value
                                '                    If dblDownPayment >= DblRental Then
                                '                        dblDownPayment = DblRental
                                '                        DblRental = DblRental - dblDownPayment
                                '                    Else
                                '                        dblDownPayment = dblDownPayment
                                '                        DblRental = DblRental - dblDownPayment
                                '                    End If
                                '                    oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                                '                    oDoc.DownPaymentsToDraw.Add()
                                '                    oDoc.DownPaymentsToDraw.SetCurrentLine(intCount)
                                '                    oDoc.DownPaymentsToDraw.DocEntry = oDPI.DocEntry
                                '                    oDoc.DownPaymentsToDraw.AmountToDraw = dblDownPayment ' oDPI.DownPaymentAmount
                                '                    intCount = intCount + 1
                                '                End If
                                '            End If
                                '        End If
                                '        otemp2.MoveNext()
                                '    Next
                                'End If
                                If oDoc.Add <> 0 Then
                                    WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
                                    WriteErrorlog("Error Creating Invoice  Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Error : " & oApplication.Company.GetLastErrorDescription, sPath)
                                    blnErrorflag = True
                                Else
                                    oApplication.Company.GetNewObjectCode(strDocNum)
                                    otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                                    ' otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
                                    Dim str As String = "Commission Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Invoice Number : " & otemp2.Fields.Item("DocNum").Value
                                    WriteErrorlog(str, sPath)
                                    '  WriteErrorlog("Commission  Invoice created successfully Contract Number : " & oTemp.Fields.Item("U_ContractNumber").Value & ": Invoice Number : " & otemp2.field.item("DocNum").value, sPath)
                                End If

                            End If
                        End If
                    End If
                    oTemp.MoveNext()
                End If

            Next
            If blnErrorflag = True Then
                If oApplication.Company.InTransaction() Then
                    'oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                End If
                WriteErrorlog("Process completed with errors, Some Contracts are not posted.check the log file ", sFailureLog)
                WriteErrorlog("Please correct the errors and Re-run the Bill Generation Wizard. ", sFailureLog)
                aForm.Freeze(False)
                Return False
            End If
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
            End If
            Committrans("Add")
            oHeaderGrid = aForm.Items.Item("11").Specific
            oLineGird = aForm.Items.Item("12").Specific
            oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
            oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
            Formatgrid(aForm, "Payroll")
            aForm.Items.Item("5").Enabled = False
            aForm.Freeze(False)
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            WriteErrorlog(" Error : " & oApplication.Company.GetLastErrorDescription, sPath)
            WriteErrorlog(" Error : " & oApplication.Company.GetLastErrorDescription, sFailureLog)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            aForm.Freeze(False)
            Return False
        End Try
    End Function
#End Region

#Region "Write into ErrorLog File"
    Private Sub WriteErrorHeader(ByVal apath As String)
        Dim aSw As System.IO.StreamWriter
        Dim aMessage As String
        aMessage = "FileName : " & apath
        If File.Exists(apath) Then
        End If
        aSw = New StreamWriter(sPath, True)
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
    Private Sub WriteErrorlog(ByVal aMessage As String, ByVal aPath As String)
        Dim aSw As System.IO.StreamWriter
        If File.Exists(aPath) Then
        End If
        aSw = New StreamWriter(aPath, True)
        aMessage = Now.ToString("dd-MM-yyyy hh:mm") & "--> " & aMessage
        aSw.WriteLine(aMessage)
        aSw.Flush()
        aSw.Close()
    End Sub
#End Region
#Region "AddtoUDT"
    Private Function AddtoUDT1(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode, strECode, strESocial, strEname, strETax, strGLAcc As String
        Dim OCHECKBOXCOLUMN As SAPbouiCOM.CheckBoxColumn
        oGrid = aform.Items.Item("5").Specific
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            If oGrid.DataTable.GetValue(2, intRow) <> "" Then
                strCode = oGrid.DataTable.GetValue(0, intRow)
                strECode = oGrid.DataTable.GetValue(2, intRow)
                strEname = oGrid.DataTable.GetValue(3, intRow)
                strGLAcc = oGrid.DataTable.GetValue(4, intRow)
                oCombobox = oGrid.Columns.Item(5)
                strESocial = oCombobox.GetSelectedValue(intRow).Value
                strETax = oGrid.DataTable.GetValue(6, intRow)
                'strbindesc = oGrid.DataTable.GetValue(5, intRow)
                oUserTable = oApplication.Company.UserTables.Item("Z_OBILL")
                If oGrid.DataTable.GetValue(0, intRow) = "" Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_OBILL", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_CODE").Value = oGrid.DataTable.GetValue(2, intRow).ToString.ToUpper()
                    oUserTable.UserFields.Fields.Item("U_Z_NAME").Value = (oGrid.DataTable.GetValue(3, intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_GLACC").Value = (oGrid.DataTable.GetValue(4, intRow))
                    oUserTable.UserFields.Fields.Item("U_Z_TYPE").Value = strESocial
                    oUserTable.UserFields.Fields.Item("U_Z_RATE").Value = strETax
                    If oUserTable.Add() <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Committrans("Cancel")
                        Return False
                    End If
                End If
            End If
        Next
        oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Committrans("Add")
        DataBind(aform)
    End Function
#End Region

#Region "CommitTrans"
    Private Sub Committrans(ByVal strChoice As String)
        Dim oTemprec, oItemRec As SAPbobsCOM.Recordset
        oTemprec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oItemRec = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If strChoice = "Cancel" Then
            oTemprec.DoQuery("Update [@Z_OBILL] set NAME=CODE where Name Like '%D'")
            oTemprec.DoQuery("Update [@Z_OBILL1] set NAME=CODE where Name Like '%D'")
            oTemprec.DoQuery("Delete from  [@Z_OBILL]  where NAME Like '%N'")
            oTemprec.DoQuery("Delete from  [@Z_OBILL1]  where NAME Like '%N'")
            oTemprec.DoQuery("Delete from  [@Z_OBILL2]  where NAME Like '%N'")
        Else
            oTemprec.DoQuery("Delete from  [@Z_OBILL]  where NAME Like '%D'")
            oTemprec.DoQuery("Update [@Z_OBILL] set NAME=CODE where Name Like '%N'")
            oTemprec.DoQuery("Update [@Z_OBILL1] set NAME=CODE where Name Like '%N'")
            oTemprec.DoQuery("Update [@Z_OBILL2] set NAME=CODE where Name Like '%N'")

        End If

    End Sub
#End Region
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal aForm As SAPbouiCOM.Form, ByVal aOption As String)
        Dim aGrid As SAPbouiCOM.Grid

        Select Case aOption
            Case "Load"
                aGrid = aForm.Items.Item("11").Specific
                aGrid.Columns.Item(0).TitleObject.Caption = "Year"
                aGrid.Columns.Item(1).TitleObject.Caption = "Month"
                aGrid.Columns.Item(2).TitleObject.Caption = "Contract ID"
                aGrid.Columns.Item(3).TitleObject.Caption = "UnitCode"
                aGrid.Columns.Item(4).TitleObject.Caption = "Tenent Code"
                aGrid.Columns.Item(5).TitleObject.Caption = "Annual Rent"
                aGrid.Columns.Item(6).TitleObject.Caption = "Number of Months"
                'oEditTextColumn = agrid.Columns.Item(0)
                'oEditTextColumn.LinkedObjectType = "171"
            Case "Payroll"



                'AddFields("Z_OBILL2", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
                'AddFields("Z_OBILL2", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)
                'AddFields("Z_OBILL2", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
                'AddFields("Z_OBILL2", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
                'AddFields("Z_OBILL2", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
                'AddFields("Z_OBILL2", "Z_AMOUNT", "Expense Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
                'AddFields("Z_OBILL2", "Z_TotalSq", "Total Sq.Meter", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
                'AddFields("Z_OBILL2", "Z_Rate", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

                aGrid = aForm.Items.Item("15").Specific
                aGrid.Columns.Item(0).TitleObject.Caption = "Code"
                aGrid.Columns.Item(0).Visible = False
                aGrid.Columns.Item(1).TitleObject.Caption = "Name"
                aGrid.Columns.Item(1).Visible = False
                aGrid.Columns.Item(2).TitleObject.Caption = "Year"
                aGrid.Columns.Item(2).Editable = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Month"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "ExpCode"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Exp.Name"
                aGrid.Columns.Item(5).Editable = False
                aGrid.Columns.Item(6).TitleObject.Caption = "G/L Account"
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "Total Amount"
                aGrid.Columns.Item(7).Editable = True
                oEditTextColumn = aGrid.Columns.Item(7)
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item(8).TitleObject.Caption = "Total Sq.Meter"
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).TitleObject.Caption = "Rate per Sq.Meter"
                aGrid.Columns.Item(9).Editable = False

                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None

                aGrid = aForm.Items.Item("11").Specific
                aGrid.Columns.Item("Code").TitleObject.Caption = "Code"
                aGrid.Columns.Item("Code").Visible = False
                aGrid.Columns.Item("Name").TitleObject.Caption = "Name"
                aGrid.Columns.Item("Name").Visible = False
                aGrid.Columns.Item("U_YEAR").TitleObject.Caption = "Year"
                aGrid.Columns.Item("U_YEAR").Visible = False
                aGrid.Columns.Item("U_MONTH").TitleObject.Caption = "Month"
                aGrid.Columns.Item("U_MONTH").Visible = False
                aGrid.Columns.Item("U_CONTRACTID").TitleObject.Caption = "Contract ID"
                aGrid.Columns.Item("U_CONTRACTID").Type = SAPbouiCOM.BoGridColumnType.gct_EditText
                oEditTextColumn = aGrid.Columns.Item("U_CONTRACTID")
                oEditTextColumn.LinkedObjectType = "@Z_CONTRACT"
                Try
                    aGrid.Columns.Item("U_CNTNO").TitleObject.Caption = "Contract No"
                    aGrid.Columns.Item("U_CNTNO").Type = SAPbouiCOM.BoGridColumnType.gct_EditText

                Catch ex As Exception

                End Try
               
                aGrid.Columns.Item("U_UNITCODE").TitleObject.Caption = "Unit Code"
                oEditTextColumn = aGrid.Columns.Item("U_UNITCODE")
                Try
                    aGrid.Columns.Item("U_EXTUNITCODE").TitleObject.Caption = "Ext.Unit Code"
                Catch ex As Exception

                End Try
               
                '  oEditTextColumn.LinkedObjectType = "4"
                aGrid.Columns.Item("U_SPACE").TitleObject.Caption = "Space"
                aGrid.Columns.Item("U_SPACE").Visible = False
                oEditTextColumn = aGrid.Columns.Item("U_SPACE")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item("U_CARDCODE").TitleObject.Caption = "Tenent Code"
                oEditTextColumn = aGrid.Columns.Item("U_CARDCODE")
                oEditTextColumn.LinkedObjectType = "2"
                aGrid.Columns.Item("U_ANNUALRENT").TitleObject.Caption = "Annual Rent"
                aGrid.Columns.Item("U_ANNUALRENT").Visible = False
                aGrid.Columns.Item("U_PAYTRMS").TitleObject.Caption = "Payment Terms"
                Try


                    aGrid.Columns.Item("U_PAYTRMS").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    Dim otemp As SAPbobsCOM.Recordset
                    otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otemp.DoQuery("SELECT T0.[GroupNum], T0.[PymntGroup] FROM OCTG T0")
                    oCombo = aGrid.Columns.Item("U_PAYTRMS")
                    For introw As Integer = 0 To otemp.RecordCount - 1
                        oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                        otemp.MoveNext()
                    Next
                    oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                Catch ex As Exception

                End Try
                aGrid.Columns.Item("U_PAYTRMS").Visible = False
                aGrid.Columns.Item("U_CONTRACTNUMBER").TitleObject.Caption = "Contract Number with Seq.No"
                aGrid.Columns.Item("U_CONTRACTNUMBER").Editable = False
                aGrid.Columns.Item("U_CHGMONTH").TitleObject.Caption = "Number of Months"
                aGrid.Columns.Item("U_CHGMONTH").Visible = False
                aGrid.Columns.Item("U_MONTHRENT").TitleObject.Caption = "Monthly Rental"
                aGrid.Columns.Item("U_RENTGL").TitleObject.Caption = "Account Code"
                oEditTextColumn = aGrid.Columns.Item("U_RENTGL")
                oEditTextColumn.LinkedObjectType = "1"
                aGrid.Columns.Item("U_RENTGL").Visible = False
                aGrid.Columns.Item("U_EXPENSES").TitleObject.Caption = "Expenses"
                aGrid.Columns.Item("U_TOTAL").TitleObject.Caption = "Total"
                aGrid.Columns.Item("U_INVOICED").TitleObject.Caption = "Invoiced"
                aGrid.Columns.Item("U_INVENTRY").TitleObject.Caption = "Invoice DocEntry"
                oEditTextColumn = aGrid.Columns.Item("U_INVENTRY")
                oEditTextColumn.LinkedObjectType = SAPbobsCOM.BoObjectTypes.oInvoices
                aGrid.Columns.Item("U_INVNUMBER").TitleObject.Caption = "Invoice Document Number"
                Try
                    aGrid.Columns.Item("U_REMARKS").Visible = False
                    aGrid.Columns.Item("U_Z_PROTYPE").TitleObject.Caption = "Property Type"
                    aGrid.Columns.Item("U_Z_PROTYPE").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                    oCombo = aGrid.Columns.Item("U_Z_PROTYPE")
                    oCombo.ValidValues.Add("A", "Owned")
                    oCombo.ValidValues.Add("T", "Third Party")
                    oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
                    aGrid.Columns.Item("U_COMPER").TitleObject.Caption = "Commission Percentage"
                    aGrid.Columns.Item("U_COMMISSION").TitleObject.Caption = "Commission Amount"
                    aGrid.Columns.Item("U_COMMGL").TitleObject.Caption = "Commission Account Code"
                    aGrid.Columns.Item("U_COMMGL").Visible = False
                    aGrid.Columns.Item("U_OWNERCODE").TitleObject.Caption = "Owner Code"
                    aGrid.Columns.Item("U_SEQ").TitleObject.Caption = "Sequance Number"
                    aGrid.Columns.Item("U_SEQ").Visible = False
                    oEditTextColumn = aGrid.Columns.Item("U_OWNERCODE")
                    oEditTextColumn.LinkedObjectType = "2"
                Catch ex As Exception
                End Try
                'AddFields("Z_OBILL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
                'AddFields("Z_OBILL", "ComPer", "Commission Percentage", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
                'AddFields("Z_OBILL", "Commission", "Commission Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
                'AddFields("Z_OBILL", "CommGL", "Reciable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
                'AddFields("Z_OBILL", "OwnerCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

                'aGrid.Columns.Item(18).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox



                oEditTextColumn = aGrid.Columns.Item("U_ANNUALRENT")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = aGrid.Columns.Item("U_MONTHRENT")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                oEditTextColumn = aGrid.Columns.Item("U_EXPENSES")
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

                aGrid = aForm.Items.Item("12").Specific
                ' aGrid = aForm.Items.Item("11").Specific
                aGrid.Columns.Item(0).TitleObject.Caption = "Code"
                aGrid.Columns.Item(0).Visible = False
                aGrid.Columns.Item(1).TitleObject.Caption = "Name"
                aGrid.Columns.Item(1).Visible = False
                aGrid.Columns.Item(2).TitleObject.Caption = "RefNo"
                aGrid.Columns.Item(2).Visible = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Expense Code"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "Expense Name"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Account Code"
                aGrid.Columns.Item(5).Editable = False
                oEditTextColumn = aGrid.Columns.Item(5)
                oEditTextColumn.LinkedObjectType = "1"
                aGrid.Columns.Item(6).TitleObject.Caption = "Exp.Type"
                aGrid.Columns.Item(6).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oCombo = aGrid.Columns.Item(6)
                oCombo.ValidValues.Add("S", "Per Sqr.Mtr")
                oCombo.ValidValues.Add("F", "Fixed")
                oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "Rate"
                aGrid.Columns.Item(7).Editable = False
                aGrid.Columns.Item(8).TitleObject.Caption = "Expense Amount"
                oEditTextColumn = aGrid.Columns.Item(8)
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).Visible = False
                aGrid.Columns.Item(10).Visible = False
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        End Select

        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region

#Region "Events"

#Region "Display Expenses"
    Private Sub DisplayExpenses(ByVal aForm As SAPbouiCOM.Form)
        Dim strCode As String = ""
        Dim aGrid As SAPbouiCOM.Grid
        oGrid = aForm.Items.Item("11").Specific
        Try
            aForm.Freeze(True)

            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                If oGrid.Rows.IsSelected(intRow) Then
                    strCode = oGrid.DataTable.GetValue("Code", intRow)
                    Exit For
                End If
            Next
            If strCode <> "" Then
                aGrid = aForm.Items.Item("12").Specific
                aGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where U_Z_RefNo='" & strCode & "'")
                aGrid.Columns.Item(0).TitleObject.Caption = "Code"
                aGrid.Columns.Item(0).Visible = False
                aGrid.Columns.Item(1).TitleObject.Caption = "Name"
                aGrid.Columns.Item(1).Visible = False
                aGrid.Columns.Item(2).TitleObject.Caption = "RefNo"
                aGrid.Columns.Item(2).Visible = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Expense Code"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "Expense Name"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Reciable Account Code"
                aGrid.Columns.Item(5).Editable = False
                oEditTextColumn = aGrid.Columns.Item(5)
                oEditTextColumn.LinkedObjectType = "1"
                aGrid.Columns.Item(6).TitleObject.Caption = "Exp.Type"
                aGrid.Columns.Item(6).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
                oCombo = aGrid.Columns.Item(6)
                oCombo.ValidValues.Add("S", "Per Sqr.Mtr")
                oCombo.ValidValues.Add("F", "Fixed")
                oCombo.ValidValues.Add("P", "Percentage")
                oCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "Rate"
                aGrid.Columns.Item(7).Editable = False
                aGrid.Columns.Item(8).TitleObject.Caption = "Expense Amount"
                oEditTextColumn = aGrid.Columns.Item(8)
                oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).Visible = False
                aGrid.Columns.Item(10).Visible = False
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

            End If
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_BillGeneration
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS, mnu_ADD, mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If

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
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD) Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)

            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_BillGeneration Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "4"
                                        If oForm.PaneLevel = 2 Then
                                            If GenerateBilling(oForm) = False Then
                                                Committrans("Cancel")
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        ElseIf oForm.PaneLevel = 3 Then
                                            If GenerateBilling(oForm) = False Then
                                                Committrans("Cancel")
                                                BubbleEvent = False
                                                Exit Sub
                                            End If

                                        End If
                                        'oForm.PaneLevel = oForm.PaneLevel + 1
                                    Case "2"
                                        If oApplication.SBO_Application.MessageBox("Do you want to Cancel?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        Committrans("Cancel")
                                End Select
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" And pVal.ColUID = "U_Z_AMOUNT" And pVal.CharPressed = 9 Then
                                    oGrid = oForm.Items.Item("15").Specific
                                    Dim dblamount, dblTotalSq, dblrate As Double
                                    dblamount = oGrid.DataTable.GetValue("U_Z_AMOUNT", pVal.Row)
                                    dblTotalSq = oGrid.DataTable.GetValue("U_Z_TOTALSQ", pVal.Row)
                                    If dblamount <= 0 Or dblTotalSq <= 0 Then
                                        dblrate = 0
                                    Else
                                        dblrate = dblamount / dblTotalSq
                                    End If
                                    oGrid.DataTable.SetValue("U_Z_RATE", pVal.Row, dblrate)
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                    Case "5"
                                        If oApplication.SBO_Application.MessageBox("Do you want to calcualte the Billing ?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        If GenerateInvoice(oForm) = True Then
                                            oApplication.Utilities.Message("Operation Completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            Dim x As System.Diagnostics.ProcessStartInfo
                                            x = New System.Diagnostics.ProcessStartInfo
                                            x.UseShellExecute = True
                                            '   sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                            Dim oTemp As SAPbobsCOM.Recordset
                                            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oTemp.DoQuery("Select * from [@Z_LogDetail]")
                                            If oTemp.RecordCount > 0 Then
                                                If oTemp.Fields.Item("U_Z_LOG_PATH").Value <> "" Then
                                                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value
                                                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\ImportLog_Invoice.txt"
                                                    sFailureLog = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\FailureLog.txt"
                                                Else
                                                    sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                                    sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                                                End If
                                            Else
                                                sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                                sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                                            End If
                                            x.FileName = sPath
                                            System.Diagnostics.Process.Start(x)
                                            x = Nothing
                                            Exit Sub
                                        Else
                                            oApplication.SBO_Application.MessageBox("Operation compleated with errors")
                                            Dim x As System.Diagnostics.ProcessStartInfo
                                            x = New System.Diagnostics.ProcessStartInfo
                                            x.UseShellExecute = True
                                            '  sPath = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                                            Dim oTemp As SAPbobsCOM.Recordset
                                            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            oTemp.DoQuery("Select * from [@Z_LogDetail]")
                                            If oTemp.RecordCount > 0 Then
                                                If oTemp.Fields.Item("U_Z_LOG_PATH").Value <> "" Then
                                                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value
                                                    sPath = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\ImportLog_Invoice.txt"
                                                    sFailureLog = oTemp.Fields.Item("U_Z_LOG_PATH").Value & "\FailureLog.txt"
                                                Else
                                                    sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                                    sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                                                End If
                                            Else
                                                sPath = System.Windows.Forms.Application.StartupPath & "\ImportLog_Invoice.txt"
                                                sFailureLog = System.Windows.Forms.Application.StartupPath & "\FailureLog.txt"
                                            End If
                                            x.FileName = sFailureLog ' sPath
                                            System.Diagnostics.Process.Start(x)
                                            x = Nothing
                                            Exit Sub
                                        End If
                                    Case "11"
                                        If pVal.ColUID = "RowsHeader" And pVal.Row >= 0 Then
                                            DisplayExpenses(oForm)
                                        End If
                                    Case "14"
                                        oForm.Close()

                                End Select
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
                                        oForm.Freeze(True)
                                        If pVal.ItemUID = "20" Or pVal.ItemUID = "24" Then
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, oDataTable.GetValue("U_Z_CODE", 0))
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        If pVal.ItemUID = "22" Or pVal.ItemUID = "26" Then
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, oDataTable.GetValue("U_Z_PROITEMCODE", 0))
                                            Catch ex As Exception

                                            End Try

                                        End If

                                        oForm.Freeze(False)
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
