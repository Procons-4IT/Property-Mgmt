Public Class clsPostingWizard
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
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
    Private InvBaseDocNo, sPath As String
    Private InvForConsumedItems, intpane As Integer
    Private blnFlag As Boolean = False
#Region "Methods"
    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_Postingwizard, frm_PostingWizard)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.PaneLevel = 0
        DataBind(oForm)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
    Private Sub DataBind(ByVal aForm As SAPbouiCOM.Form)

        aForm.DataSources.UserDataSources.Add("intYear1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aForm.DataSources.UserDataSources.Add("intMonth1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)

        aForm.DataSources.UserDataSources.Add("intYear", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        aForm.DataSources.UserDataSources.Add("intMonth", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
        oEditText = aForm.Items.Item("7").Specific
        oEditText.DataBind.SetBound(True, "", "intYear1")
        oEditText.ChooseFromListUID = "CFL1"
        oEditText.ChooseFromListAlias = "U_Z_Conno"
        oEditText = aForm.Items.Item("9").Specific
        oEditText.DataBind.SetBound(True, "", "intMonth1")
        oEditText.ChooseFromListUID = "CFL2"
        oEditText.ChooseFromListAlias = "U_Z_Conno"

        oCombobox = aForm.Items.Item("16").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 2010 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombobox.DataBind.SetBound(True, "", "intYear")
        aForm.Items.Item("16").DisplayDesc = True

        oCombobox = aForm.Items.Item("15").Specific
        oCombobox.ValidValues.Add("0", "")
        For intRow As Integer = 1 To 12
            oCombobox.ValidValues.Add(intRow, MonthName(intRow))
        Next
        oCombobox.DataBind.SetBound(True, "", "intMonth")
        oCombobox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        aForm.Items.Item("15").DisplayDesc = True

        oGrid = aForm.Items.Item("10").Specific
        dtTemp = oGrid.DataTable
        Dim strsql, strsql1 As String

        If intpane = 0 Then
            strsql = "SELECT T0.[DocEntry], T0.[U_Z_CONNO], T0.[U_Z_CONTDATE],T0.U_Z_UnitCode,T0.[U_Z_TENCODE], T0.[U_Z_TENNAME], T0.[U_Z_STARTDATE], T0.[U_Z_ENDDATE], T0.[U_Z_STATUS], T0.[U_Z_ANNUALRENT], T0.[U_Z_DEPOSIT], T0.[U_Z_OWNERCODE] FROM [dbo].[@Z_CONTRACT]  T0 where 1=2"
            'dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code  where EmpId=10000000")
        Else
            strsql = "SELECT T0.[DocEntry], T0.[U_Z_CONNO], T0.[U_Z_CONTDATE],T0.U_Z_UnitCode,T0.[U_Z_TENCODE], T0.[U_Z_TENNAME], T0.[U_Z_STARTDATE], T0.[U_Z_ENDDATE], T0.[U_Z_STATUS], T0.[U_Z_ANNUALRENT], T0.[U_Z_DEPOSIT], T0.[U_Z_OWNERCODE] FROM [dbo].[@Z_CONTRACT]  T0 where 1=2"
            '            dtTemp.ExecuteQuery("SELECT T0.[empID], T0.[firstName]+T0.[LastName] 'Emplopyee name', T0.[jobTitle],T1.[Name], T0.[salary], T0.[salaryUnit], T2.[PrcName] FROM OHEM T0  INNER JOIN OUDP T1 ON T0.dept = T1.Code ")
        End If
        dtTemp.ExecuteQuery(strsql)
        oGrid.DataTable = dtTemp
        oGrid = aForm.Items.Item("10").Specific
        Formatgrid(oForm, "Load")
    End Sub

    Private Function PopulateContractDetails(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim strcondition, strFrom, strTo As String
        Dim dtFrom, dtTo As Date
        Dim intMonth, intYear As Integer
        strFrom = oApplication.Utilities.getEdittextvalue(aForm, "19")
        strTo = oApplication.Utilities.getEdittextvalue(aForm, "29")
        If strFrom <> "" Then
            strcondition = " T0.DocEntry >=" & CInt(strFrom)
            'dtFrom = oApplication.Utilities.GetDateTimeValue(strFrom)
        Else
            strcondition = " 1= 1"
            ' oApplication.Utilities.Message("Contract expirty from date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ' Return False
        End If
        If strTo <> "" Then
            strcondition = strcondition & " and T0.DocEntry <=" & CInt(strTo)
            ' dtTo = oApplication.Utilities.GetDateTimeValue(strTo)
        Else
            strcondition = strcondition & " and 1=1"
            ' oApplication.Utilities.Message("Contract Expirty End date is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ' Return False
        End If
        oCombobox = aForm.Items.Item("15").Specific
        Try
            If oCombobox.Selected.Description = "" Then
                oApplication.Utilities.Message("Select Contract Expirty Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else

                intMonth = oCombobox.Selected.Value
                strcondition = strcondition & " and ( Month(T0.U_Z_EndDate)=" & intMonth
            End If
        Catch ex As Exception
            oApplication.Utilities.Message("Select Contract Expirty Month", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try
        oCombobox = aForm.Items.Item("16").Specific
        Try

        
            If oCombobox.Selected.Description = "" Then
                oApplication.Utilities.Message("Select Contract Expirty Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            Else
                intYear = oCombobox.Selected.Value
                strcondition = strcondition & " and  Year(T0.U_Z_EndDate)=" & intYear & ")"

            End If
        Catch ex As Exception
            oApplication.Utilities.Message("Select Contract Expirty Year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End Try

        'If dtFrom > dtTo Then
        '    oApplication.Utilities.Message("Contract Expirty start date should be less than End date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '    Return False
        'End If
        Dim strSql As String
        'strcondition = " T0.U_Z_EndDate between '" & dtFrom.ToString("yyyy-MM-dd") & "' and '" & dtTo.ToString("yyyy-MM-dd") & "'"
        strSql = "SELECT T0.[DocEntry], T0.[U_Z_CONNO], T0.[U_Z_CONTDATE],T0.U_Z_UnitCode,T0.[U_Z_TENCODE], T0.[U_Z_TENNAME], T0.[U_Z_STARTDATE], T0.[U_Z_ENDDATE], case T0.[U_Z_STATUS] When 'PED' then 'Pending for approval' when 'APP' then 'Approved' when 'AGR' then 'Agreed' when 'TER' then 'Terminated' else 'Cancelled' end, T0.[U_Z_ANNUALRENT], T0.[U_Z_DEPOSIT], T0.[U_Z_OWNERCODE] , T0.U_Z_DPNUMBER 'DPNumber' ,' ' 'Select' FROM [dbo].[@Z_CONTRACT]  T0"
        strSql = strSql & " where " & strcondition
        oGrid = aForm.Items.Item("10").Specific
        oGrid.DataTable.ExecuteQuery(strSql)
        Formatgrid(aForm, "Payroll")
        Return True

    End Function
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
        Dim strECode, strESocial, strEname, strETax, strGLAcc, strStartDate, strEndDate As String
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
                    If oApplication.SBO_Application.MessageBox("Bill Generation already completed for this Selected year and month. Do you want to continue?", , "Yes", "No") = 2 Then
                        aForm.Freeze(False)
                        Return False
                    End If
                End If
                oTemp1.DoQuery("Select * from [@Z_OBILL2]  where U_Month=" & intMonth & " and U_Year=" & intYear)
                If oTemp1.RecordCount <= 0 Then
                    'strSQL = "SELECT isnull(sum( T1.[U_Z_SPACE]),0) FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    'strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where T0.U_Z_STATUS='AGR' and (" & intMonth & "  between Month(T0.U_Z_Startdate) and month(T0.U_Z_EndDate))"
                    'strSQL = strSQL & " and (" & intYear & "  between Year(T0.U_Z_Startdate) and Year(T0.U_Z_EndDate))"


                    strSQL = "SELECT isnull(sum( T1.[U_Z_SPACE]),0) FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where T0.U_Z_STATUS='AGR' and ('" & strStartDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"
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
                    If oApplication.SBO_Application.MessageBox("Bill Generation already completed for this Selected year and month. Do you want to continue?", , "Yes", "No") = 2 Then
                        aForm.Freeze(False)
                        Return False
                    End If
                Else
                    oTemp1.DoQuery("Delete from [@Z_OBILL] where U_Month=" & intMonth & " and U_Year=" & intYear)
                    oTemp1.DoQuery("Delete from [@Z_OBILL1] where U_Month=" & intMonth & " and U_Year=" & intYear)
                End If
            End If
            oTemp.DoQuery("Select * from [@Z_OBILL] where U_Month=" & intMonth & " and U_Year=" & intYear)
            If aForm.PaneLevel = 3 Then
                If oTemp.RecordCount > 0 Then
                    oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
                    oExpGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL2]  where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    Formatgrid(aForm, "Payroll")
                    aForm.Items.Item("5").Enabled = False
                Else
                    oTemp1.DoQuery("Delete from [@Z_OBILL] where U_Month=" & intMonth & " and U_Year=" & intYear)
                    oTemp1.DoQuery("Delete from [@Z_OBILL1] where U_Month=" & intMonth & " and U_Year=" & intYear)
                    aForm.Items.Item("5").Enabled = True
                    'strSQL = "SELECT T0.[DocEntry], T0.[U_Z_UNITCODE], T1.[U_Z_SPACE], T0.[U_Z_TENCODE], T0.[U_Z_ANNUALRENT], T0.[U_Z_PAYTRMS], T0.[U_Z_CHGMONTH],T0.[U_Z_ANNUALRENT]/T0.[U_Z_CHGMONTH],T0.U_Z_ACCTCODE ,'N', 0.000,0.000 FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    'strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where T0.U_Z_STATUS='AGR' and (" & intMonth & "  between Month(T0.U_Z_Startdate) and month(T0.U_Z_EndDate))"
                    'strSQL = strSQL & " and (" & intYear & "  between Year(T0.U_Z_Startdate) and Year(T0.U_Z_EndDate))"

                    strSQL = "SELECT T0.[DocEntry], T0.[U_Z_UNITCODE], T1.[U_Z_SPACE], T0.[U_Z_TENCODE], T0.[U_Z_ANNUALRENT], T0.[U_Z_PAYTRMS], T0.[U_Z_CHGMONTH],T0.[U_Z_ANNUALRENT]/T0.[U_Z_CHGMONTH],CASE T0.U_Z_TYPE  when 'T' then T0.U_Z_ACCTCODE1 else T0.U_Z_LiaAc  end  ,'N', 0.000,0.000 ,T0.[U_Z_ConNo] 'ContrctID' ,T0.U_Z_ProType 'PropertyType',T0.U_Z_CommAc 'CommissionAccount',T0.U_Z_Comm 'ComPer', T0.U_Z_OwnerCode 'Owner' FROM [dbo].[@Z_CONTRACT]  T0  inner Join  [dbo].[@Z_PROPUNIT]  T1 "
                    ' strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE] where T0.U_Z_STATUS='AGR' and (" & intMonth & "  between Month(T0.U_Z_Startdate) and month(T0.U_Z_EndDate))"
                    ' strSQL = strSQL & " and (" & intYear & "  between Year(T0.U_Z_Startdate) and Year(T0.U_Z_EndDate))"

                    strSQL = strSQL & " on T1.[U_Z_PROITEMCODE]=T0.[U_Z_UNITCODE]  inner Join OCRD T2 on T2.Cardcode=T0.U_Z_TENCODE where T0.U_Z_STATUS='AGR' and ('" & strStartDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"
                    strSQL = strSQL & " or ('" & strEndDate & "'  between (T0.U_Z_Startdate) and (T0.U_Z_EndDate))"

                    oTemp1.DoQuery(strSQL)
                    For intRow1 As Integer = 0 To oTemp1.RecordCount - 1
                        oUserTable = oApplication.Company.UserTables.Item("Z_OBILL")
                        strCode = oApplication.Utilities.getMaxCode("@Z_OBILL", "Code")
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
                        oUserTable.UserFields.Fields.Item("U_Month").Value = intMonth
                        oUserTable.UserFields.Fields.Item("U_Year").Value = intYear
                        oUserTable.UserFields.Fields.Item("U_ContractNumber").Value = oTemp1.Fields.Item("ContrctID").Value
                        oUserTable.UserFields.Fields.Item("U_ContractID").Value = oTemp1.Fields.Item(0).Value
                        oUserTable.UserFields.Fields.Item("U_UnitCode").Value = oTemp1.Fields.Item(1).Value
                        oUserTable.UserFields.Fields.Item("U_Space").Value = oTemp1.Fields.Item(2).Value
                        oUserTable.UserFields.Fields.Item("U_Annualrent").Value = oTemp1.Fields.Item(4).Value
                        oUserTable.UserFields.Fields.Item("U_PayTrms").Value = oTemp1.Fields.Item(5).Value
                        oUserTable.UserFields.Fields.Item("U_ChgMonth").Value = oTemp1.Fields.Item(6).Value
                        Dim dblMonthRent, dblCommPer, dblCommissionAmount As Double
                        dblMonthRent = oTemp1.Fields.Item(7).Value
                        dblCommPer = oTemp1.Fields.Item("ComPer").Value
                        dblCommissionAmount = dblMonthRent * dblCommPer / 100

                        'T0.Z_ProType 'PropertyType',T0.U_Z_CommAc 'CommissionAccount',T0.U_Z_ConNo 'ComPer', T0.U_Z_OwnerCode 'Owner' 
                        If oTemp1.Fields.Item("PropertyType").Value = "A" Then
                            oUserTable.UserFields.Fields.Item("U_CardCode").Value = oTemp1.Fields.Item(3).Value
                            oUserTable.UserFields.Fields.Item("U_MonthRent").Value = oTemp1.Fields.Item(7).Value
                            oUserTable.UserFields.Fields.Item("U_RentGL").Value = oTemp1.Fields.Item(8).Value
                            oUserTable.UserFields.Fields.Item("U_Remarks").Value = "Monthly Rental  :"
                            oUserTable.UserFields.Fields.Item("U_Commission").Value = dblCommissionAmount
                            oUserTable.UserFields.Fields.Item("U_ComPer").Value = dblCommPer
                            oUserTable.UserFields.Fields.Item("U_CommGL").Value = oTemp1.Fields.Item("CommissionAccount").Value
                            oUserTable.UserFields.Fields.Item("U_OwnerCode").Value = oTemp1.Fields.Item("Owner").Value
                            oUserTable.UserFields.Fields.Item("U_Z_ProType").Value = "A"
                        Else
                            oUserTable.UserFields.Fields.Item("U_CardCode").Value = oTemp1.Fields.Item(3).Value
                            oUserTable.UserFields.Fields.Item("U_MonthRent").Value = oTemp1.Fields.Item(7).Value
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
                    If blnLineExists = False Then
                        oApplication.Utilities.Message("No Contract available for selected month and year", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        aForm.Freeze(False)
                        Return False
                    End If
                    oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
                    oExpGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL2]  where  U_Month=" & intMonth & " and U_Year=" & intYear)
                    Formatgrid(aForm, "Payroll")

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

                oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                oDoc.DocDate = Now.Date
                otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                oDoc.Project = otemp4.Fields.Item(0).Value
                strProject = otemp4.Fields.Item(0).Value
                strCostCenter = otemp4.Fields.Item(1).Value

                If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    'oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If

                If oTemp.Fields.Item("U_Z_ProType").Value = "A" Then
                    oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                    oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                    oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                Else
                    'oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_CommGL").Value)
                    'oDoc.Lines.LineTotal = oTemp.Fields.Item("U_Commission").Value

                    oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                    oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                End If

                oDoc.Lines.ItemDescription = oTemp.Fields.Item("U_Remarks").Value & ":" & oTemp.Fields.Item("U_UNITCODE").Value
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

                    otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus='C' and CardCode='" & strCardCode & "'  and [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                    For intLoop As Integer = 0 To otemp2.RecordCount - 1
                        oDPI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                        If oDPI.GetByKey(otemp2.Fields.Item("DocEntry").Value) Then
                            If DblRental > 0 Then
                                If oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Close Or oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Paid Then
                                    dblDownPayment = oDPI.DownPaymentAmount
                                    oDPRec.DoQuery("Select isnull(DpmAppl,0) from ODPI where DocEntry=" & oDPI.DocEntry)
                                    dblDownPayment = dblDownPayment - oDPRec.Fields.Item(0).Value
                                    If dblDownPayment >= DblRental Then
                                        dblDownPayment = DblRental
                                        DblRental = DblRental - dblDownPayment
                                    Else
                                        dblDownPayment = dblDownPayment
                                        DblRental = DblRental - dblDownPayment
                                    End If
                                    If dblDownPayment > 0 Then
                                        oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                                        oDoc.DownPaymentsToDraw.Add()
                                        oDoc.DownPaymentsToDraw.SetCurrentLine(intCount)
                                        oDoc.DownPaymentsToDraw.DocEntry = oDPI.DocEntry
                                        oDoc.DownPaymentsToDraw.AmountToDraw = dblDownPayment ' oDPI.DownPaymentAmount
                                        intCount = intCount + 1
                                    End If
                                End If
                            End If
                        End If
                        otemp2.MoveNext()
                    Next
                End If

                oTemp1.DoQuery("Select * from [@Z_OBILL1] where U_Z_RefNo='" & oTemp.Fields.Item("Code").Value & "'")
                For intLoop As Integer = 0 To oTemp1.RecordCount - 1
                    oDoc.Lines.Add()
                    oDoc.Lines.SetCurrentLine(intLoop + 1)
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
                        End If
                    End If
                    oTemp1.MoveNext()
                Next

                If oDoc.Add <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    If oApplication.Company.InTransaction() Then
                        oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                    End If
                    aForm.Freeze(False)
                    Return False
                Else
                    Dim strDocNum As String
                    oApplication.Company.GetNewObjectCode(strDocNum)
                    otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                    otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")

                    If 1 = 2 Then ' oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                        oDoc = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)
                        oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                        If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                            strCardCode = oTemp.Fields.Item("U_CARDCODE").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                        Else
                            strCardCode = oTemp.Fields.Item("U_OwnerCode").Value
                            oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                        End If
                        oDoc.CardCode = strCardCode ' oTemp.Fields.Item("U_CARDCODE").Value
                        oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = oTemp.Fields.Item("U_ContractNumber").Value
                        oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = oTemp.Fields.Item("U_ContractID").Value
                        oDoc.DocDate = Now.Date
                        otemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & oTemp.Fields.Item("U_UNITCODE").Value & "'")
                        oDoc.Project = otemp4.Fields.Item(0).Value
                        strProject = otemp4.Fields.Item(0).Value
                        strCostCenter = otemp4.Fields.Item(1).Value
                        If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_RENTGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_MonthRent").Value
                        Else
                            oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(oTemp.Fields.Item("U_CommGL").Value)
                            oDoc.Lines.LineTotal = oTemp.Fields.Item("U_Commission").Value
                        End If
                        oDoc.Lines.ItemDescription = "Monthly Rental :" & oTemp.Fields.Item("U_UNITCODE").Value
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
                        If oTemp.Fields.Item("U_Z_ProType").Value = "T" Then
                            otemp2.DoQuery("Select DocEntry,isnull(U_Z_ContID,0) from ODPI where docstatus='C' and CardCode='" & strCardCode & "' and  [U_Z_DPType]='A' and  isnull(U_Z_ContID,0)=" & aCode)
                            For intLoop As Integer = 0 To otemp2.RecordCount - 1
                                oDPI = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments)
                                If oDPI.GetByKey(otemp2.Fields.Item("DocEntry").Value) Then
                                    If DblRental > 0 Then
                                        If oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Close Or oDPI.DocumentStatus = SAPbobsCOM.BoStatus.bost_Paid Then
                                            dblDownPayment = oDPI.DownPaymentAmount
                                            oDPRec.DoQuery("Select isnull(DpmAppl,0) from ODPI where DocEntry=" & oDPI.DocEntry)
                                            dblDownPayment = dblDownPayment - oDPRec.Fields.Item(0).Value
                                            If dblDownPayment >= DblRental Then
                                                dblDownPayment = DblRental
                                                DblRental = DblRental - dblDownPayment
                                            Else
                                                dblDownPayment = dblDownPayment
                                                DblRental = DblRental - dblDownPayment
                                            End If
                                            oDoc.DownPaymentType = SAPbobsCOM.DownPaymentTypeEnum.dptInvoice
                                            oDoc.DownPaymentsToDraw.Add()
                                            oDoc.DownPaymentsToDraw.SetCurrentLine(intCount)
                                            oDoc.DownPaymentsToDraw.DocEntry = oDPI.DocEntry
                                            oDoc.DownPaymentsToDraw.AmountToDraw = dblDownPayment ' oDPI.DownPaymentAmount
                                            intCount = intCount + 1
                                        End If
                                    End If
                                End If
                                otemp2.MoveNext()
                            Next
                        End If
                        If oDoc.Add <> 0 Then
                            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            If oApplication.Company.InTransaction() Then
                                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                            End If
                            aForm.Freeze(False)
                            Return False
                        Else
                            oApplication.Company.GetNewObjectCode(strDocNum)
                            otemp2.DoQuery("Select * from OINV where Docentry=" & strDocNum)
                            otemp3.DoQuery("Update [@Z_OBILL] set U_Invoiced='Y' , U_InvEntry=" & strDocNum & ",U_InvNumber=" & otemp2.Fields.Item("DocNum").Value & " where Code='" & oTemp.Fields.Item("Code").Value & "'")
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
            oHeaderGrid.DataTable.ExecuteQuery("Select * from [@Z_OBILL] where  U_Month=" & intMonth & " and U_Year=" & intYear)
            oLineGird.DataTable.ExecuteQuery("Select * from [@Z_OBILL1] where Code='xxxx'")
            Formatgrid(aForm, "Payroll")
            aForm.Items.Item("5").Enabled = False
            aForm.Freeze(False)
            oApplication.Utilities.Message("Operation completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Return True
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            If oApplication.Company.InTransaction() Then
                oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            End If
            aForm.Freeze(False)
            Return False
        End Try
    End Function
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
        Databind(aform)
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


#Region "Create Documents"

    Private Function createDocuments(ByVal aForm As SAPbouiCOM.Form, ByVal aChoice As String) As Boolean
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        oApplication.Company.StartTransaction()
        oGrid = aForm.Items.Item("10").Specific
        Dim oCheckBox As SAPbouiCOM.CheckBoxColumn
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            oCheckBox = oGrid.Columns.Item("Select")
            If oCheckBox.IsChecked(intRow) Then
                If aChoice = "Both" Then
                    If CreateDownPaymentInvoice(oGrid.DataTable.GetValue(0, intRow)) = False Then
                        If oApplication.Company.InTransaction() Then
                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    End If
                    If CreateSecurityDepost(oGrid.DataTable.GetValue(0, intRow)) = False Then
                        If oApplication.Company.InTransaction() Then
                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    End If
                ElseIf aChoice = "DP" Then
                    If CreateDownPaymentInvoice(oGrid.DataTable.GetValue(0, intRow)) = False Then
                        If oApplication.Company.InTransaction() Then
                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    End If
                Else
                    If CreateSecurityDepost(oGrid.DataTable.GetValue(0, intRow)) = False Then
                        If oApplication.Company.InTransaction() Then
                            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
                        End If
                        Return False
                    End If

                End If

            End If
        Next
        If oApplication.Company.InTransaction Then
            oApplication.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
        End If
        Return True
    End Function

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
        otemp.DoQuery("select isnull(U_Z_Annualrent,0), * from [@Z_CONTRACT] where DocEntry=" & acode)
        dbldownpayment = otemp.Fields.Item(0).Value
        If blnRecordExits = True Then
            dbldownpayment = dbldownpayment - dbltotalDownpayment
        End If
        If dbldownpayment > 0 Then
            otemp.DoQuery("select isnull(U_Z_Annualrent,0) , isnull(U_Z_Type,'O') 'Type', * from [@Z_CONTRACT] where DocEntry=" & acode)
            If otemp.Fields.Item(0).Value > 0 Then
                strcustCode = otemp.Fields.Item("Type").Value
                Dim strCreditAccountcode As String
                If otemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    '  oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                Else
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If

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

                If otemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                Else
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = acode
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = otemp.Fields.Item("U_Z_ConNo").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = otemp.Fields.Item("U_Z_ConNo").Value
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
                oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAccountcode)
                oDoc.Lines.ItemDescription = "Advance Annual Rent for UntiCode : " & otemp.Fields.Item("U_Z_UnitCode").Value
                'oDoc.Lines.TaxCode = "Exempt"
                If strCostCenter <> "" Then
                    oDoc.Lines.CostingCode = strCostCenter
                End If
                If strProject <> "" Then
                    oDoc.Lines.ProjectCode = strProject
                End If
                oDoc.Lines.LineTotal = dblPrice ' otemp.Fields.Item("U_MonthRent").Value
                If oBP.GetByKey(strOwner) Then
                    If oBP.VatGroup <> "" Then
                        oDoc.Lines.TaxCode = oBP.VatGroup
                    End If
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
        Return True
    End Function

    Private Function CreateSecurityDepost(ByVal acode As Integer) As Boolean
        Dim otemp, otemp1 As SAPbobsCOM.Recordset
        Dim strcustCode, stritemcode, strqty, strprice, strtotal, strdriver, strcardname As String
        Dim oDoc As SAPbobsCOM.Documents
        Dim oBP As SAPbobsCOM.BusinessPartners
        Dim dblQty, dblPrice As Double
        Dim blnRecordExits As Boolean = False
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp1 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oBP = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)

        otemp1.DoQuery("SELECT sum(DocTotal)-sum(VatSum)  FROM ODPI T0 WHERE T0.[U_Z_DPType]='S' and T0.[U_Z_ContID] =" & acode)
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
        otemp.DoQuery("select isnull(U_Z_Deposit,0), * from [@Z_CONTRACT] where DocEntry=" & acode)
        dbldownpayment = otemp.Fields.Item(0).Value
        If blnRecordExits = True Then
            dbldownpayment = dbldownpayment - dbltotalDownpayment
        End If
        If dbldownpayment > 0 Then
            ' otemp.DoQuery("select isnull(U_Z_Deposit,0) + isnull(U_Z_Salik,0)+isnull(U_Z_DPAmount,0), * from [@Z_ORDR] where DocEntry=" & acode)
            otemp.DoQuery("select isnull(U_Z_Deposit,0), isnull(U_Z_Type,'O') 'Type', * from [@Z_CONTRACT] where DocEntry=" & acode)
            If otemp.Fields.Item(0).Value > 0 Then
                strcustCode = otemp.Fields.Item("Type").Value
                Dim strCreditAccountcode As String
                If otemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    'oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                Else
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If
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

                If otemp.Fields.Item("U_Z_ProType").Value = "T" Then
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "O"
                Else
                    strCreditAccountcode = otemp.Fields.Item("U_Z_LiaAc").Value
                    ' oDoc.UserFields.Fields.Item("U_Z_INVTYPE").Value = "T"
                End If
                oDoc.UserFields.Fields.Item("U_Z_CONTID").Value = acode
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = otemp.Fields.Item("U_Z_ConNo").Value

                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = otemp.Fields.Item("U_Z_ConNo").Value
                oDoc.UserFields.Fields.Item("U_Z_STARTDATE").Value = otemp.Fields.Item("U_Z_StartDate").Value
                oDoc.UserFields.Fields.Item("U_Z_ENDDATE").Value = otemp.Fields.Item("U_Z_EndDate").Value
                oDoc.UserFields.Fields.Item("U_Z_DPTYPE").Value = "S"
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                Dim oTemp4 As SAPbobsCOM.Recordset
                Dim strProject, strCostCenter As String
                oTemp4 = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oTemp4.DoQuery("Select isnull(U_Z_PropCode,''),isnull(U_Z_CostCenter,'') from [@Z_PROPUNIT] where U_Z_ProItemCode='" & otemp.Fields.Item("U_Z_UnitCode").Value & "'")
                oDoc.Project = oTemp4.Fields.Item(0).Value
                strProject = oTemp4.Fields.Item(0).Value
                strCostCenter = oTemp4.Fields.Item(1).Value
                oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAccountcode)
                oDoc.Lines.ItemDescription = "Security Deposit for UntiCode : " & otemp.Fields.Item("U_Z_UnitCode").Value
                If strCostCenter <> "" Then
                    oDoc.Lines.CostingCode = strCostCenter
                End If
                If strProject <> "" Then
                    oDoc.Lines.ProjectCode = strProject
                End If
                oDoc.Lines.LineTotal = dblPrice ' otemp.Fields.Item("U_MonthRent").Value
                If oBP.GetByKey(strOwner) Then
                    oDoc.Lines.TaxCode = oBP.VatGroup
                End If
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
                '  otemp1.DoQuery("Update [@Z_CONTRACT] set U_Z_DPEntry=" & otemp.Fields.Item(0).Value & ",U_Z_DPNumber='" & otemp.Fields.Item(1).Value & "' where docentry=" & acode)
                oApplication.Utilities.Message("Down Payment Created sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                Return True
            End If
        End If
        'End If
        'CreateIncomingPayment(acode)
        Return True
    End Function

#End Region
#Region "FormatGrid"
    Private Sub Formatgrid(ByVal aForm As SAPbouiCOM.Form, ByVal aOption As String)
        Dim aGrid As SAPbouiCOM.Grid

        Select Case aOption
            Case "Load"
                aGrid = aForm.Items.Item("10").Specific
                aGrid.Columns.Item(0).TitleObject.Caption = "Contract ID"
                aGrid.Columns.Item(0).Editable = False
                aGrid.Columns.Item(1).TitleObject.Caption = "Contract Number"
                aGrid.Columns.Item(1).Editable = False
                aGrid.Columns.Item(2).TitleObject.Caption = "Contract Date"
                aGrid.Columns.Item(2).Editable = False
                aGrid.Columns.Item(3).TitleObject.Caption = "UnitCode"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "Tenant Code"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Tenant Name"
                aGrid.Columns.Item(5).Editable = False
                aGrid.Columns.Item(6).TitleObject.Caption = "Start Date"
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "End Date"
                aGrid.Columns.Item(7).Editable = False
                aGrid.Columns.Item(8).TitleObject.Caption = "Status"
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).TitleObject.Caption = "Annual Rent"
                aGrid.Columns.Item(9).Editable = False
                aGrid.Columns.Item(10).TitleObject.Caption = "Security Deposit"
                aGrid.Columns.Item(10).Editable = False
                aGrid.Columns.Item(11).TitleObject.Caption = "Owner Code"
                aGrid.Columns.Item(11).Editable = False
            Case "Payroll"
                aGrid = aForm.Items.Item("10").Specific
                aGrid.Columns.Item(0).TitleObject.Caption = "Contract ID"
                aGrid.Columns.Item(0).Editable = False
                oEditTextColumn = aGrid.Columns.Item(0)
                oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                aGrid.Columns.Item(1).TitleObject.Caption = "Contract Number"
                aGrid.Columns.Item(1).Editable = False
                aGrid.Columns.Item(2).TitleObject.Caption = "Contract Date"
                aGrid.Columns.Item(2).Editable = False
                aGrid.Columns.Item(3).TitleObject.Caption = "UnitCode"
                aGrid.Columns.Item(3).Editable = False
                aGrid.Columns.Item(4).TitleObject.Caption = "Tenant Code"
                aGrid.Columns.Item(4).Editable = False
                aGrid.Columns.Item(5).TitleObject.Caption = "Tenant Name"
                aGrid.Columns.Item(5).Editable = False
                aGrid.Columns.Item(6).TitleObject.Caption = "Start Date"
                aGrid.Columns.Item(6).Editable = False
                aGrid.Columns.Item(7).TitleObject.Caption = "End Date"
                aGrid.Columns.Item(7).Editable = False
                aGrid.Columns.Item(8).TitleObject.Caption = "Status"
                aGrid.Columns.Item(8).Editable = False
                aGrid.Columns.Item(9).TitleObject.Caption = "Annual Rent"
                aGrid.Columns.Item(9).Editable = False
                aGrid.Columns.Item(10).TitleObject.Caption = "Security Deposit"
                aGrid.Columns.Item(10).Editable = False
                aGrid.Columns.Item(11).TitleObject.Caption = "Owner Code"
                aGrid.Columns.Item(11).Editable = False
                aGrid.Columns.Item("DPNumber").TitleObject.Caption = "Down Payment number"
                aGrid.Columns.Item("DPNumber").Editable = False
                aGrid.Columns.Item("Select").TitleObject.Caption = "Select"
                aGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
                aGrid.AutoResizeColumns()
                aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single

        End Select

        agrid.AutoResizeColumns()
        agrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
    End Sub
#End Region
    Private Sub SelectAll(ByVal aform As SAPbouiCOM.Form, ByVal blnValue As Boolean)
        aform.Freeze(True)
        oGrid = aform.Items.Item("10").Specific
        Dim ocheckbox As SAPbouiCOM.CheckBoxColumn
        For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
            ocheckbox = oGrid.Columns.Item("Select")
            If blnValue = True Then
                oGrid.DataTable.SetValue("Select", intRow, "Y")
            Else
                oGrid.DataTable.SetValue("Select", intRow, "N")
            End If
        Next
        aform.Freeze(False)
    End Sub

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
                Case mnu_postingWizard
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
            If pVal.FormTypeEx = frm_PostingWizard Then
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
                                            If PopulateContractDetails(oForm) = False Then
                                                BubbleEvent = False
                                                Exit Sub
                                            End If
                                        End If
                                    Case "2"
                                        If oApplication.SBO_Application.MessageBox("Do you want to Cancel?", , "Yes", "No") = 2 Then
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                        '  Committrans("Cancel")
                                End Select
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "10" And pVal.ColUID = "DocEntry" Then
                                    oGrid = oForm.Items.Item("10").Specific
                                    Dim oobj As New clsTenContracts
                                    oobj.LoadForm_Contract_View(oGrid.DataTable.GetValue("U_Z_CONNO", pVal.Row))
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "3"
                                        oForm.PaneLevel = oForm.PaneLevel - 1
                                    Case "4"
                                        oForm.PaneLevel = oForm.PaneLevel + 1
                                    Case "17"
                                        SelectAll(oForm, True)
                                    Case "18"
                                        SelectAll(oForm, False)
                                    Case "5"
                                        Dim intChoice As String = "Both"
                                        If oApplication.SBO_Application.MessageBox("Do you want to Post Missing Downpayment Invoices?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        If createDocuments(oForm, intChoice) = True Then
                                            oApplication.Utilities.Message("Operation Completed successfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                                            oForm.Close()
                                        End If
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
                                        If pVal.ItemUID = "9" Or pVal.ItemUID = "7" Then
                                            val = oDataTable.GetValue("U_Z_CONNO", 0)
                                            val1 = oDataTable.GetValue("DocEntry", 0)
                                            If pVal.ItemUID = "7" Then
                                                oApplication.Utilities.setEdittextvalue(oForm, "19", val1)
                                            Else
                                                oApplication.Utilities.setEdittextvalue(oForm, "29", val1)
                                            End If
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                            End Try
                                            oForm.Freeze(False)
                                        End If
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
