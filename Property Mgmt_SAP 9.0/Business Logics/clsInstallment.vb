Public Class clsInstallment
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

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub

    Private Sub FillRentType(aForm As SAPbouiCOM.Form)
        oCombo = aForm.Items.Item("17").Specific
        For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
            Try
                oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Catch ex As Exception

            End Try
        Next
        oCombo.ValidValues.Add("", "")
        oCombo.ValidValues.Add("D", "Daily")
        oCombo.ValidValues.Add("W", "Weekly")
        oCombo.ValidValues.Add("M", "Monthly")
        oCombo.ValidValues.Add("Q", "Quarterly")
        oCombo.ValidValues.Add("S", "Semi Annual")
        oCombo.ValidValues.Add("A", "Annual")
        oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        oCombo.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        aForm.Items.Item("17").DisplayDesc = True
    End Sub
    Public Sub LoadForm(ByVal aContractNo As String)
        oForm = oApplication.Utilities.LoadForm(xml_Installments, frm_Installments)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataSources.UserDataSources.Add("dtDate", SAPbouiCOM.BoDataType.dt_DATE)
        oEditText = oForm.Items.Item("9").Specific
        oEditText.DataBind.SetBound(True, "", "dtDate")
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        FillRentType(oForm)
        oTest.DoQuery("Select * from [@Z_CONTRACT] where DocEntry=" & CInt(aContractNo))
        If oTest.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(oForm, "3", oTest.Fields.Item("DocEntry").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "5", oTest.Fields.Item("U_Z_CntNo").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "7", oTest.Fields.Item("U_Z_StartDate").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "9", oTest.Fields.Item("U_Z_EndDate").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "11", oTest.Fields.Item("U_Z_ChgMonth").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "13", oTest.Fields.Item("U_Z_Annualrent").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "19", oTest.Fields.Item("U_Z_Monthly").Value)
            oCombo = oForm.Items.Item("17").Specific
            oCombo.Select(oTest.Fields.Item("U_Z_RentType").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount',U_Z_StartDate1 'From',U_Z_EndDate1 'To',U_Z_Manual 'Manual',U_Z_Status 'Paid Status' from [@Z_CONINS] where U_Z_ConId=" & CInt(aContractNo) & " Order by U_Z_StartDate1")

        Else
            oGrid = oForm.Items.Item("15").Specific
            oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_StartDate1 'From',U_Z_EndDate1 'To',U_Z_Amount 'Amount',U_Z_Manual 'Manual',U_Z_Status 'Paid Status' from [@Z_CONINS] where 1=2 Order by U_Z_StartDate1")
        End If
        Dim strUser As String = oApplication.Company.UserName
        Dim aChoice As Boolean = False
        oTest.DoQuery("Select isnull(U_Z_Install,'N'),* from OUSR where USER_CODE='" & strUser & "'")
        If oTest.Fields.Item(0).Value = "Y" Then
            If AddtoUDT(oForm) = True Then
                oGrid = oForm.Items.Item("15").Specific
                oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount',U_Z_StartDate1 'From',U_Z_EndDate1 'To',U_Z_Manual 'U_Z_Manual',U_Z_Status 'Paid Status' from [@Z_CONINS]  where U_Z_ConId=" & CInt(aContractNo) & " Order by U_Z_StartDate1")
            End If
            oForm.Items.Item("14").Visible = True
            aChoice = True
        Else
            If AddtoUDT(oForm) = True Then
                oGrid = oForm.Items.Item("15").Specific
                oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount',U_Z_StartDate1 'From',U_Z_EndDate1 'To',U_Z_Manual 'U_Z_Manual',U_Z_Status 'Paid Status' from [@Z_CONINS]  where U_Z_ConId=" & CInt(aContractNo) & " Order by U_Z_StartDate1")
            End If
            oForm.Items.Item("14").Visible = True
            aChoice = False
        End If

        Formatgrid(oGrid, aChoice)

        'oForm.DataBrowser.BrowseBy = "1000002"
        'oForm.EnableMenu(mnu_DELETE_ROW, True)
        'oForm.EnableMenu(mnu_ADD_ROW, True)
        'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        'Dim oTest As SAPbobsCOM.Recordset
        'oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'Dim st As String
        'st = "Update [@Z_CONTRACT] set U_Z_CntNo= U_Z_ConNo +'_'+ convert(varchar,isnull(U_Z_SeqNo,'1'))"
        'oTest.DoQuery(st)
        'oForm.Items.Item("18").Visible = False
        'oForm.Items.Item("19").Visible = False
        'oForm.Items.Item("120").Visible = False
        'oForm.Items.Item("60").Visible = False
        'oForm.PaneLevel = 1
        oForm.Freeze(False)
        If blnInstallmentfromContract = True Then
            blnInstallmentfromContract = False
            oForm.Close()
        End If
    End Sub
#End Region

    Private Sub Formatgrid(ByVal aGrid As SAPbouiCOM.Grid, ByVal aflag As Boolean)
        aGrid.Columns.Item(0).Visible = False
        aGrid.Columns.Item(1).Visible = False
        aGrid.Columns.Item(2).TitleObject.Caption = "Month"
        aGrid.Columns.Item(2).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item(2)
        For intRow As Integer = 1 To 12
            oCombobox.ValidValues.Add(intRow, MonthName(intRow))
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly

        aGrid.Columns.Item(3).TitleObject.Caption = "Year"
        aGrid.Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item(3)
        For intRow As Integer = 2005 To 2050
            oCombobox.ValidValues.Add(intRow, intRow)
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        aGrid.Columns.Item(2).Editable = False
        aGrid.Columns.Item(3).Editable = False
        aGrid.Columns.Item(4).Editable = aflag
        oEditTextColumn = aGrid.Columns.Item(4)
        oEditTextColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto
        aGrid.Columns.Item("U_Z_Manual").TitleObject.Caption = "Manual"
        aGrid.Columns.Item("U_Z_Manual").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox
        aGrid.Columns.Item("From").Editable = False
        aGrid.Columns.Item("To").Editable = False
        aGrid.Columns.Item("Paid Status").TitleObject.Caption = "Installment Paid Status"
        aGrid.Columns.Item("Paid Status").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oCombobox = aGrid.Columns.Item("Paid Status")
        oCombobox.ValidValues.Add("Y", "Paid")
        oCombobox.ValidValues.Add("N", "Pending")
        oCombobox.ExpandType = SAPbouiCOM.BoExpandType.et_DescriptionOnly
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both
        aGrid.Columns.Item("Paid Status").Editable = False
        aGrid.RowHeaders.TitleObject.Caption = "#"
        For intRow As Integer = 0 To aGrid.DataTable.Rows.Count - 1
            aGrid.RowHeaders.SetText(intRow, intRow + 1)
        Next

        aGrid.AutoResizeColumns()
        aGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
    End Sub


    Private Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
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
        Dim strRentType1 As String = otest.Fields.Item("U_Z_RENTTYPE").Value
        Dim dtContStart, dtContEnd As Date
        dtContStart = otest.Fields.Item("U_Z_StartDate").Value
        dtContEnd = otest.Fields.Item("U_Z_EndDate").Value
        otest.DoQuery("Select * from [@Z_CONINS] where U_Z_Status='Y' and  U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
        If otest.RecordCount <= 0 Then
            otest.DoQuery("Select * from [@Z_CONINS] where U_Z_ConID=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
            If otest.Fields.Item("U_Z_RENTTYPE").Value <> strRentType1 Then
                otest.DoQuery("Delete from [@Z_CONINS] where   U_Z_ConID=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
            Else
                otest.DoQuery("Delete from [@Z_CONINS] where U_Z_Manual<>'Y' and  U_Z_ConID=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
            End If
        Else
            Return True
        End If

        otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
        Dim dtTo1 As Date
        If 1 = 1 Then ' otest.RecordCount <= 0 Then
            Dim dblAssignedRent As Double = 0
            Dim dblRent As Double
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
                dblRent = dblAnnualRent / dblNoofMonths
                If dt12 >= dtTo Then
                    dtTo1 = dtTo
                    dblRent = dblAnnualRent - dblAssignedRent
                End If


                otest.DoQuery("Select Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount',U_Z_Manual from [@Z_CONINS] where U_Z_Month=" & Month(dtFrom) & " and U_Z_Year=" & Year(dtFrom) & " and   U_Z_ConId=" & CInt(oApplication.Utilities.getEdittextvalue(aForm, "3")))
                If otest.RecordCount <= 0 Then
                    strCode = oApplication.Utilities.getMaxCode("@Z_CONINS", "Code")
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                    oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = dtContStart ' (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = dtContEnd '(oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                    oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                    oUserTable.UserFields.Fields.Item("U_Z_MonthRent").Value = oApplication.Utilities.getEdittextvalue(aForm, "19")

                    oUserTable.UserFields.Fields.Item("U_Z_StartDate1").Value = dtFrom
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate1").Value = dtTo1 '(oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_Month").Value = Month(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Year").Value = Year(dtFrom)
                    oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = dblRent ' dblAnnualRent / dblNoofMonths
                    oUserTable.UserFields.Fields.Item("U_Z_RentType").Value = oCombo.Selected.Value
                    oUserTable.UserFields.Fields.Item("U_Z_Status").Value = "N"
                    oUserTable.UserFields.Fields.Item("U_Z_Manual").Value = "N"
                    dblAssignedRent = dblAssignedRent + dblRent
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
                        oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = dtContStart ' (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                        oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = dtContEnd ' (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                        oUserTable.UserFields.Fields.Item("U_Z_StartDate1").Value = dtFrom
                        oUserTable.UserFields.Fields.Item("U_Z_EndDate1").Value = dtTo1 '(oApplication.Utilities.getEdittextvalue(aForm, "9"))
                        oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                        oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                        oUserTable.UserFields.Fields.Item("U_Z_MonthRent").Value = oApplication.Utilities.getEdittextvalue(aForm, "19")
                        oUserTable.UserFields.Fields.Item("U_Z_Month").Value = Month(dtFrom)
                        oUserTable.UserFields.Fields.Item("U_Z_Year").Value = Year(dtFrom)
                        If oUserTable.UserFields.Fields.Item("U_Z_Manual").Value = "Y" Then
                            dblRent = oUserTable.UserFields.Fields.Item("U_Z_Amount").Value
                            oUserTable.UserFields.Fields.Item("U_Z_Manual").Value = "Y"
                        End If
                        oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = dblRent ' dblAnnualRent / dblNoofMonths
                        oUserTable.UserFields.Fields.Item("U_Z_RentType").Value = oCombo.Selected.Value
                        oUserTable.UserFields.Fields.Item("U_Z_Status").Value = "N"
                        dblAssignedRent = dblAssignedRent + dblRent
                        '  oUserTable.UserFields.Fields.Item("U_Z_Status").Value = "N"
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

    Private Function Validation(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim dtFrom, dtTo As Date
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        Dim dblAnnualRent, dblNoofMonths As Double
        Dim oCheckBox As SAPbouiCOM.CheckBoxColumn
        dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "7"))
        dtTo = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "9"))
        Dim otest As SAPbobsCOM.Recordset
        Dim dblRent As Double = 0
        dblAnnualRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "13"))
        dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "11"))
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If 1 = 1 Then ' otest.RecordCount <= 0 Then
            oGrid = aForm.Items.Item("15").Specific
            ' oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                oCheckBox = oGrid.Columns.Item("U_Z_Manual")
                dblRent = dblRent + oGrid.DataTable.GetValue("Amount", intRow)
            Next
        End If
        If Math.Round(dblRent, 3) <> Math.Round(dblAnnualRent, 3) Then

            oApplication.Utilities.Message("Total Monthly rental not matched with Annual Rent. Difference : " & dblAnnualRent - dblRent, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        Return True

    End Function


    Private Function AddtoUDT_Table(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim dtFrom, dtTo As Date
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim strCode As String
        Dim dblAnnualRent, dblNoofMonths As Double
        Dim oCheckBox As SAPbouiCOM.CheckBoxColumn
        dtFrom = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "7"))
        dtTo = oApplication.Utilities.GetDateTimeValue(oApplication.Utilities.getEdittextvalue(aForm, "9"))
        Dim otest As SAPbobsCOM.Recordset
        Dim dblRent As Double = 0
        oCombo = aForm.Items.Item("17").Specific
        dblAnnualRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "13"))
        dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(aForm, "11"))
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
         If 1 = 1 Then ' otest.RecordCount <= 0 Then
            oUserTable = oApplication.Company.UserTables.Item("Z_CONINS")
            oGrid = aForm.Items.Item("15").Specific
            ' oGrid.DataTable.ExecuteQuery("Select  Code,Name,Cast(U_Z_Month as varchar),cast(U_Z_Year as varchar),U_Z_Amount 'Amount' from [@Z_CONINS] where 1=2")
            For intRow As Integer = 0 To oGrid.DataTable.Rows.Count - 1
                strCode = oGrid.DataTable.GetValue("Code", intRow)
                oCheckBox = oGrid.Columns.Item("U_Z_Manual")
                If oUserTable.GetByKey(strCode) Then
                    oUserTable.Code = strCode
                    oUserTable.Name = strCode
                    oUserTable.UserFields.Fields.Item("U_Z_ConID").Value = CInt(oApplication.Utilities.getEdittextvalue(aForm, "3"))
                    oUserTable.UserFields.Fields.Item("U_Z_ConNo").Value = (oApplication.Utilities.getEdittextvalue(aForm, "5"))
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "7"))
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = (oApplication.Utilities.getEdittextvalue(aForm, "9"))
                    oUserTable.UserFields.Fields.Item("U_Z_NoofMonths").Value = (oApplication.Utilities.getEdittextvalue(aForm, "11"))
                    oUserTable.UserFields.Fields.Item("U_Z_AnnualRent").Value = (oApplication.Utilities.getEdittextvalue(aForm, "13"))
                    oUserTable.UserFields.Fields.Item("U_Z_Month").Value = oGrid.DataTable.GetValue(2, intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Year").Value = oGrid.DataTable.GetValue(3, intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_Amount").Value = oGrid.DataTable.GetValue("Amount", intRow)
                    dblRent = dblRent + oGrid.DataTable.GetValue("Amount", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_StartDate").Value = oGrid.DataTable.GetValue("From", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_EndDate").Value = oGrid.DataTable.GetValue("To", intRow)
                    oUserTable.UserFields.Fields.Item("U_Z_RentType").Value = oCombo.Selected.Value
                    If oCheckBox.IsChecked(intRow) Then
                        oUserTable.UserFields.Fields.Item("U_Z_Manual").Value = "Y"
                    Else
                        oUserTable.UserFields.Fields.Item("U_Z_Manual").Value = "N"
                    End If
                    oUserTable.UserFields.Fields.Item("U_Z_Status").Value = oGrid.DataTable.GetValue("Paid Status", intRow)
                    If oUserTable.Update <> 0 Then
                        oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
            Next
        End If
        Return True

    End Function
#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_ItemGroup
                    If pVal.BeforeAction = False Then
                        '  LoadForm()
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
            If pVal.FormTypeEx = frm_Installments Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" And pVal.ColUID = "Amount" And pVal.CharPressed <> 9 Then
                                    oGrid = oForm.Items.Item("15").Specific
                                    Dim ocheck As SAPbouiCOM.CheckBoxColumn
                                    ocheck = oGrid.Columns.Item("U_Z_Manual")
                                    If ocheck.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" And pVal.ColUID = "Amount" Then
                                    oGrid = oForm.Items.Item("15").Specific
                                    Dim ocheck As SAPbouiCOM.CheckBoxColumn
                                    ocheck = oGrid.Columns.Item("U_Z_Manual")
                                    If ocheck.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If

                            Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "15" And pVal.ColUID = "Amount" Then
                                    oGrid = oForm.Items.Item("15").Specific
                                    Dim ocheck As SAPbouiCOM.CheckBoxColumn
                                    ocheck = oGrid.Columns.Item("U_Z_Manual")
                                    If ocheck.IsChecked(pVal.Row) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                If pVal.ItemUID = "14" Then
                                    If Validation(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "14"
                                        If oApplication.SBO_Application.MessageBox("Do you want to save the details ?", , "Yes", "No") = 2 Then
                                            Exit Sub
                                        End If
                                        If AddtoUDT_Table(oForm) = True Then
                                            oForm.Close()
                                        End If
                                End Select
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
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
