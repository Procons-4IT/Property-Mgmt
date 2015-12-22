Public Class clsTenContracts
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
    Private oColumn As SAPbouiCOM.Column
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
    Private strSelectedFilepath, strSelectedFolderPath As String

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        oForm = oApplication.Utilities.LoadForm(xml_TenContracts, frm_TenContracts)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "1000002"
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        AddChooseFromList(oForm)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim st As String
        st = "Update [@Z_CONTRACT] set U_Z_CntNo= U_Z_ConNo +'_'+ convert(varchar,isnull(U_Z_SeqNo,'1'))"
        oTest.DoQuery(st)
        databind(oForm)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
#End Region


#Region "ShowFileDialog"

    '*****************************************************************
    'Type               : Procedure
    'Name               : ShowFileDialog
    'Parameter          :
    'Return Value       :
    'Author             : Senthil Kumar B 
    'Created Date       : 
    'Last Modified By   : 
    'Modified Date      : 
    'Purpose            : To open a File Browser
    '******************************************************************

    Private Sub fillopen()
        Dim mythr As New System.Threading.Thread(AddressOf ShowFileDialog)
        mythr.SetApartmentState(Threading.ApartmentState.STA)
        mythr.Start()
        mythr.Join()

    End Sub

    Private Sub ShowFileDialog()
        Dim oDialogBox As New OpenFileDialog
        Dim strFileName, strMdbFilePath As String
        Dim oEdit As SAPbouiCOM.EditText
        Dim oProcesses() As Process
        Try
            oProcesses = Process.GetProcessesByName("SAP Business One")
            If oProcesses.Length <> 0 Then
                For i As Integer = 0 To oProcesses.Length - 1
                    Dim MyWindow As New WindowWrapper(oProcesses(i).MainWindowHandle)
                    If oDialogBox.ShowDialog(MyWindow) = DialogResult.OK Then
                        strMdbFilePath = oDialogBox.FileName
                        strSelectedFilepath = oDialogBox.FileName
                        strFileName = strSelectedFilepath
                        strSelectedFolderPath = strFileName
                    Else
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub
#End Region
    Private Function dateDifference(ByVal aDate1 As Date, ByVal adate2 As Date) As Double
        Dim intDifferenance As Integer
        intDifferenance = DateDiff(DateInterval.Day, adate2, aDate1)
        Return intDifferenance
    End Function
    Private Function ValidateUnitExists(ByVal aform As SAPbouiCOM.Form, ByVal dtStartdate As Date, ByVal dtEndDate As Date) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            Dim intDocEntry As Integer
            Dim strUnitCode, strSql As String
            Dim adate As Date
            intDocEntry = CInt(oApplication.Utilities.getEdittextvalue(aform, "1000002"))
            strUnitCode = oApplication.Utilities.getEdittextvalue(aform, "43")
            'strSql = "Select * from [@Z_CONTRACT] where U_Z_UnitCode='" & strUnitCode & "' and DocEntry<>" & intDocEntry & " and '" & dtStartdate.ToString("yyyy-MM-yy") & "' between U_Z_StartDate and U_Z_EndDate"
            strSql = "Select * from [@Z_CONTRACT] where U_Z_UnitCode='" & strUnitCode & "' and  U_Z_STATUS<>'CAN'  and DocEntry<>" & intDocEntry & " order by DocEntry desc"

            oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTemp.DoQuery(strSql)
            If oTemp.RecordCount > 0 Then
                adate = oTemp.Fields.Item("U_Z_EndDate").Value
                Dim intDifferenance As Integer
                intDifferenance = DateDiff(DateInterval.Day, adate, dtStartdate)
                If intDifferenance <= 0 Then
                    oApplication.Utilities.Message("Contract already exits for the selected Unit Code. The contract start date should be greater than previous contract end date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return True
                End If
            End If
            'strSql = "Select * from [@Z_CONTRACT] where U_Z_UnitCode='" & strUnitCode & "' and DocEntry<>" & intDocEntry & " and '" & dtEndDate.ToString("yyyy-MM-yy") & "' between U_Z_StartDate and U_Z_EndDate"
            'oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'oTemp.DoQuery(strSql)
            'If oTemp.RecordCount > 0 Then
            '    oApplication.Utilities.Message("Contract already exits for the selected Unit Code. The contract start date should be greater than previous contract end date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '    Return True
            'End If
            Return False


        End If
    End Function

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
            oCon.Alias = "U_Z_PROFLG"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_5")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_6")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)
            oCFL = oCFLs.Item("CFL_7")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()
            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
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
            oCombo = aForm.Items.Item("23").Specific
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

            'oCombo = aForm.Items.Item("65").Specific
            'For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
            '    Try
            '        oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            '    Catch ex As Exception

            '    End Try
            'Next
            'oCombo.ValidValues.Add("", "")

            'otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'otemp.DoQuery("SELECT T0.[DocEntry], T0.[U_Z_PolicyNumber] FROM [@Z_INSURANCE] T0")
            'For introw As Integer = 0 To otemp.RecordCount - 1
            '    oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            '    otemp.MoveNext()
            'Next
            'aForm.Items.Item("65").DisplayDesc = True

            oMatrix = aForm.Items.Item("50").Specific
            oColumn = oMatrix.Columns.Item("V_0")
            otemp.DoQuery("Select Code,U_Z_Name from [@Z_OEXP] order by Convert(numeric,Code)")
            oColumn.ValidValues.Add("", "")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oColumn.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oColumn.DisplayDesc = False
            oColumn = oMatrix.Columns.Item("V_2")
            oColumn.DisplayDesc = True
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oCombo = aForm.Items.Item("54").Specific
            oCombo.Select("T", SAPbouiCOM.BoSearchKey.psk_ByValue)
            aForm.Items.Item("6").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            aForm.Items.Item("54").Enabled = False
            aForm.Items.Item("68").Enabled = False
            Dim oButCombo As SAPbouiCOM.ButtonCombo
            oButCombo = aForm.Items.Item("77").Specific
            oButCombo.Caption = "Print Contract"
            oButCombo.ValidValues.Add("Print-Default", "Contract-Default")
            oButCombo.ValidValues.Add("Print-Awqaf", "Awqaf")
            oButCombo.ValidValues.Add("Print-RentGeneral", "Rent-General")
            oButCombo.ValidValues.Add("Print-Al Bayan", "Al Bayan")
            oButCombo.ValidValues.Add("Private Eng office", "Private Eng office")

            aForm.Items.Item("77").DisplayDesc = False
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
        strstring = "Select DocEntry,U_Z_Desc from [@Z_PROPUNIT] where U_Z_PropCode=" & strPropertyCode & " order by DocEntry"
        aForm.Freeze(True)
        oApplication.Utilities.FillComboBox(oCombo, strstring)
        aForm.Freeze(False)
    End Sub



    Private Sub PopulateExpenseDetails(ByVal amatrix As SAPbouiCOM.Matrix, ByVal intRow As Integer, ByVal aAmount As Double)
        oCombo = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific
        Dim strCode As String
        strCode = oCombo.Selected.Value
        Dim otemp As SAPbobsCOM.Recordset
        otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otemp.DoQuery("Select * from [@Z_OEXP]  where code='" & strCode & "' order by Convert(numeric,Code)")
        If otemp.RecordCount > 0 Then
            oApplication.Utilities.SetMatrixValues(amatrix, "V_5", intRow, otemp.Fields.Item("U_Z_CODE").Value)
            oApplication.Utilities.SetMatrixValues(amatrix, "V_4", intRow, otemp.Fields.Item("U_Z_NAME").Value)
            oApplication.Utilities.SetMatrixValues(amatrix, "V_3", intRow, otemp.Fields.Item("U_Z_GLACC").Value)
            oCombo = oMatrix.Columns.Item("V_2").Cells.Item(intRow).Specific
            oCombo.Select(otemp.Fields.Item("U_Z_TYPE").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            'oApplication.Utilities.SetMatrixValues(amatrix, "V_2", intRow, otemp.Fields.Item("U_Z_TYPE").Value)
            oApplication.Utilities.SetMatrixValues(amatrix, "V_6", intRow, otemp.Fields.Item("U_Z_RATE").Value)
            If otemp.Fields.Item("U_Z_TYPE").Value = "P" Then
                Dim dblPercentage As Double
                dblPercentage = otemp.Fields.Item("U_Z_RATE").Value
                aAmount = aAmount * dblPercentage / 100
                oApplication.Utilities.SetMatrixValues(amatrix, "V_1", intRow, aAmount)
            Else
                oApplication.Utilities.SetMatrixValues(amatrix, "V_1", intRow, otemp.Fields.Item("U_Z_RATE").Value)
            End If
        End If
    End Sub

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
            Exit Sub
        End If
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("50").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT1")
            Case "3"
                oMatrix = aForm.Items.Item("62").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT2")
        End Select
        Try
            aForm.Freeze(True)
            If oMatrix.RowCount <= 0 Then
                oMatrix.AddRow()
            End If

            Try

                Select Case aForm.PaneLevel
                    Case "1"
                        oCombo = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                        If oCombo.Selected.Value <> "" Then
                            oMatrix.AddRow()
                            oCombo = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                            oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
                            ' oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_5", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_4", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_3", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_6", oMatrix.RowCount, "0")
                            '                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", oMatrix.RowCount, "")
                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, "")
                            oCombo = oMatrix.Columns.Item("V_2").Cells.Item(oMatrix.RowCount).Specific
                            oCombo.Select("F", SAPbouiCOM.BoSearchKey.psk_ByValue)
                            oCombo = oMatrix.Columns.Item("V_7").Cells.Item(oMatrix.RowCount).Specific
                            oCombo.Select("M", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        End If
                    Case "3"
                        oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                        Try
                            If oEditText.Value <> "" Then
                                oMatrix.AddRow()
                                Select Case aForm.PaneLevel
                                    Case "3"
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
                                End Select
                            End If

                        Catch ex As Exception
                            aForm.Freeze(False)
                            oMatrix.AddRow()
                        End Try

                End Select

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
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
            Exit Sub
        End If
        Select Case aForm.PaneLevel
            Case "1"
                oMatrix = aForm.Items.Item("50").Specific
            Case "3"
                Exit Sub
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
        AssignLineNo(aForm)
        Return True
    End Function
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
            Exit Sub
        End If
        Me.MatrixId = "50"
        If aForm.PaneLevel = 3 Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT2")
            frmSourceMatrix = aForm.Items.Item("62").Specific
        ElseIf aForm.PaneLevel = 1 Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT1")
            frmSourceMatrix = aForm.Items.Item("50").Specific
        End If
        'If Me.MatrixId.ToString = "50" Then

        'Else

        'End If
        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT1")
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



#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strReservedate, strStartDate, strEndDate, strIssueDate, strInvoiceamt, strAnnualRent, strNumberofMonths, strContractNumber As String
        Dim dtreserverdate, dtstartdate, dtenddate, dtissuedate As String
        Dim dblAnnualrent, dblNumberofMonths As Double
        Dim dblInvoiceamt As Double
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If oApplication.Utilities.getEdittextvalue(aform, "43") = "" Then
                oApplication.Utilities.Message("Property Unit Code can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aform, "9") = "" Then
                oApplication.Utilities.Message("Tenent Code can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "4") = "" Then
            oApplication.Utilities.Message("Start date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strStartDate = oApplication.Utilities.getEdittextvalue(aform, "4")
            dtstartdate = oApplication.Utilities.GetDateTimeValue(strStartDate)
        End If

        If oApplication.Utilities.getEdittextvalue(aform, "6") = "" Then
            oApplication.Utilities.Message("End date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            strEndDate = oApplication.Utilities.getEdittextvalue(aform, "6")
            dtenddate = oApplication.Utilities.GetDateTimeValue(strEndDate)
        End If
        If dtstartdate < dtreserverdate Then
            ' oApplication.Utilities.Message("Start date should be greater than or equal to Reservation Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'Return False
        End If
        Dim intDifferenance As Integer
        intDifferenance = dateDifference(dtenddate, dtstartdate)
        If intDifferenance < 0 Then
            oApplication.Utilities.Message("End date should be greater than Start Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        End If
        If dtstartdate > dtissuedate Then
            ' oApplication.Utilities.Message("Issue date should be greater than Start Date", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'Return False
        End If
        If ValidateUnitExists(aform, dtstartdate, dtenddate) = True Then
            Return False
        End If
        strAnnualRent = oApplication.Utilities.getEdittextvalue(aform, "17")
        If strAnnualRent = "" Then
            oApplication.Utilities.Message("Annual rent is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            dblAnnualrent = oApplication.Utilities.getDocumentQuantity(strAnnualRent)
            If dblAnnualrent <= 0 Then
                oApplication.Utilities.Message("Annual rent should be greater than zero...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        '  MsgBox(oApplication.Utilities.getWords(dblAnnualrent.ToString))
        '   oApplication.Utilities.setEdittextvalue(aform, "47", (oApplication.Utilities.SFormatNumber(dblAnnualrent)))

        strNumberofMonths = oApplication.Utilities.getEdittextvalue(aform, "28")
        If strNumberofMonths = "" Then
            oApplication.Utilities.Message("Number of Months should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Return False
        Else
            dblNumberofMonths = oApplication.Utilities.getDocumentQuantity(strNumberofMonths)
            If dblNumberofMonths <= 0 Then
                oApplication.Utilities.Message("Number of Months should be greater than zero", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        If oApplication.Utilities.getEdittextvalue(aform, "49") = "" Then
            'oApplication.Utilities.Message("Reciable Account code is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            'Return False
        End If
        oCombo = aform.Items.Item("54").Specific
        If oCombo.Selected.Value = "T" Then
            If oApplication.Utilities.getEdittextvalue(aform, "56") = "" Then
                oApplication.Utilities.Message("Tenant receivable account code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If oApplication.Utilities.getEdittextvalue(aform, "58") = "" Then
                oApplication.Utilities.Message("Owner code is missing for Tenant Contrac Type", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        Dim strInsurance As String
        oCombo = aform.Items.Item("25").Specific
        Try
            strInsurance = oCombo.Selected.Value
        Catch ex As Exception
            strInsurance = "N"
        End Try

        If strInsurance = "Y" Then
            If oApplication.Utilities.getEdittextvalue(aform, "65") = "" Then
                oApplication.Utilities.Message("Insurance Policy Number is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If
        oMatrix = aform.Items.Item("50").Specific
        Dim strcode, strcode1 As String
        For intLoop As Integer = 1 To oMatrix.VisualRowCount
            oCombo = oMatrix.Columns.Item("V_0").Cells.Item(intLoop).Specific
            strcode = oCombo.Selected.Value
            If strcode <> "" Then
                For intLoop1 As Integer = 1 To oMatrix.VisualRowCount
                    oCombo = oMatrix.Columns.Item("V_0").Cells.Item(intLoop1).Specific
                    strcode1 = oCombo.Selected.Value
                    If strcode <> "" Then
                        If strcode = strcode1 And intLoop <> intLoop1 Then
                            oApplication.Utilities.Message("Expenes already exists", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Return False
                        End If

                    End If
                Next
                oCombo = oMatrix.Columns.Item("V_2").Cells.Item(intLoop).Specific
                If oCombo.Selected.Value = "P" Then
                    Dim dblPercentage, aamount As Double
                    dblPercentage = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getMatrixValues(oMatrix, "V_6", intLoop))
                    aamount = (dblAnnualrent / dblNumberofMonths) * dblPercentage / 100
                    oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", intLoop, aamount)
                End If
            End If
        Next
        Return True
    End Function
#End Region

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("50").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_CONTRACT1")
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

    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("62").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                Dim strFilename As String
                strFilename = oMatrix.Columns.Item("V_0").Cells.Item(intRow).Specific.value
                Dim x As System.Diagnostics.ProcessStartInfo
                x = New System.Diagnostics.ProcessStartInfo
                x.UseShellExecute = True
                x.FileName = strFilename
                System.Diagnostics.Process.Start(x)
                x = Nothing
                Exit Sub
            End If
        Next
        oApplication.Utilities.Message("No file has been selected...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

    End Sub

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
            ' otemp.DoQuery("select isnull(U_Z_Deposit,0) + isnull(U_Z_Salik,0)+isnull(U_Z_DPAmount,0), * from [@Z_ORDR] where DocEntry=" & acode)
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
                oDoc.UserFields.Fields.Item("U_SEQ").Value = otemp.Fields.Item("U_Z_SeqNo").Value
                oDoc.UserFields.Fields.Item("U_Z_STARTDATE").Value = otemp.Fields.Item("U_Z_StartDate").Value
                oDoc.UserFields.Fields.Item("U_Z_ENDDATE").Value = otemp.Fields.Item("U_Z_EndDate").Value
                oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = otemp.Fields.Item("U_Z_CntNo").Value
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
                    '    MsgBox(oBP.DownPaymentClearAct)
                    'T0.U_Z_ACCTCODE else T0.U_Z_LiaAc
                    '   MsgBox(oBP.DownPaymentClearAct)
                    '  oDoc.Lines.AccountCode = oBP.DownPaymentClearAct
                    ' oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(otemp.Fields.Item("U_Z_AcctCode").Value)
                    ' oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(otemp.Fields.Item("U_Z_LiaAc").Value)
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
        'CreateIncomingPayment(acode)
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
                oDoc.UserFields.Fields.Item("U_SEQ").Value = otemp.Fields.Item("U_Z_SeqNo").Value
                oDoc.UserFields.Fields.Item("U_Z_CONTNUMBER").Value = otemp.Fields.Item("U_Z_ConNo").Value
                oDoc.UserFields.Fields.Item("U_Z_CNTNUMBER").Value = otemp.Fields.Item("U_Z_CntNo").Value
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
                    '  oDoc.Lines.AccountCode = oBP.DownPaymentClearAct
                    'T0.U_Z_ACCTCODE else T0.U_Z_LiaAc
                    ' oDoc.Lines.AccountCode = otemp.Fields.Item("U_Z_AcctCode").Value ' oBP.DownPaymentClearAct
                    'oDoc.Lines.AccountCode = oApplication.Utilities.getAccountCode(strCreditAccountcode)
                    'oDoc.Lines.ItemDescription = "Security Deposit for UntiCode : " & otemp.Fields.Item("U_Z_UnitCode").Value
                    'If oBP.VatGroup <> "" Then
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

            End If
        End If
        'End If
        'CreateIncomingPayment(acode)
    End Function

#Region "Events"



    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_TenContracts Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "InvList1"
                        oCreationPackage.String = "Financial Transaction Details"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "Agreement1"
                        oCreationPackage.String = "Print Contract Agreement"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)

                        oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                        oCreationPackage.UniqueID = "Renewal"
                        oCreationPackage.String = "Renewal History"
                        oCreationPackage.Enabled = True
                        oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        oMenus.AddEx(oCreationPackage)
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                        oApplication.SBO_Application.Menus.RemoveEx("InvList1")
                        oApplication.SBO_Application.Menus.RemoveEx("Renewal")
                        oApplication.SBO_Application.Menus.RemoveEx("Agreement1")
                    End If
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If
    End Sub

    Private Sub PopulateCustomerAddress(ByVal aCode As String)
        Dim oTest As SAPbobsCOM.Recordset
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If oApplication.Utilities.getEdittextvalue(oForm, "58") <> "" Then
            oTest.DoQuery("Select * from OCRD where CardCode='" & aCode & "'")
            If oTest.RecordCount <= 0 Then
                Exit Sub
            End If
        Else
            '  Exit Sub
        End If
        oTest.DoQuery("select CntctPrsn , MailAddres ,Phone1,isnull(Phone2,'') + ',' + isnull(Cellular,''), Fax,E_Mail ,* from OCRD where CardCode='" & aCode & "'")
        Dim str As String
        str = oTest.Fields.Item(0).Value & "," & oTest.Fields.Item(1).Value & "," & oTest.Fields.Item(2).Value & "," & oTest.Fields.Item(3).Value & "," & oTest.Fields.Item(4).Value & "," & oTest.Fields.Item(5).Value
        Try
            oApplication.Utilities.setEdittextvalue(oForm, "9", aCode)
            oApplication.Utilities.setEdittextvalue(oForm, "13", str)
        Catch ex As Exception
        End Try
        oApplication.Utilities.setEdittextvalue(oForm, "13", str)
    End Sub
#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_TenContracts
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case "Renewal"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm
                    If oForm.TypeEx = frm_TenContracts Then
                        Dim strcode As String
                        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                        Dim strNo As String
                        Try
                            strNo = oForm.Items.Item("1000002").Specific.value
                        Catch ex As Exception
                            strNo = ""
                        End Try
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                            Dim oObj As New clsRenewalHistory
                            oObj.LoadForm(strNo, "Booking")
                        End If
                    End If

                Case "InvList1"
                    Dim strcode As String
                    oForm = oApplication.SBO_Application.Forms.ActiveForm
                    If oForm.TypeEx = frm_TenContracts Then
                        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                        Dim strNo As String
                        Try
                            strNo = oForm.Items.Item("1000002").Specific.value
                        Catch ex As Exception
                            strNo = ""
                        End Try
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                            Dim oObj As New clsDocumentsView
                            oObj.LoadForm(strNo, "Booking")
                        End If
                    End If

                Case "Agreement1"
                    Dim strcode As String
                    oForm = oApplication.SBO_Application.Forms.ActiveForm
                    If oForm.TypeEx = frm_TenContracts Then
                        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                        Dim strNo As String
                        Try
                            strNo = oForm.Items.Item("1000002").Specific.value
                        Catch ex As Exception
                            strNo = ""
                        End Try
                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                            'Dim oObj As New clsPrint
                            'oObj.PrintContract_Tenant(strNo)
                            Dim oObj As New clsPrint
                            'oObj.PrintContract(strNo)
                            oCombo = oForm.Items.Item("54").Specific
                            Dim stType As String
                            Try
                                stType = oCombo.Selected.Value
                            Catch ex As Exception
                                stType = "O"
                            End Try
                            If stType = "T" Then
                                oObj.PrintContract_Tenant(strNo, "ContractTen")
                            Else
                                oObj.PrintContract(strNo, "Contract")
                            End If
                        End If
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "1000002", oApplication.Utilities.getMaxCode("@Z_CONTRACT", "DocEntry"))
                        oForm.Items.Item("40").Enabled = True
                        oApplication.Utilities.setEdittextvalue(oForm, "40", "t")
                        oApplication.SBO_Application.SendKeys("{TAB}")
                        oForm.Items.Item("40").Enabled = True
                        oForm.Items.Item("68").Enabled = False
                        oForm.Items.Item("43").Enabled = True
                        oCombo = oForm.Items.Item("25").Specific
                        oCombo.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)
                        GetNextNumber(oForm)
                    End If

                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    oForm.Items.Item("43").Enabled = True
                    oForm.Items.Item("68").Enabled = True
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    End If
                Case mnu_DELETE_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        RefereshDeleteRow(oForm)
                    Else
                        If ValidateDeletion(oForm) = False Then
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm = oApplication.SBO_Application.Forms.ActiveForm()
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub LoadForm_Contract(ByVal aCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_TenContracts, frm_TenContracts)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "1000002"
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.Items.Item("54").Enabled = True
        oCombo = oForm.Items.Item("54").Specific
        oCombo.Select("T", SAPbouiCOM.BoSearchKey.psk_ByValue)
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from [@Z_PropUnit] where DocEntry=" & aCode)
        If otest.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(oForm, "43", otest.Fields.Item("U_Z_ProItemCode").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "edDesc", otest.Fields.Item("U_Z_Desc").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "58", otest.Fields.Item("U_Z_OwnerCode").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "edComm", otest.Fields.Item("U_Z_Comm").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "40", "T")
            oApplication.SBO_Application.SendKeys("{TAB}")
            GetNextNumber(oForm)
        End If
        oCombo = oForm.Items.Item("25").Specific
        oCombo.Select("N", SAPbouiCOM.BoSearchKey.psk_ByValue)

        oForm.Items.Item("54").Enabled = False
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub

    Public Sub LoadForm_Contract_View(ByVal aCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_TenContracts, frm_TenContracts)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.DataBrowser.BrowseBy = "1000002"
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

        oForm.Items.Item("68").Enabled = True
        oApplication.Utilities.setEdittextvalue(oForm, "68", aCode)
        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    oForm.Items.Item("40").Enabled = False
                    oForm.Items.Item("43").Enabled = False
                Else
                    oForm.Items.Item("40").Enabled = True
                    oForm.Items.Item("43").Enabled = True
                End If
                oCombo = oForm.Items.Item("54").Specific
                If oCombo.Selected.Value = "O" Then
                    oForm.Items.Item("18").Visible = False
                    oForm.Items.Item("19").Visible = False
                    oForm.Items.Item("120").Visible = False
                    oForm.Items.Item("60").Visible = False
                    oForm.Items.Item("49").Visible = False
                    oForm.Items.Item("1000004").Visible = False
                Else

                    oForm.Items.Item("18").Visible = True
                    oForm.Items.Item("19").Visible = True
                    oForm.Items.Item("120").Visible = True
                    oForm.Items.Item("60").Visible = True
                    oForm.Items.Item("49").Visible = False
                    oForm.Items.Item("1000004").Visible = False

                End If
                oCombo = oForm.Items.Item("21").Specific
                If oCombo.Selected.Value = "TER" Then
                    oForm.Items.Item("1").Enabled = False
                    oForm.Items.Item("120").Enabled = False
                    oForm.Items.Item("60").Enabled = False
                Else
                    oForm.Items.Item("1").Enabled = True
                    oForm.Items.Item("120").Enabled = True
                    oForm.Items.Item("60").Enabled = True

                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#Region "Get Next Number"

    Private Sub GetNextNumber(ByVal aForm As SAPbouiCOM.Form)
        Dim oTest As SAPbobsCOM.Recordset
        Dim strType As String
        If aForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            oCombo = aForm.Items.Item("54").Specific
            If oCombo.Selected.Value <> "T" Then
                oCombo.Select("T", SAPbouiCOM.BoSearchKey.psk_ByValue)
            End If
            Dim stQuery As String
            strType = oCombo.Selected.Value
            oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTest.DoQuery("Select count(*) +1 from [@Z_CONTRACT] where U_Z_Type='" & strType & "'")
            Dim dblCount As Integer
            dblCount = oTest.Fields.Item(0).Value
            Select Case strType
                Case "O"
                    stQuery = "Select isnull(max(convert(numeric,replace(U_Z_CONNO,'CM',''))),0) +1from [@Z_CONTRACT] where U_Z_Type='" & strType & "'"
                    oTest.DoQuery(stQuery)
                    dblCount = oTest.Fields.Item(0).Value
                    strType = "CM" & dblCount.ToString("0000")
                Case "T"
                    stQuery = "Select isnull(max(convert(numeric,replace(U_Z_CONNO,'CR',''))),0) +1from [@Z_CONTRACT] where U_Z_Type='" & strType & "'"
                    oTest.DoQuery(stQuery)
                    dblCount = oTest.Fields.Item(0).Value
                    strType = "CR" & dblCount.ToString("0000")
            End Select
            oApplication.Utilities.setEdittextvalue(aForm, "68", strType)
            oApplication.Utilities.setEdittextvalue(aForm, "81", "1")
            strType = strType & "_1"
            oApplication.Utilities.setEdittextvalue(aForm, "82", strType)

        End If
    End Sub


    Private Function getNewContractID(ByVal aForm As SAPbouiCOM.Form) As String
        Dim oTest As SAPbobsCOM.Recordset
        Dim strType As String

        strType = "CR0001"
        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTest.DoQuery("Select count(*) +1 from [@Z_CONTRACT] where U_Z_Type='T'")
        Dim dblCount As Integer
        dblCount = oTest.Fields.Item(0).Value
        Select Case strType
            Case "O"
                strType = "CM" & dblCount.ToString("0000")
            Case "T"
                strType = "CR" & dblCount.ToString("0000")
        End Select

        Return strType

    End Function
#End Region

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_TenContracts Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "56" Or pVal.ItemUID = "58") Then
                                    oCombo = oForm.Items.Item("54").Specific
                                    Dim stType As String
                                    Try
                                        stType = oCombo.Selected.Value
                                    Catch ex As Exception
                                        stType = "O"
                                    End Try
                                    If stType = "O" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "56" Or pVal.ItemUID = "58") Then
                                    oCombo = oForm.Items.Item("54").Specific
                                    Dim stType As String
                                    Try
                                        stType = oCombo.Selected.Value
                                    Catch ex As Exception
                                        stType = "O"
                                    End Try
                                    If stType = "O" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                If pVal.ItemUID = "50" Or pVal.ItemUID = "62" Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = pVal.ItemUID
                                    frmSourceMatrix = oMatrix
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                                Select Case pVal.ItemUID
                                    Case "85"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            Dim oObj As New clsInstallment

                                            oObj.LoadForm(oApplication.Utilities.getEdittextvalue(oForm, "1000002"))
                                            BubbleEvent = False
                                            Exit Sub
                                        End If
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("50").Specific
                                If pVal.ItemUID = "50" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "50"
                                    frmSourceMatrix = oMatrix
                                ElseIf pVal.ItemUID = "65" Then
                                    Dim strInsurance As String
                                    oCombo = oForm.Items.Item("25").Specific
                                    strInsurance = oCombo.Selected.Value
                                    If strInsurance = "N" Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If



                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Try
                                    oForm.Items.Item("41").Width = oForm.Items.Item("51").Left + oForm.Items.Item("51").Width + 5
                                Catch ex As Exception
                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If (pVal.ItemUID = "17" Or pVal.ItemUID = "28") And pVal.CharPressed = 9 Then
                                    Dim dblNoMonth, dblAnnual, dblMonth As Double
                                    dblNoMonth = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "28"))
                                    dblAnnual = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "17"))

                                    dblMonth = dblAnnual / dblNoMonth
                                    oApplication.Utilities.setEdittextvalue(oForm, "84", dblMonth)
                                End If
                                If (pVal.ItemUID = "4" Or pVal.ItemUID = "6") And pVal.CharPressed = 9 Then
                                    Dim strFromDate, strToDate As String
                                    Dim dtFromDate, dttoDate As Date
                                    strFromDate = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    strToDate = oApplication.Utilities.getEdittextvalue(oForm, "6")
                                    If strFromDate <> "" Then
                                        dtFromDate = oApplication.Utilities.GetDateTimeValue(strFromDate)
                                    End If
                                    If strToDate <> "" Then
                                        dttoDate = oApplication.Utilities.GetDateTimeValue(strToDate)
                                    End If
                                    If strFromDate <> "" And strToDate <> "" Then
                                        Dim intNoofMonths As Integer
                                        intNoofMonths = DateDiff(DateInterval.Month, dtFromDate, dttoDate)
                                        If intNoofMonths = 0 Then
                                            intNoofMonths = 1
                                        End If
                                        ' oApplication.Utilities.setEdittextvalue(oForm, "28", DateDiff(DateInterval.Month, dtFromDate, dttoDate))
                                        oApplication.Utilities.setEdittextvalue(oForm, "28", intNoofMonths)
                                        Dim dblNoMonth, dblAnnual, dblMonth As Double
                                        dblNoMonth = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "28"))
                                        dblAnnual = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "17"))
                                        dblMonth = dblAnnual / dblNoMonth
                                        oApplication.Utilities.setEdittextvalue(oForm, "84", dblMonth)
                                    End If


                                End If
                                If pVal.ItemUID = "65" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    Dim strIns As String
                                    oCombo = oForm.Items.Item("25").Specific
                                    If oCombo.Selected.Value = "N" Then
                                        Exit Sub
                                    End If
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "65")
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_INSURANCE] where U_Z_Unitcode='" & oApplication.Utilities.getEdittextvalue(oForm, "43") & "' and U_Z_Policynumber='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "65", "")
                                        strIns = ""
                                    Else
                                        Exit Sub
                                    End If
                                    Dim objChooseForm As SAPbouiCOM.Form
                                    Dim objChoose As New clsChooseFromList
                                    clsChooseFromList.ItemUID = pVal.ItemUID
                                    clsChooseFromList.SourceFormUID = FormUID
                                    clsChooseFromList.SourceLabel = 0
                                    clsChooseFromList.CFLChoice = "Insurance" 'oCombo.Selected.Value
                                    clsChooseFromList.choice = "Bin"
                                    clsChooseFromList.Documentchoice = oApplication.Utilities.getEdittextvalue(oForm, "9") 'TenCode
                                    clsChooseFromList.ItemCode = oApplication.Utilities.getEdittextvalue(oForm, "43") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsChooseFromList.sourceColumID = ""
                                    oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
                                If pVal.ItemUID = "43" And pVal.CharPressed = 9 Then
                                    oForm = oApplication.SBO_Application.Forms.Item(FormUID)

                                    Dim strIns As String
                                    strIns = oApplication.Utilities.getEdittextvalue(oForm, "43")
                                    Dim otest As SAPbobsCOM.Recordset
                                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    otest.DoQuery("Select * from [@Z_PROPUNIT] where U_Z_ProItemCode='" & strIns & "'")
                                    If otest.RecordCount <= 0 Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "43", "")
                                        strIns = ""
                                    Else
                                        otest.DoQuery("Select isnull(U_Z_PRICE,0) from [@Z_PRICELIST] where U_Z_PrlNam='" & strIns & "'")
                                        oApplication.Utilities.setEdittextvalue(oForm, "17", otest.Fields.Item(0).Value)
                                        oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
                                    clsChooseFromList.ItemCode = "" 'oApplication.Utilities.getEdittextvalue(oForm, "6") 'Unit Code
                                    ' clsChooseFromList.BinDescrUID = "BinToBinHeader"
                                    clsChooseFromList.sourceColumID = "edDesc"
                                    oApplication.Utilities.LoadForm("\CFL.xml", frm_ChoosefromList)
                                    objChooseForm = oApplication.SBO_Application.Forms.ActiveForm()
                                    objChoose.databound(objChooseForm)
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oMatrix = oForm.Items.Item("50").Specific
                                If pVal.ItemUID = "54" Then
                                    GetNextNumber(oForm)
                                End If

                                If pVal.ItemUID = "77" Then
                                    Dim strcode As String
                                    oForm = oApplication.SBO_Application.Forms.ActiveForm
                                    If oForm.TypeEx = frm_TenContracts Then
                                        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                        Dim strNo As String
                                        Try
                                            strNo = oForm.Items.Item("1000002").Specific.value
                                        Catch ex As Exception
                                            strNo = ""
                                        End Try
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                            oCombo = oForm.Items.Item("54").Specific
                                            Dim stType As String
                                            Try
                                                stType = oCombo.Selected.Value
                                            Catch ex As Exception
                                                stType = "O"
                                            End Try
                                            Dim ocombobutton As SAPbouiCOM.ButtonCombo
                                            ocombobutton = oForm.Items.Item("77").Specific
                                            Dim oObj As New clsPrint
                                            If stType = "T" Then
                                                oObj.PrintContract_Tenant(strNo, ocombobutton.Selected.Description)
                                            Else
                                                oObj.PrintContract(strNo, "Contract")
                                            End If

                                        End If
                                    End If


                                End If
                                If pVal.ItemUID = "50" And pVal.ColUID = "V_0" Then
                                    oForm.Freeze(True)
                                    Dim dblRent, dblNoofMonths As Double
                                    dblRent = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "17"))
                                    dblNoofMonths = oApplication.Utilities.getDocumentQuantity(oApplication.Utilities.getEdittextvalue(oForm, "28"))
                                    dblRent = dblRent / dblNoofMonths
                                    PopulateExpenseDetails(oMatrix, pVal.Row, dblRent)
                                    oCombo = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                    oCombo.Select("M", SAPbouiCOM.BoSearchKey.psk_ByValue)
                                    oForm.Freeze(False)
                                ElseIf pVal.ItemUID = "25" Then
                                    oCombo = oForm.Items.Item("25").Specific
                                    If oCombo.Selected.Value = "Y" Then
                                        oApplication.Utilities.setEdittextvalue(oForm, "65", "")
                                        'oCombo = oForm.Items.Item("65").Specific
                                        'oApplication.Utilities.FillComboBox(oCombo, "Select DocEntry,U_Z_PolicyNumber from [@Z_Insurance]  where U_Z_UnitCode='" & oApplication.Utilities.getEdittextvalue(oForm, "43") & "' and ( isnull(U_Z_TenCode,'')='' or  U_Z_TenCode='" & oApplication.Utilities.getEdittextvalue(oForm, "9") & "') order by Docentry ")
                                    Else
                                        oApplication.Utilities.setEdittextvalue(oForm, "65", "")
                                        'oCombo = oForm.Items.Item("65").Specific
                                        'oApplication.Utilities.FillComboBox(oCombo, "Select DocEntry,U_Z_PolicyNumber from [@Z_Insurance]  where 1=2 order by Docentry ")
                                    End If
                                    '                                    oForm.Items.Item("65").DisplayDesc = True
                                End If
                                If pVal.ItemUID = "50" And pVal.ColUID = "V_7" Then
                                    Dim dtDate, dtNextDate As Date
                                    Dim strDate, strFrequency, strMonth As String
                                    strDate = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                    If strDate <> "" Then
                                        dtDate = oApplication.Utilities.GetDateTimeValue(strDate)
                                        oCombo = oMatrix.Columns.Item("V_7").Cells.Item(pVal.Row).Specific
                                        strFrequency = oCombo.Selected.Value
                                        strMonth = ""
                                        Select Case strFrequency
                                            Case "M"
                                                'strMonth = ""
                                                For intRow As Integer = 1 To 12
                                                    If strMonth = "" Then
                                                        strMonth = "'" & MonthName(intRow) & "'"
                                                    Else
                                                        strMonth = strMonth & ",'" & MonthName(intRow) & "'"
                                                    End If
                                                Next
                                                '                                                strMonth = "1,2,3,4,5,6,7,8,9,10,11,12"
                                            Case "Q"
                                                strMonth = MonthName(dtDate.Month)
                                                dtDate = DateAdd(DateInterval.Month, 3, dtDate)
                                                strMonth = "'" & strMonth & "','" & MonthName(dtDate.Month) & "',"
                                                dtDate = DateAdd(DateInterval.Month, 3, dtDate)
                                                strMonth = strMonth & "'" & MonthName(dtDate.Month) & "',"
                                                dtDate = DateAdd(DateInterval.Month, 3, dtDate)
                                                strMonth = strMonth & "'" & MonthName(dtDate.Month) & "'"
                                            Case "H"
                                                strMonth = MonthName(dtDate.Month)
                                                dtDate = DateAdd(DateInterval.Month, 6, dtDate)
                                                strMonth = "'" & strMonth & "','" & MonthName(dtDate.Month) & "'"
                                            Case "Y"
                                                strMonth = "'" & MonthName(dtDate.Month) & "'"
                                        End Select
                                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_8", pVal.Row, strMonth)


                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "42"
                                        oForm.PaneLevel = 1
                                    Case "45"
                                        oForm.PaneLevel = 2
                                    Case "61"
                                        oForm.PaneLevel = 3
                                    Case "51"
                                        AddRow(oForm)
                                    Case "77"
                                        Dim strcode As String
                                        oForm = oApplication.SBO_Application.Forms.ActiveForm
                                        If oForm.TypeEx = frm_TenContracts Then
                                            strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                            Dim strNo As String
                                            Try
                                                strNo = oForm.Items.Item("1000002").Specific.value
                                            Catch ex As Exception
                                                strNo = ""
                                            End Try
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                                oCombo = oForm.Items.Item("54").Specific
                                                Dim stType As String
                                                Try
                                                    stType = oCombo.Selected.Value
                                                Catch ex As Exception
                                                    stType = "O"
                                                End Try
                                                Dim ocombobutton As SAPbouiCOM.ButtonCombo
                                                ocombobutton = oForm.Items.Item("77").Specific
                                                Dim oObj As New clsPrint
                                                If stType = "T" Then
                                                    oObj.PrintContract_Tenant(strNo, ocombobutton.Selected.Description)
                                                Else
                                                    oObj.PrintContract(strNo, "Contract")
                                                End If

                                            End If
                                        End If
                                    Case "52"
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                        'Case "69"
                                        '    Dim strcode As String
                                        '    oForm = oApplication.SBO_Application.Forms.ActiveForm
                                        '    If oForm.TypeEx = frm_TenContracts Then
                                        '        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                        '        Dim strNo As String
                                        '        Try
                                        '            strNo = oForm.Items.Item("1000002").Specific.value
                                        '        Catch ex As Exception
                                        '            strNo = ""
                                        '        End Try
                                        '        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                        '            oCombo = oForm.Items.Item("54").Specific
                                        '            Dim stType As String
                                        '            Try
                                        '                stType = oCombo.Selected.Value
                                        '            Catch ex As Exception
                                        '                stType = "O"
                                        '            End Try
                                        '            Dim oObj As New clsPrint
                                        '            If stType = "T" Then
                                        '                oObj.PrintContract_Tenant(strNo, "ContractTen")
                                        '            Else
                                        '                oObj.PrintContract(strNo, "Contract")
                                        '            End If

                                        '        End If
                                        '    End If


                                        'Case "77"
                                        '    Dim strcode As String
                                        '    oForm = oApplication.SBO_Application.Forms.ActiveForm
                                        '    If oForm.TypeEx = frm_TenContracts Then
                                        '        strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                        '        Dim strNo As String
                                        '        Try
                                        '            strNo = oForm.Items.Item("1000002").Specific.value
                                        '        Catch ex As Exception
                                        '            strNo = ""
                                        '        End Try
                                        '        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                        '            oCombo = oForm.Items.Item("54").Specific
                                        '            Dim stType As String
                                        '            Try
                                        '                stType = oCombo.Selected.Value
                                        '            Catch ex As Exception
                                        '                stType = "O"
                                        '            End Try

                                        '            Dim oObj As New clsPrint
                                        '            If stType = "T" Then
                                        '                oObj.PrintContract_Tenant(strNo, "Rent-General")
                                        '            Else
                                        '                oObj.PrintContract(strNo, "Contract")
                                        '            End If

                                        '        End If
                                        '    End If

                                    Case "78"
                                        Dim strcode As String
                                        oForm = oApplication.SBO_Application.Forms.ActiveForm
                                        If oForm.TypeEx = frm_TenContracts Then
                                            strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                            Dim strNo As String
                                            Try
                                                strNo = oForm.Items.Item("1000002").Specific.value
                                            Catch ex As Exception
                                                strNo = ""
                                            End Try
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                                oCombo = oForm.Items.Item("54").Specific
                                                Dim stType As String
                                                Try
                                                    stType = oCombo.Selected.Value
                                                Catch ex As Exception
                                                    stType = "O"
                                                End Try
                                                Dim oObj As New clsPrint
                                                If stType = "T" Then
                                                    oObj.PrintContract_Tenant(strNo, "Awqaf")
                                                Else
                                                    oObj.PrintContract(strNo, "Contract")
                                                End If

                                            End If
                                        End If

                                    Case "79"
                                        Dim strcode As String
                                        oForm = oApplication.SBO_Application.Forms.ActiveForm
                                        If oForm.TypeEx = frm_TenContracts Then
                                            strcode = oApplication.Utilities.getEdittextvalue(oForm, "1000002")
                                            Dim strNo As String
                                            Try
                                                strNo = oForm.Items.Item("1000002").Specific.value
                                            Catch ex As Exception
                                                strNo = ""
                                            End Try
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE And strNo <> "" Then
                                                oCombo = oForm.Items.Item("54").Specific
                                                Dim stType As String
                                                Try
                                                    stType = oCombo.Selected.Value
                                                Catch ex As Exception
                                                    stType = "O"
                                                End Try
                                                Dim oObj As New clsPrint
                                                If stType = "T" Then
                                                    oObj.PrintContract_Tenant(strNo, "Al Bayan")
                                                Else
                                                    oObj.PrintContract(strNo, "Contract")
                                                End If

                                            End If
                                        End If

                                    Case "63"
                                        fillopen()
                                        oMatrix = oForm.Items.Item("62").Specific
                                        AddRow(oForm)
                                        Try
                                            oForm.Freeze(True)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, strSelectedFilepath)
                                            Dim strDate As String
                                            Dim dtdate As Date
                                            dtdate = Now.Date
                                            strDate = Now.Date.Today().ToString
                                            ''  strdate=
                                            Dim oColumn As SAPbouiCOM.Column
                                            oColumn = oMatrix.Columns.Item("V_1")
                                            oColumn.Editable = True

                                            oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                                            oEditText.String = "t"
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            oForm.Items.Item("13").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            oColumn.Editable = False

                                            'oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, dtdate)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oForm.Freeze(False)
                                        Catch ex As Exception
                                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oForm.Freeze(False)

                                        End Try
                                    Case "64"
                                        LoadFiles(oForm)
                                    Case "67"
                                        oMatrix = oForm.Items.Item("62").Specific
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)

                                    Case "120"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oCombo = oForm.Items.Item("21").Specific
                                            If oCombo.Selected.Value = "TER" Or oCombo.Selected.Value = "CAN" Then
                                                oApplication.Utilities.Message("Contract Already terminated  / Canceled", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                Exit Sub
                                            End If
                                            CreateDownPaymentInvoice(CInt(oApplication.Utilities.getEdittextvalue(oForm, "1000002")))
                                            ' CreateSecurityDepost(CInt(oApplication.Utilities.getEdittextvalue(oForm, "1000002")))
                                        End If

                                    Case "60"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                            oCombo = oForm.Items.Item("21").Specific
                                            If oCombo.Selected.Value = "TER" Or oCombo.Selected.Value = "CAN" Then
                                                oApplication.Utilities.Message("Contract Already terminated  / Canceled", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                                Exit Sub
                                            End If
                                            CreateSecurityDepost(CInt(oApplication.Utilities.getEdittextvalue(oForm, "1000002")))
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
                                        If pVal.ItemUID = "9" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val2 = oDataTable.GetValue("CardName", 0)
                                            val1 = oDataTable.GetValue("GroupNum", 0)
                                            oCombo = oForm.Items.Item("23").Specific
                                            oCombo.Select(val1, SAPbouiCOM.BoSearchKey.psk_ByValue)
                                            oApplication.Utilities.setEdittextvalue(oForm, "48", val2)
                                            'oCombo = oForm.Items.Item("65").Specific
                                            'oApplication.Utilities.FillComboBox(oCombo, "Select DocEntry,U_Z_PolicyNumber from [@Z_Insurance]  where U_Z_UnitCode='" & oApplication.Utilities.getEdittextvalue(oForm, "43") & "' and ( isnull(U_Z_TenCode,'')='' or  U_Z_TenCode='" & val & "') order by Docentry ")
                                            ' oForm.Items.Item("65").DisplayDesc = True
                                            oApplication.Utilities.setEdittextvalue(oForm, "65", "")
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                            End Try

                                            PopulateCustomerAddress(val)
                                            'ElseIf pVal.ItemUID = "43" Then
                                            '    val = oDataTable.GetValue("ItemCode", 0)
                                            '    oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "58" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "49" Or pVal.ItemUID = "76" Or pVal.ItemUID = "56" Or pVal.ItemUID = "72" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
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
