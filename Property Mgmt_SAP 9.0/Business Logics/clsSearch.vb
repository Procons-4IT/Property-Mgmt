Public Class clsSearch
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

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Search) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Search, frm_Search)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Databind(oForm)
    End Sub
    Private Sub Databind(ByVal aForm As SAPbouiCOM.Form)
        Dim otest As SAPbobsCOM.Recordset
        Try
            aForm.Freeze(True)
            aForm.DataSources.UserDataSources.Add("frmprice", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            aForm.DataSources.UserDataSources.Add("Toprice", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            aForm.DataSources.UserDataSources.Add("frmSpace", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            aForm.DataSources.UserDataSources.Add("ToSpace", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER)
            aForm.DataSources.UserDataSources.Add("dt", SAPbouiCOM.BoDataType.dt_DATE)

            otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otest.DoQuery("select DocEntry,U_Z_UNITTYPE from [@Z_UnitType] order by DocEntry")
            oCombo = aForm.Items.Item("4").Specific
            oCombo.ValidValues.Add("", "")
            For intRow As Integer = 0 To otest.RecordCount - 1
                oCombo.ValidValues.Add(otest.Fields.Item(0).Value, otest.Fields.Item(1).Value)
                otest.MoveNext()
            Next
            aForm.Items.Item("4").DisplayDesc = True
            oCombo = aForm.Items.Item("6").Specific
            oCombo.ValidValues.Add("C", "With Contract")
            oCombo.ValidValues.Add("O", "With out Contract")
            oCombo.ValidValues.Add("E", "Expired Contract")
            aForm.Items.Item("6").DisplayDesc = True

            oApplication.Utilities.setUserDatabind(aForm, "8", "dt")
            oApplication.Utilities.setUserDatabind(aForm, "11", "frmprice")
            oApplication.Utilities.setUserDatabind(aForm, "13", "ToPrice")
            oApplication.Utilities.setUserDatabind(aForm, "16", "frmSpace")
            oApplication.Utilities.setUserDatabind(aForm, "18", "ToSpace")
            strSQL = ""
            strSQL = " select T1.U_Z_PROPDESC,T1.U_Z_PROITEMCODE,T1.U_Z_CODE,T1.U_Z_DESC,T2.U_Z_UNITTYPE,T0.DocEntry,U_Z_TENCODE,U_Z_TENNAme,U_Z_STARTDATE,U_Z_ENDDATE, case T0.U_Z_STATUS when 'PED' then 'Pending' when 'APP' then 'Approved' when 'AGR' then 'Agreed' when 'TER' then 'Terminated' else 'Cancelled' end , T1.U_Z_SPACE,T1.U_Z_PRICE,T1.U_Z_TOTALAREA,T1.U_Z_TOTALPRICE, U_Z_OFFADDRESS,U_Z_ANNUALRENT from [@Z_CONTRACT] T0"
            strSQL = strSQL & " Left outer join [@Z_PROPUNIT] T1 on T1.U_Z_PROITEMCODE=T0.U_Z_UNITCODE  inner join   [@Z_UnitType] T2 on T2.DocEntry=T1.U_Z_UNITTYPE"
            strSQL = strSQL & " where 1=2"
            oGrid = aForm.Items.Item("20").Specific
            oGrid.DataTable.ExecuteQuery(strSQL)
            FormatGrid("C", oGrid)
            aForm.Items.Item("20").Enabled = False


            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try


    End Sub
#End Region

#Region "Display Result"
    Private Sub DisplayResult(ByVal aForm As SAPbouiCOM.Form)
        Dim otest, otest1 As SAPbobsCOM.Recordset
        Dim strFromPrice, strToPrice, strFromSpace, strToSpace, strDate, strPriceCondition, strSpaceCondition, strUnitCondition, strCondition, strreportChoice, strLocation, strsql As String
        Dim dblFromPrice, dblToPrice, dblFromSpace, dblToSpace As Double
        Dim dtDate As Date


        Try
            aForm.Freeze(True)
            strFromPrice = oApplication.Utilities.getEdittextvalue(aForm, "11")
            strToPrice = oApplication.Utilities.getEdittextvalue(aForm, "13")
            strFromSpace = oApplication.Utilities.getEdittextvalue(aForm, "16")
            strToSpace = oApplication.Utilities.getEdittextvalue(aForm, "18")
            strDate = oApplication.Utilities.getEdittextvalue(aForm, "8")

            If strDate <> "" Then
                dtDate = oApplication.Utilities.GetDateTimeValue(strDate)
            End If
            If strFromPrice <> "" Then
                dblFromPrice = oApplication.Utilities.getDocumentQuantity(strFromPrice)
                strPriceCondition = " T1.U_Z_TOTALPRICE > =" & dblFromPrice
            Else
                strPriceCondition = " 1=1"
            End If
            If strToPrice <> "" Then
                dblToPrice = oApplication.Utilities.getDocumentQuantity(strToPrice)
                strPriceCondition = strPriceCondition & " and T1.U_Z_TOTALPRICE < =" & dblToPrice
            Else
                strPriceCondition = strPriceCondition & " and 1=1"
            End If
            If strFromSpace <> "" Then
                dblFromSpace = oApplication.Utilities.getDocumentQuantity(strFromSpace)
                strSpaceCondition = " T1.U_Z_TOTALAREA > =" & dblFromSpace
            Else
                strSpaceCondition = " 1=1"
            End If
            If strToSpace <> "" Then
                dblToSpace = oApplication.Utilities.getDocumentQuantity(strToSpace)
                strSpaceCondition = strSpaceCondition & " and T1.U_Z_TOTALAREA < =" & dblToSpace
            Else
                strSpaceCondition = strSpaceCondition & " and 1=1"
            End If

            oCombo = aForm.Items.Item("4").Specific
            strLocation = oCombo.Selected.Value



            If strLocation <> "" Then
                strUnitCondition = " T1.U_Z_UNITTYPE=" & strLocation
            Else
                strUnitCondition = " 1=1"
            End If

            oCombo = aForm.Items.Item("6").Specific
            Try
                strreportChoice = oCombo.Selected.Value
            Catch ex As Exception
                oApplication.Utilities.Message("Property Available Method is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
            End Try


            If strreportChoice = "" Then
                oApplication.Utilities.Message("Select the available property method", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                aForm.Freeze(False)
                Exit Sub
            ElseIf strreportChoice = "E" Then
                If strDate = "" Then
                    oApplication.Utilities.Message("Expiring date can not be empty", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    aForm.Freeze(False)
                    Exit Sub
                Else
                    dtDate = oApplication.Utilities.GetDateTimeValue(strDate)
                End If
            End If


            Select Case strreportChoice
                Case "C" 'with Contract
                    strsql = " select T1.U_Z_PROPDESC,T1.U_Z_PROITEMCODE,T1.U_Z_CODE,T1.U_Z_DESC,T2.U_Z_UNITTYPE,  T1.U_Z_SPACE,T1.U_Z_PRICE,T1.U_Z_TOTALAREA,T1.U_Z_TOTALPRICE,T0.DocEntry,U_Z_TENCODE,U_Z_TENNAme,U_Z_STARTDATE,U_Z_ENDDATE, case T0.U_Z_STATUS when 'PED' then 'Pending' when 'APP' then 'Approved' when 'AGR' then 'Agreed' when 'TER' then 'Terminated' else 'Cancelled' end , U_Z_OFFADDRESS,U_Z_ANNUALRENT from [@Z_CONTRACT] T0"
                    strsql = strsql & " Left outer join [@Z_PROPUNIT] T1 on T1.U_Z_PROITEMCODE=T0.U_Z_UNITCODE  inner join   [@Z_UnitType] T2 on T2.DocEntry=T1.U_Z_UNITTYPE"
                    strsql = strsql & " where " & strUnitCondition & " and " & strPriceCondition & " and " & strSpaceCondition & " order by T1.U_Z_PROPCODE,T1.U_Z_CODE"
                Case "O" 'with out contract
                    strsql = "select U_Z_PROPDESC,U_Z_PROITEMCODE,U_Z_CODE,T1.U_Z_DESC,T0.U_Z_UNITTYPE,case U_Z_UNITSTATUS when 'M' then 'Under Maintenance' when 'O' then 'Offered' when 'R' then 'Reserved (before signature)' when 'L' then 'Reserved (After Signature)' when 'S' then 'Sold' when 'A' then 'Available' else 'Not Available' end,U_Z_SPACE,U_Z_PRICE,U_Z_TOTALAREA,U_Z_TOTALPRICE,U_Z_COSTCENTER from [@Z_PROPUNIT] T1 "
                    strsql = strsql & " inner join [@Z_UnitType] T0 on T0.DocEntry=T1.U_Z_UNITTYPE "
                    strsql = strsql & " where " & strUnitCondition & " and T1.U_Z_PROITEMCODE not in (Select U_Z_UNITCODE from [@Z_CONTRACT] ) and " & strPriceCondition & " and " & strSpaceCondition & " order by T1.U_Z_PROPCODE,T1.U_Z_CODE"
                Case "E" 'expired after the given date 
                    strsql = " select T1.U_Z_PROPDESC,T1.U_Z_PROITEMCODE,T1.U_Z_CODE,T1.U_Z_DESC,T2.U_Z_UNITTYPE,  T1.U_Z_SPACE,T1.U_Z_PRICE,T1.U_Z_TOTALAREA,T1.U_Z_TOTALPRICE,T0.DocEntry,U_Z_TENCODE,U_Z_TENNAme,U_Z_STARTDATE,U_Z_ENDDATE,case T0.U_Z_STATUS when 'PED' then 'Pending' when 'APP' then 'Approved' when 'AGR' then 'Agreed' when 'TER' then 'Terminated' else 'Cancelled' end , U_Z_OFFADDRESS,U_Z_ANNUALRENT from [@Z_CONTRACT] T0"
                    strsql = strsql & " RIGHT outer join [@Z_PROPUNIT] T1 on T1.U_Z_PROITEMCODE=T0.U_Z_UNITCODE  inner join   [@Z_UnitType] T2 on T2.DocEntry=T1.U_Z_UNITTYPE"
                    ' strsql = strsql & " where " & strUnitCondition & " and " & strPriceCondition & " and " & strSpaceCondition & " and (T0.U_Z_ENDDATE is null or ( T0.U_Z_ENDDATE < '" & dtDate.ToString("yyyy-MM-dd") & "' and  T0.U_Z_STARTDATE > '" & dtDate.ToString("yyyy-MM-dd") & "')) order by T1.U_Z_PROPCODE,T1.U_Z_CODE "
                    strsql = strsql & " where " & strUnitCondition & " and " & strPriceCondition & " and " & strSpaceCondition & " and (T0.U_Z_ENDDATE is null or ( T0.U_Z_ENDDATE < '" & dtDate.ToString("yyyy-MM-dd") & "')) order by T1.U_Z_PROPCODE,T1.U_Z_CODE "
            End Select
            oGrid = aForm.Items.Item("20").Specific
            oGrid.DataTable.ExecuteQuery(strsql)
            FormatGrid(strreportChoice, oGrid)
            oGrid.AutoResizeColumns()
            aForm.Items.Item("20").Enabled = False
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_None
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try

    End Sub
#End Region

    Private Sub FormatGrid(ByVal aChoice As String, ByVal aGrid As SAPbouiCOM.Grid)
        Select Case aChoice
            Case "C"
                aGrid.Columns.Item(0).TitleObject.Caption = "Property Description"
                aGrid.Columns.Item(1).TitleObject.Caption = "Unit Code"
                oEditTextColumn = aGrid.Columns.Item(1)
                oEditTextColumn.LinkedObjectType = "Z_PROPUNIT"
                aGrid.Columns.Item(2).TitleObject.Caption = "Unit Code"
                aGrid.Columns.Item(2).Visible = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Unit Description"
                aGrid.Columns.Item(4).TitleObject.Caption = "Unit Type"
                aGrid.Columns.Item(5).TitleObject.Caption = "Space Provided"
                aGrid.Columns.Item(6).TitleObject.Caption = "Sales price per Sqr.Meter"
                aGrid.Columns.Item(7).TitleObject.Caption = "Total Area"
                aGrid.Columns.Item(8).TitleObject.Caption = "Total Price"

                aGrid.Columns.Item(9).TitleObject.Caption = "Contract ID"
                oEditTextColumn = aGrid.Columns.Item(9)
                oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                aGrid.Columns.Item(10).TitleObject.Caption = "Tenent Code"
                oEditTextColumn = aGrid.Columns.Item(10)
                oEditTextColumn.LinkedObjectType = "2"
                aGrid.Columns.Item(11).TitleObject.Caption = "Tenent name"
                aGrid.Columns.Item(12).TitleObject.Caption = "Start Date"
                aGrid.Columns.Item(13).TitleObject.Caption = "End Date"
                aGrid.Columns.Item(14).TitleObject.Caption = "Contract Status"
                aGrid.Columns.Item(15).TitleObject.Caption = "Office Address"
                aGrid.Columns.Item(16).TitleObject.Caption = "Annual Rent"

            Case "O"
                aGrid.Columns.Item(0).TitleObject.Caption = "Property Description"
                aGrid.Columns.Item(1).TitleObject.Caption = "Item Code"
                oEditTextColumn = aGrid.Columns.Item(1)
                oEditTextColumn.LinkedObjectType = "4"
                aGrid.Columns.Item(2).TitleObject.Caption = "Unit Code"
                aGrid.Columns.Item(2).Visible = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Unit Description"

                aGrid.Columns.Item(4).TitleObject.Caption = "Unit Type"
                aGrid.Columns.Item(5).TitleObject.Caption = "Unit Status"
                aGrid.Columns.Item(6).TitleObject.Caption = "Space"
                aGrid.Columns.Item(7).TitleObject.Caption = "Price"
                aGrid.Columns.Item(8).TitleObject.Caption = "Total Area"
                aGrid.Columns.Item(9).TitleObject.Caption = "Total Price"
                aGrid.Columns.Item(10).TitleObject.Caption = "Cost Center"
            Case "E"
                aGrid.Columns.Item(0).TitleObject.Caption = "Property Description"
                aGrid.Columns.Item(1).TitleObject.Caption = "Unit Code"
                oEditTextColumn = aGrid.Columns.Item(1)
                oEditTextColumn.LinkedObjectType = "4"
                aGrid.Columns.Item(2).TitleObject.Caption = "Unit Code"
                aGrid.Columns.Item(2).Visible = False
                aGrid.Columns.Item(3).TitleObject.Caption = "Unit Description"
                aGrid.Columns.Item(4).TitleObject.Caption = "Unit Type"
                aGrid.Columns.Item(5).TitleObject.Caption = "Space Provided"
                aGrid.Columns.Item(6).TitleObject.Caption = "Sales price per Sqr.Meter"
                aGrid.Columns.Item(7).TitleObject.Caption = "Total Area"
                aGrid.Columns.Item(8).TitleObject.Caption = "Total Price"
                aGrid.Columns.Item(9).TitleObject.Caption = "Contract ID"
                oEditTextColumn = aGrid.Columns.Item(9)
                oEditTextColumn.LinkedObjectType = "Z_CONTRACT"
                aGrid.Columns.Item(10).TitleObject.Caption = "Tenent Code"
                oEditTextColumn = aGrid.Columns.Item(10)
                oEditTextColumn.LinkedObjectType = "2"
                aGrid.Columns.Item(11).TitleObject.Caption = "Tenent name"
                aGrid.Columns.Item(12).TitleObject.Caption = "Start Date"
                aGrid.Columns.Item(13).TitleObject.Caption = "End Date"
                aGrid.Columns.Item(14).TitleObject.Caption = "Contract Status"
                aGrid.Columns.Item(15).TitleObject.Caption = "Office Address"
                aGrid.Columns.Item(16).TitleObject.Caption = "Annual Rent"



        End Select

    End Sub

#Region "Events"



#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Search
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
            If pVal.FormTypeEx = frm_Search Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "20" And pVal.ColUID = "U_Z_PROITEMCODE" Then
                                    oGrid = oForm.Items.Item("20").Specific
                                    For intRow As Integer = pVal.Row To pVal.Row
                                        Dim UnitCode As String = oGrid.DataTable.GetValue("U_Z_PROITEMCODE", intRow)
                                        Dim oObjUnit As New clsPropertyUnitDetails
                                        oObjUnit.LoadForm(UnitCode)
                                        BubbleEvent = False
                                        Exit Sub
                                    Next
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "19"
                                        DisplayResult(oForm)
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
