Public Class clsPropertyUnitDetails
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oDBSrc_Line As SAPbouiCOM.DBDataSource
    Private oMatrix As SAPbouiCOM.Matrix
    Private oEditText As SAPbouiCOM.EditText
    Private ostatic As SAPbouiCOM.StaticText
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
    Private strSelectedFilepath, strSelectedFolderPath As String
#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_PropertyUnitDetails) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_PropertyUnitSetup, frm_PropertyUnitDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "4"
        oForm.Freeze(True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' oForm.Items.Item("6").Enabled = False
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu("1287", True)
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.Items.Item("38").DisplayDesc = True
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
    Public Sub LoadForm(ByVal UnitCode As String)
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_PropertyUnitDetails) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_PropertyUnitSetup, frm_PropertyUnitDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.Freeze(True)
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu(mnu_ADD_ROW, True)
        oForm.EnableMenu("1287", True)
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.Items.Item("38").DisplayDesc = True
        oForm.PaneLevel = 1
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        oForm.Items.Item("40").Enabled = True
        oApplication.Utilities.setEdittextvalue(oForm, "40", UnitCode)
        oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oForm.Items.Item("40").Enabled = False
        oForm.Freeze(False)
    End Sub
    Public Sub CreateUnitFromProperty(ByVal aCode As String)
        oForm = oApplication.Utilities.LoadForm(xml_PropertyUnitSetup, frm_PropertyUnitDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "4"
        oForm.Freeze(True)
        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' oForm.Items.Item("6").Enabled = False
        oForm.EnableMenu(mnu_DELETE_ROW, True)
        oForm.EnableMenu("1287", True)
        AddChooseFromList(oForm)
        databind(oForm)
        oForm.Items.Item("38").DisplayDesc = True
        oForm.PaneLevel = 1

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
        oApplication.Utilities.setEdittextvalue(oForm, "4", oApplication.Utilities.getMaxCode("@Z_PROPUNIT", "DocEntry"))
        oApplication.Utilities.setEdittextvalue(oForm, "8", aCode)
        oForm.Items.Item("8").Enabled = True
        oForm.Items.Item("40").Enabled = False
        oForm.Items.Item("39").Enabled = False
        oForm.Items.Item("41").Enabled = True
        oForm.Items.Item("8").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
        oApplication.SBO_Application.SendKeys("{TAB}")
        oApplication.Utilities.setEdittextvalue(oForm, "92", Now.Date)
        PopulatePropetyDetails(oForm, aCode)
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
            oCFLs = objForm.ChooseFromLists
            oCFLCreationParams = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)

            oCFL = oCFLs.Item("CFL_3")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")
            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "CardType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "C"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_6")
            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "DimCode"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "1"
            oCFL.SetConditions(oCons)


            oCFL = oCFLs.Item("CFL_Asset")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "ItemType"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "F"
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
            '            oMatrix = aForm.Items.Item("12").Specific
            oCombo = aForm.Items.Item("38").Specific
            For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception
                End Try
            Next
            oCombo.ValidValues.Add("", "")
            Dim otemp As SAPbobsCOM.Recordset
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("Select DocEntry,U_Z_UnitType from [@Z_UnitType] order by DocEntry")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next

            oCombo = aForm.Items.Item("34").Specific
            For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                Try
                    oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
                Catch ex As Exception

                End Try
            Next
            oCombo.ValidValues.Add("", "")
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("SELECT T0.[ListNum], T0.[ListName] FROM OPLN T0")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oForm.Items.Item("34").DisplayDesc = True
            oCombo = oForm.Items.Item("32").Specific
            Dim intFieldid As Integer
            otemp.DoQuery("SELECT FieldId ,* FROM CUFD where Upper(TableID)='@Z_PROPUNIT' and Upper(AliasID)='Z_UNITSTATUS'")
            intFieldid = otemp.Fields.Item(0).Value
            For intRow As Integer = oCombo.ValidValues.Count - 1 To 0 Step -1
                oCombo.ValidValues.Remove(intRow, SAPbouiCOM.BoSearchKey.psk_Index)
            Next
            otemp.DoQuery("SELECT FldValue,Descr from UFD1 where Upper(tableID)='@Z_PROPUNIT' and FieldID=" & intFieldid)
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            Try
                oCombo.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception

            End Try
            
            'oCombo = aForm.Items.Item("41").Specific
            'oCombo.ValidValues.Add("", "")
            'otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            'otemp.DoQuery("select prcCode,PrcName from OPRC where Locked='N'")
            'For introw As Integer = 0 To otemp.RecordCount - 1
            '    oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
            '    otemp.MoveNext()
            'Next
            'oForm.Items.Item("41").DisplayDesc = True
            '  oForm.Items.Item("11").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
            ' oForm.Items.Item("18").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub


#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "1"
                Exit Sub
                oMatrix = aForm.Items.Item("12").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ1")
            Case "3"
                oMatrix = aForm.Items.Item("86").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT1")
                ' Exit Sub
            Case "5"
                oMatrix = aForm.Items.Item("160").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT2")
            Case "6"
                oMatrix = aForm.Items.Item("163").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT3")
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
                    oMatrix.ClearRowData(oMatrix.RowCount)
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
            Case "5"
                oMatrix = aForm.Items.Item("160").Specific
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
        ' Return True
        If validation(aForm) = False Then
            Return False
        End If
        'AssignLineNo(aForm)
        Return True
    End Function
#Region "Delete Row"
    Private Sub RefereshDeleteRow(ByVal aForm As SAPbouiCOM.Form)
        If 1 = 1 Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT1")
            'Else
            '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PRJ2")
        End If
        If aForm.PaneLevel = 5 Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT2")
            frmSourceMatrix = aForm.Items.Item("160").Specific
        ElseIf aForm.PaneLevel = 6 Then
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT3")
            frmSourceMatrix = aForm.Items.Item("163").Specific
        Else
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT1")
            frmSourceMatrix = aForm.Items.Item("86").Specific
        End If

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
        Dim strCode, strcode1, strProject, strUnitCode, strActivity1, strExtUnitCode As String
        Dim oTemp As SAPbobsCOM.Recordset

        strUnitCode = oApplication.Utilities.getEdittextvalue(aform, "39")
        strExtUnitCode = oApplication.Utilities.getEdittextvalue(aform, "158")
        strProject = oApplication.Utilities.getEdittextvalue(aform, "8")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            getUnitCode(aform)
            If strUnitCode = "" Then
                oApplication.Utilities.Message("Unit code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If strExtUnitCode = "" Then
                oApplication.Utilities.Message("Ext.Unit code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            If strProject = "" Then
                oApplication.Utilities.Message("Project code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oCombo = aform.Items.Item("38").Specific
            If oCombo.Selected.Value = "" Then
                oApplication.Utilities.Message("Property Unit Type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oTemp.DoQuery("Select * from [@Z_PROPUNIT] where U_Z_PropCode='" & strProject & "' and  U_Z_Code='" & strUnitCode & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Unit Code already defined for this project...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            oTemp.DoQuery("Select * from [@Z_PROPUNIT] where U_Z_PropCode='" & strProject & "' and  U_Z_ExtProNo='" & strExtUnitCode & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Ext.Unit Code already defined for this project...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            'oCombo = aform.Items.Item("41").Specific
            'If oCombo.Selected.Value = "" Then
            '    oApplication.Utilities.Message("Cost center is missing..", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If

        End If
        Dim strAnnualRent As String
        Dim dblAnnualRent As Double
        strAnnualRent = oApplication.Utilities.getEdittextvalue(aform, "17")
        If strAnnualRent = "" Then
            '    oApplication.Utilities.Message("Space provide by Sq.Mtr is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '   Return False
        Else
            dblAnnualrent = oApplication.Utilities.getDocumentQuantity(strAnnualRent)
            If dblAnnualrent <= 0 Then
                '      oApplication.Utilities.Message("Space provide by Sq.Mtr should be greater than zero...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '     Return False
            End If
        End If

        Return True
    End Function
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

    Private Sub AssignLineNo(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("25").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROPUNIT1")
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
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim intDocEntry As Integer
        Dim strProject, strProjectName, strItemCode, strUnitCode, strItemname, strPropertyCode As String

        oApplication.Utilities.ExecuteSQL(oTempRS, "Select * from [@Z_PROPUNIT] where docentry=" & strProjectCode)
        intDocEntry = 0
        strProject = ""
        strProjectName = ""
        strItemCode = ""
        Dim strPropCode As String
        If oTempRS.RecordCount > 0 Then
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strUnitCode = oTempRS.Fields.Item("U_Z_Code").Value
            strProject = oTempRS.Fields.Item("U_Z_ExtProNo").Value
            strPropCode = oTempRS.Fields.Item("U_Z_PropCode").Value
            If strProject = "" Then
                strProject = oTempRS.Fields.Item("U_Z_PropCode").Value
            End If
            strProjectName = oTempRS.Fields.Item("U_Z_Desc").Value
            Dim strCostCetner As String = oApplication.Utilities.createCostCenter(strProject, strProjectName, oTempRS.Fields.Item("U_Z_RegDate").Value)
            If strCostCetner <> "" Then
                oApplication.Utilities.ExecuteSQL(oTempRS, "Update ""@Z_PROPUNIT"" set ""U_Z_CostCenter""='" & strCostCetner & "' where isnull(""U_Z_CostCenter"",'')='' and  ""DocEntry""=" & strProjectCode)
            End If
            If strProject = "" Or strUnitCode = "" Then
                Exit Function
            End If
            strItemCode = strPropCode & "-" & strUnitCode.ToString
        Else
            Return True
        End If

        'oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        'oTempRS.DoQuery("Select * from OITM where ItemCode='" & strItemCode & "'")
        'If oTempRS.RecordCount <= 0 Then
        '    If strChoice <> "Delete" Then
        '        strChoice = "Add"
        '    End If
        'End If

        'If strChoice <> "Delete" Then
        '    Dim oCmpSrv As SAPbobsCOM.CompanyService
        '    Dim projectService As SAPbobsCOM.IProjectsService
        '    Dim project As SAPbobsCOM.Items
        '    oCmpSrv = oApplication.Company.GetCompanyService
        '    project = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '    oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    oTempRS.DoQuery("Select * from [@Z_PROPUNIT] where DocEntry=" & intDocEntry)
        '    If project.GetByKey(strItemCode) Then
        '        project.ItemName = oTempRS.Fields.Item("U_Z_PropDesc").Value & "-" & oTempRS.Fields.Item("U_Z_Desc").Value
        '        project.UserFields.Fields.Item("U_Z_PRJCODE").Value = oTempRS.Fields.Item("U_Z_PropCode").Value
        '        project.UserFields.Fields.Item("U_Z_PROFLG").Value = "Y"
        '        project.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
        '        project.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO
        '        project.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO
        '        If project.Update <> 0 Then
        '            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '            Return False
        '        End If
        '    Else
        '        project.ItemCode = strItemCode
        '        project.ItemName = oTempRS.Fields.Item("U_Z_PropDesc").Value & "-" & oTempRS.Fields.Item("U_Z_Desc").Value
        '        project.UserFields.Fields.Item("U_Z_PROFLG").Value = "Y"
        '        project.UserFields.Fields.Item("U_Z_PRJCODE").Value = oTempRS.Fields.Item("U_Z_PropCode").Value
        '        project.SalesItem = SAPbobsCOM.BoYesNoEnum.tYES
        '        project.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO
        '        project.PurchaseItem = SAPbobsCOM.BoYesNoEnum.tNO
        '        If project.Add <> 0 Then
        '            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '            Return False
        '        End If
        '    End If
        'Else
        '    Dim oCmpSrv As SAPbobsCOM.CompanyService
        '    Dim projectService As SAPbobsCOM.IProjectsService
        '    Dim project As SAPbobsCOM.Items
        '    oCmpSrv = oApplication.Company.GetCompanyService
        '    project = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        '    oTempRS.DoQuery("Select * from [@Z_PROPUNIT] where DocEntry=" & intDocEntry)
        '    If project.GetByKey(strItemCode) Then
        '        If project.Remove <> 0 Then
        '            oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '            Return False
        '        End If
        '    End If
        'End If
        oTempRS.DoQuery("Update [@Z_PROPUNIT] set U_Z_ProItemCode='" & strItemCode & "' where DocEntry=" & intDocEntry)
        oTempRS.DoQuery("Select count(*) from [@Z_PROPUNIT] where U_Z_PropCode='" & strProjectCode & "'")
        Dim dblNoofUnit As Double
        dblNoofUnit = oTempRS.Fields.Item(0).Value
        oTempRS.DoQuery("Update [@Z_PROP] set U_Z_NoofUnits='" & dblNoofUnit & "' where U_Z_Code='" & strProjectCode & "'")
        Return True
    End Function
#End Region

    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("86").Specific
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
#End Region

#Region "Events"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        oForm = oApplication.SBO_Application.Forms.Item(eventInfo.FormUID)
        'If eventInfo.FormUID = "RightClk" Then
        If oForm.TypeEx = frm_PropertyUnitDetails Then
            If (eventInfo.BeforeAction = True) Then
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
                    oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Reserve"
                    oCreationPackage.String = "Reservation Details"
                    oCreationPackage.Enabled = True
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)


                    oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Contract"
                    oCreationPackage.String = "Contract Details"
                    oCreationPackage.Enabled = True
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)

                    oCreationPackage = oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
                    oCreationPackage.UniqueID = "Billing"
                    oCreationPackage.String = "Billing Details"
                    oCreationPackage.Enabled = True
                    oMenuItem = oApplication.SBO_Application.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    oMenus.AddEx(oCreationPackage)
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            Else
                Dim oMenuItem As SAPbouiCOM.MenuItem
                Dim oMenus As SAPbouiCOM.Menus
                Try
                    oApplication.SBO_Application.Menus.RemoveEx("Reserve")
                    oApplication.SBO_Application.Menus.RemoveEx("Contract")
                    oApplication.SBO_Application.Menus.RemoveEx("Billing")
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        End If

       

    End Sub


#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_PropertyUnitSetup
                    If pVal.BeforeAction = False Then
                        LoadForm()
                    End If
                Case mnu_ADD_ROW
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        AddRow(oForm)
                    Else
                        If oForm.PaneLevel = 3 Then
                            BubbleEvent = False
                            Exit Sub
                        End If

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
                Case "1287"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "4", oApplication.Utilities.getMaxCode("@Z_PROPUNIT", "DocEntry"))
                        oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("40").Enabled = False
                        oForm.Items.Item("39").Enabled = True
                        oForm.Items.Item("41").Enabled = True
                        oApplication.Utilities.setEdittextvalue(oForm, "40", "")
                        oApplication.Utilities.setEdittextvalue(oForm, "39", "")
                        oApplication.Utilities.setEdittextvalue(oForm, "41", "")
                        getUnitCode(oForm)
                    End If
                Case "Reserve", "Contract", "Billing"
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        Dim strCode As String
                        Dim oTemprs As SAPbobsCOM.Recordset
                        strCode = oApplication.Utilities.getEdittextvalue(oForm, "40")
                        oTemprs = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        ' oTemprs.DoQuery("Select isnull(U_Z_ProFlg,'N') from OITM where ItemCode='" & strCode & "'")
                        If 1 = 1 Then 'oTemprs.Fields.Item(0).Value = "Y" Then
                            Dim objChoose As New clsPropertyUnitReport
                            objChoose.LoadForm(strCode, pVal.MenuUID)
                        End If
                    End If

                Case mnu_Remove
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        Dim strProject As String
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strProject = oApplication.Utilities.getEdittextvalue(oForm, "40")
                        If strProject <> "" Then
                            oTest.DoQuery("Select * from [@Z_RESER] where  U_Z_UnitCode='" & strProject & "'")
                            If oTest.RecordCount > 0 Then
                                oApplication.Utilities.Message("Property  Unit already mapped to Reservation. You can not remove", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            oTest.DoQuery("Select * from [@Z_CONTRACT] where U_Z_UnitCode='" & strProject & "'")
                            If oTest.RecordCount > 0 Then
                                oApplication.Utilities.Message("Property  Unit already mapped to Contract. You can not remove", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If
                            If oApplication.SBO_Application.MessageBox("Do you want to remove this property unit ?", , "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                            Dim strDocNum As String
                            Dim objedittext As SAPbouiCOM.EditText
                            objedittext = oForm.Items.Item("40").Specific
                            strDocNum = objedittext.String
                            'If AddtoItemMaster(strDocNum, "Delete") = False Then
                            '    BubbleEvent = False
                            '    Exit Sub
                            'End If

                        End If
                    End If

                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then

                    End If
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "4", oApplication.Utilities.getMaxCode("@Z_PROPUNIT", "DocEntry"))
                        oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("40").Enabled = False
                        oForm.Items.Item("39").Enabled = False
                        oForm.Items.Item("41").Enabled = True
                        oApplication.Utilities.setEdittextvalue(oForm, "92", Now.Date)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("40").Enabled = False
                        oForm.Items.Item("39").Enabled = True
                        oForm.Items.Item("41").Enabled = True
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
            Try
                If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD Or BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                    Dim strDocNum, strDocType As String
                    Dim objedittext As SAPbouiCOM.EditText
                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    objedittext = oForm.Items.Item("4").Specific
                    strDocNum = objedittext.String
                    Dim otest As SAPbobsCOM.Recordset
                    otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    otest.DoQuery("Select isnull(max(DocEntry),0) from [@Z_PROPUNIT]")
                    strDocNum = otest.Fields.Item(0).Value
                    AddtoItemMaster(strDocNum, "Add")
                End If

                If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                    oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("8").Enabled = False
                        oForm.Items.Item("40").Enabled = False
                        oForm.Items.Item("39").Enabled = False
                        oForm.Items.Item("41").Enabled = False
                    Else
                        oForm.Items.Item("4").Enabled = False
                        oForm.Items.Item("8").Enabled = True
                        oForm.Items.Item("40").Enabled = False
                        oForm.Items.Item("39").Enabled = True
                        oForm.Items.Item("41").Enabled = True
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

    Private Sub getUnitCode(ByVal aform As SAPbouiCOM.Form)
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim strProject As String
        strProject = oApplication.Utilities.getEdittextvalue(aform, "8")
        otest.DoQuery("Select count(*) from [@Z_PROPUNIT] where U_Z_PropCode='" & strProject & "'")
        Dim intCount As Integer
        intCount = otest.Fields.Item(0).Value
        intCount = intCount + 1
        strProject = strProject & "-" & intCount.ToString("00")
        oApplication.Utilities.setEdittextvalue(aform, "39", intCount)
        oApplication.Utilities.setEdittextvalue(aform, "40", strProject)

    End Sub
    Private Sub PopulatePropetyDetails(ByVal aform As SAPbouiCOM.Form, ByVal acode As String)
        Dim otest As SAPbobsCOM.Recordset
        otest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        otest.DoQuery("Select * from [@Z_PROP] where U_Z_Code='" & acode & "'")
        If otest.RecordCount > 0 Then
            oApplication.Utilities.setEdittextvalue(oForm, "edComm", otest.Fields.Item("U_Z_COMM").Value)
            oApplication.Utilities.setEdittextvalue(oForm, "10", otest.Fields.Item("U_Z_Desc").Value)
            getUnitCode(aform)
        End If
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PropertyUnitDetails Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                End If
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                If pVal.ItemUID = "46" And pVal.CharPressed <> 9 Then
                                    ' BubbleEvent = False
                                    ' Exit Sub
                                End If


                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                                oForm = oApplication.SBO_Application.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)
                                oMatrix = oForm.Items.Item("86").Specific
                                If (pVal.ItemUID = "86" Or pVal.ItemUID = "160" Or pVal.ItemUID = "163") And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = pVal.ItemUID
                                    frmSourceMatrix = oMatrix
                                End If

                                If pVal.ItemUID = "46" Then
                                    '  BubbleEvent = False
                                    '  Exit Sub
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "152"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            Dim objCls As New clsContracts
                                            objCls.LoadForm_Contract(oApplication.Utilities.getEdittextvalue(oForm, "4"))
                                        End If
                                    Case "153"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            Dim objCls As New clsTenContracts
                                            objCls.LoadForm_Contract(oApplication.Utilities.getEdittextvalue(oForm, "4"))
                                        End If
                                    Case "43"
                                        oForm.PaneLevel = 1
                                    Case "44"
                                        oForm.PaneLevel = 2
                                    Case "85"
                                        oForm.PaneLevel = 3
                                    Case "124"
                                        oForm.PaneLevel = 4
                                    Case "159"
                                        oForm.PaneLevel = 5
                                    Case "162"
                                        oForm.PaneLevel = 6

                                    Case "161"
                                        AddRow(oForm)
                                        'oMatrix = oForm.Items.Item("160").Specific
                                        'If oMatrix.RowCount <= 0 Then
                                        '    oMatrix.AddRow()
                                        'End If
                                        'If oApplication.Utilities.getMatrixValues(oMatrix, "V_0", oMatrix.RowCount) <> "" Then
                                        '    oMatrix.AddRow()
                                        '    oMatrix.ClearRowData(oMatrix.RowCount)

                                        '    oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                        'End If
                                        'oMatrix = oForm.Items.Item("160").Specific
                                        'oApplication.Utilities.AssignRowNo(oMatrix, oForm)
                                    Case "87"
                                        fillopen()
                                        oMatrix = oForm.Items.Item("86").Specific
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
                                            oForm.Items.Item("26").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
                                    Case "88"
                                        LoadFiles(oForm)
                                    Case "89"
                                        oMatrix = oForm.Items.Item("86").Specific
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
                                End Select

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "17" Or pVal.ItemUID = "24" And pVal.CharPressed = 9 Then


                                End If
                            Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                                Dim oCFLEvento As SAPbouiCOM.IChooseFromListEvent
                                Dim oCFL As SAPbouiCOM.ChooseFromList
                                Dim val1 As String
                                Dim sCHFL_ID, val As String
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
                                        If pVal.ItemUID = "95" Then
                                            val = oDataTable.GetValue("empID", 0)
                                            val1 = oDataTable.GetValue("firstName", 0).ToString & " " & oDataTable.GetValue("middleName", 0).ToString & " " & oDataTable.GetValue("lastName", 0).ToString
                                            oApplication.Utilities.setEdittextvalue(oForm, "154", val1)
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception

                                            End Try

                                        End If


                                        If pVal.ItemUID = "160" And pVal.ColUID = "V_0" Then
                                            val = oDataTable.GetValue("U_Z_CODE", 0)
                                            val1 = oDataTable.GetValue("U_Z_DESC", 0)
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_2", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        End If

                                        If pVal.ItemUID = "163" And pVal.ColUID = "V_0" Then
                                            val = oDataTable.GetValue("ItemCode", 0)
                                            val1 = oDataTable.GetValue("ItemName", 0)
                                            oMatrix = oForm.Items.Item(pVal.ItemUID).Specific
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", pVal.Row, val1)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", pVal.Row, val)
                                        End If
                                        If pVal.ItemUID = "8" Then
                                            val = oDataTable.GetValue("U_Z_CODE", 0)
                                            val1 = oDataTable.GetValue("U_Z_DESC", 0)
                                            Dim dblValue As Double
                                            Try
                                                dblValue = oDataTable.GetValue("U_Z_COMM", 0)
                                            Catch ex As Exception
                                                dblValue = oDataTable.GetValue("U_Z_Comm", 0)
                                            End Try

                                            oApplication.Utilities.setEdittextvalue(oForm, "10", val1)
                                            Try
                                                '   oForm.Items.Item("edComm").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                                oApplication.Utilities.setEdittextvalue(oForm, "edComm", dblValue)
                                            Catch ex As Exception
                                            End Try
                                            Try
                                                val1 = ""
                                                val1 = oDataTable.GetValue("U_Z_CARDCODE", 0)
                                                ' oApplication.Utilities.setEdittextvalue(oForm, "46", val1)
                                            Catch ex As Exception
                                            End Try

                                            Try
                                                val1 = ""
                                                val1 = oDataTable.GetValue("U_Z_CARDNAME", 0)
                                                ' oApplication.Utilities.setEdittextvalue(oForm, "47", val1)
                                            Catch ex As Exception
                                            End Try
                                            Try
                                                oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                            Catch ex As Exception
                                                getUnitCode(oForm)
                                            End Try
                                            getUnitCode(oForm)
                                        ElseIf pVal.ItemUID = "46" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "47", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "105" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "41" Then
                                            val = oDataTable.GetValue("PrcCode", 0)
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
