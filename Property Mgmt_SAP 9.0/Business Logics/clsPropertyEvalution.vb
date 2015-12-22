Public Class clsPropertyEvalution
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
    Private strSelectedFilepath, strSelectedFolderPath As String
    Dim oDataSrc_Line As SAPbouiCOM.DBDataSource
    Dim oDataSrc_Line1 As SAPbouiCOM.DBDataSource

#Region "Methods"

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub
    Private Sub LoadForm()
        If oApplication.Utilities.validateAuthorization(oApplication.Company.UserSignature, frm_Evaluation) = False Then
            oApplication.Utilities.Message("You are not authorized to do this action", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Exit Sub
        End If
        oForm = oApplication.Utilities.LoadForm(xml_Evalution, frm_Evaluation)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "6"

        oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
        ' oForm.Items.Item("6").Enabled = False
        AddChooseFromList(oForm)
        'databind(oForm)
        ' oForm.Items.Item("38").DisplayDesc = True
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub
    Public Sub LoadForm(ByVal aCode As String, ByVal aName As String)
        Dim strUnitCode As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strUnitCode = aCode
        oTemp.DoQuery("Select * from [@Z_AQEVAL] where U_Z_ProCode='" & strUnitCode & "'")
        oForm = oApplication.Utilities.LoadForm(xml_Evalution, frm_Evaluation)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "6"
        oForm.Freeze(True)
        If oTemp.RecordCount > 0 Then
            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            ''  oApplication.Utilities.setEdittextvalue(oForm, "6", oTemp.Fields.Item("DocEntry").Value)
            'oApplication.Utilities.setEdittextvalue(oForm, "115", aCode)
            'oApplication.Utilities.setEdittextvalue(oForm, "116", aName)
            'oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getMaxCode("@Z_AQEVAL", "DocEntry"))
            oApplication.Utilities.setEdittextvalue(oForm, "115", aCode)
            oApplication.Utilities.setEdittextvalue(oForm, "116", aName)
            oForm.Items.Item("115").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.SBO_Application.SendKeys("{TAB}")
        Else
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getMaxCode("@Z_AQEVAL", "DocEntry"))
            oApplication.Utilities.setEdittextvalue(oForm, "115", aCode)
            oApplication.Utilities.setEdittextvalue(oForm, "116", aName)
            oForm.Items.Item("115").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.SBO_Application.SendKeys("{TAB}")
        End If
        AddChooseFromList(oForm)
        oForm.PaneLevel = 1
        oForm.Freeze(False)
    End Sub

    Public Sub LoadForm_View(ByVal aCode As String, ByVal aName As String, ByVal aPrcCode As String)
        Dim strUnitCode As String
        Dim oTemp As SAPbobsCOM.Recordset
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        strUnitCode = aCode
        oTemp.DoQuery("Select * from [@Z_AQEVAL] where DocEntry='" & strUnitCode & "'")
        oForm = oApplication.Utilities.LoadForm(xml_Evalution, frm_Evaluation)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        oForm.DataBrowser.BrowseBy = "6"
        oForm.Freeze(True)
        AddChooseFromList(oForm)
        If oTemp.RecordCount > 0 Then
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            '  oApplication.Utilities.setEdittextvalue(oForm, "6", oTemp.Fields.Item("DocEntry").Value)
            oForm.Items.Item("6").Enabled = True
            oApplication.Utilities.setEdittextvalue(oForm, "6", aCode)

            'oApplication.Utilities.setEdittextvalue(oForm, "116", aName)

            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            'oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            'oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getMaxCode("@Z_AQEVAL", "DocEntry"))
            'oApplication.Utilities.setEdittextvalue(oForm, "115", aCode)
            'oApplication.Utilities.setEdittextvalue(oForm, "116", aName)
            'oForm.Items.Item("115").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            'oApplication.SBO_Application.SendKeys("{TAB}")
        Else
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getMaxCode("@Z_AQEVAL", "DocEntry"))
            oApplication.Utilities.setEdittextvalue(oForm, "115", aPrcCode)
            oApplication.Utilities.setEdittextvalue(oForm, "116", aName)
            oForm.Items.Item("115").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
            oApplication.SBO_Application.SendKeys("{TAB}")
        End If

        oForm.PaneLevel = 1
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
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
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
            oCombo.Select("A", SAPbouiCOM.BoSearchKey.psk_ByValue)

            oCombo = aForm.Items.Item("41").Specific
            oCombo.ValidValues.Add("", "")
            otemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            otemp.DoQuery("select prcCode,PrcName from OPRC where Locked='N'")
            For introw As Integer = 0 To otemp.RecordCount - 1
                oCombo.ValidValues.Add(otemp.Fields.Item(0).Value, otemp.Fields.Item(1).Value)
                otemp.MoveNext()
            Next
            oForm.Items.Item("41").DisplayDesc = True
            oForm.Items.Item("11").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
            oForm.Items.Item("18").TextStyle = SAPbouiCOM.BoTextStyle.ts_UNDERLINE
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub


#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "3"
                oMatrix = aForm.Items.Item("119").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AQEVAL1")
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
                        Case "3"
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
    'Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
    '    Select Case aForm.PaneLevel
    '        Case "1"
    '            oMatrix = aForm.Items.Item("25").Specific
    '            'Case "2"
    '            '    oMatrix = aForm.Items.Item("13").Specific
    '    End Select
    '    For introw As Integer = 1 To oMatrix.RowCount
    '        If oMatrix.IsRowSelected(introw) Then
    '            oMatrix.DeleteRow(introw)
    '        End If
    '    Next

    'End Sub

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
        'If aForm.PaneLevel = 3 Then
        '    oMatrix = aForm.Items.Item("119").Specific
        '    oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AQEVAL1")
        'End If
        'oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_AQEVAL1")
        'intSelectedMatrixrow = -1
        'For introw As Integer = 0 To oMatrix.RowCount
        '    If oMatrix.IsRowSelected(introw) Then
        '        intSelectedMatrixrow = introw
        '        Exit For
        '    End If
        'Next
        'If intSelectedMatrixrow <= 0 Then
        '    Exit Sub
        'End If
        'Me.RowtoDelete = intSelectedMatrixrow
        'oDataSrc_Line.RemoveRecord(Me.RowtoDelete - 1)
        'oMatrix = oMatrix
        'oMatrix.FlushToDataSource()
        'For count = 1 To oDataSrc_Line.Size - 1
        '    oDataSrc_Line.SetValue("LineId", count - 1, count)
        'Next
        'oMatrix.LoadFromDataSource()
        'If oMatrix.RowCount > 0 Then
        '    oMatrix.DeleteRow(oMatrix.RowCount)
        'End If
    End Sub
#End Region
#End Region

#End Region

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strUnitCode, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset

        strUnitCode = oApplication.Utilities.getEdittextvalue(aform, "115")
        ' strProject = oApplication.Utilities.getEdittextvalue(aform, "8")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If strUnitCode = "" Then
                oApplication.Utilities.Message("Property Code code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            'If strProject = "" Then
            '    oApplication.Utilities.Message("Project code is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            'oCombo = aform.Items.Item("38").Specific
            'If oCombo.Selected.Value = "" Then
            '    oApplication.Utilities.Message("Property Unit Type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            '    Return False
            'End If
            oTemp.DoQuery("Select * from [@Z_AQEVAL] where U_Z_ProCode='" & strUnitCode & "'")
            If oTemp.RecordCount > 0 Then
                'oApplication.Utilities.Message("Evalution already already exists for this Property code ...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'Return False
            End If
        End If
        Return True
    End Function
#End Region
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
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim intDocEntry As Integer
        Dim strProject, strProjectName, strItemCode, strUnitCode, strItemname, strPropertyCode As String

        oApplication.Utilities.ExecuteSQL(oTempRS, "Select * from [@Z_PROPUNIT] where docentry=" & strProjectCode)
        intDocEntry = 0
        strProject = ""
        strProjectName = ""
        strItemCode = ""
        If oTempRS.RecordCount > 0 Then
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strUnitCode = oTempRS.Fields.Item("U_Z_Code").Value
            strProjectCode = oTempRS.Fields.Item("U_Z_PropCode").Value
            strProject = oTempRS.Fields.Item("U_Z_PropCode").Value
            strProjectName = oTempRS.Fields.Item("U_Z_Desc").Value
            If strProject = "" Or strUnitCode = "" Then
                Exit Function
            End If
            strItemCode = strProject & "-" & strUnitCode.ToString
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

    'Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)
    '    Select Case aForm.PaneLevel
    '        Case "7"
    '            oMatrix = aForm.Items.Item("175").Specific
    '            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_RDR3")
    '    End Select
    '    Try
    '        aForm.Freeze(True)
    '        oMatrix = aForm.Items.Item("175").Specific
    '        oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_RDR3")
    '        If oMatrix.RowCount <= 0 Then
    '            oMatrix.AddRow()
    '        End If
    '        oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
    '        Try
    '            If oEditText.Value <> "" Then
    '                oMatrix.AddRow()
    '                Select Case aForm.PaneLevel
    '                    Case "7"
    '                        oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, "")
    '                End Select
    '            End If

    '        Catch ex As Exception
    '            aForm.Freeze(False)
    '            oMatrix.AddRow()
    '        End Try

    '        oMatrix.FlushToDataSource()
    '        For count As Integer = 1 To oDataSrc_Line.Size
    '            oDataSrc_Line.SetValue("LineId", count - 1, count)
    '        Next
    '        oMatrix.LoadFromDataSource()

    '        aForm.Freeze(False)
    '    Catch ex As Exception
    '        oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        aForm.Freeze(False)

    '    End Try
    'End Sub
    Private Sub deleterow(ByVal aForm As SAPbouiCOM.Form)
        Select Case aForm.PaneLevel
            Case "3"
                oMatrix = aForm.Items.Item("119").Specific
                oDataSrc_Line = aForm.DataSources.DBDataSources.Item("@Z_AQEVAL1")
        End Select
        oMatrix.FlushToDataSource()
        For introw As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(introw) Then
                oMatrix.DeleteRow(introw)
                oDataSrc_Line.RemoveRecord(introw - 1)
                'oMatrix = frmSourceMatrix

                For count As Integer = 1 To oDataSrc_Line.Size
                    oDataSrc_Line.SetValue("LineId", count - 1, count)
                Next
                oMatrix.LoadFromDataSource()
                If aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                    aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                End If
                Exit Sub
            End If
        Next


    End Sub
    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("119").Specific
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

   


#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_Evaluation Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "1" And (oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE) Then
                                    oApplication.Utilities.setEdittextvalue(oForm, "96", oApplication.Company.UserName)
                                    If AddtoUDT(oForm) = False Then
                                        BubbleEvent = False
                                        Exit Sub
                                    End If
                                ElseIf pVal.ItemUID = "1" And oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                    oApplication.Utilities.setEdittextvalue(oForm, "96", oApplication.Company.UserName)
                                End If
                        End Select

                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'oItem = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
                                Try
                                    If oForm.TypeEx = frm_Evaluation Then
                                        oForm.Items.Item("1000007").Width = oForm.Items.Item("24").Left + oForm.Items.Item("24").Width + 5
                                    End If
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "19"
                                        oForm.PaneLevel = 1
                                    Case "20"
                                        oForm.PaneLevel = 2
                                    Case "118"
                                        oForm.PaneLevel = 3
                                    Case "117"
                                        Dim obj As New clsPrint
                                        obj.PrintEvalution(oApplication.Utilities.getEdittextvalue(oForm, "6"))
                                    Case "126"
                                        Dim obj As New clsPrint
                                        obj.PrintEvalution_Arabic(oApplication.Utilities.getEdittextvalue(oForm, "6"))

                                    Case "122"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub
                                        End If
                                        deleterow(oForm)
                                        RefereshDeleteRow(oForm)
                                    Case "121"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub

                                        End If
                                        LoadFiles(oForm)
                                    Case "120"
                                        If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_OK_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                            Exit Sub

                                        End If
                                        fillopen()
                                        oMatrix = oForm.Items.Item("119").Specific
                                        AddRow(oForm)
                                        Try
                                            oForm.Freeze(True)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_0", oMatrix.RowCount, strSelectedFilepath)
                                            'Dim strDate As String
                                            'Dim dtdate As Date
                                            'dtdate = Now.Date
                                            'strDate = Now.Date.Today().ToString
                                            ' ''  strdate=
                                            ''Dim oColumn As SAPbouiCOM.Column
                                            ''oColumn = oMatrix.Columns.Item("V_1")
                                            ' '' oColumn.Editable = True

                                            ''oEditText = oMatrix.Columns.Item("V_1").Cells.Item(oMatrix.RowCount).Specific
                                            ''oEditText.String = "t"
                                            ''oApplication.SBO_Application.SendKeys("{TAB}")
                                            ''oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            ''oColumn.Editable = False
                                            '' oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, dtdate)
                                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                                            End If
                                            oForm.Freeze(False)
                                        Catch ex As Exception
                                            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                            oForm.Freeze(False)

                                        End Try
                                End Select
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
                                        If pVal.ItemUID = "115" Then
                                            val = oDataTable.GetValue("U_Z_CODE", 0)
                                            val1 = oDataTable.GetValue("U_Z_DESC", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "116", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "124", oDataTable.GetValue("U_Z_CARDCODE", 0))
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        ElseIf pVal.ItemUID = "124" Then
                                            val = oDataTable.GetValue("CardCode", 0)
                                            val1 = oDataTable.GetValue("CardName", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, "10", val1)
                                            oApplication.Utilities.setEdittextvalue(oForm, "12", oDataTable.GetValue("Phone1", 0))
                                            oApplication.Utilities.setEdittextvalue(oForm, "14", oDataTable.GetValue("Cellular", 0))
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
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
#End Region

#Region "Menu Event"
    Public Overrides Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.MenuUID
                Case mnu_Evalution
                    LoadForm()
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
                Case mnu_ADD
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oApplication.Utilities.setEdittextvalue(oForm, "6", oApplication.Utilities.getMaxCode("@Z_AQEVAL", "DocEntry"))
                        oForm.Items.Item("6").Enabled = False
                        oForm.Items.Item("115").Enabled = True
                        '  oform.Items.Item("116").Enabled=True 
                        oApplication.Utilities.setEdittextvalue(oForm, "96", oApplication.Company.UserName)
                    End If
                Case mnu_FIND
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = False Then
                        oForm.Items.Item("4").Enabled = True
                        oForm.Items.Item("115").Enabled = True
                    End If
            End Select
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
    End Sub
#End Region

    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try
            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
            End If

            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    oForm.Items.Item("6").Enabled = False
                    oForm.Items.Item("115").Enabled = False
                Else
                    oForm.Items.Item("6").Enabled = False
                    oForm.Items.Item("115").Enabled = True
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
