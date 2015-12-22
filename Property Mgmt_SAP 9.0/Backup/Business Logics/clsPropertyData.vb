Public Class clsPropertyData
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
    Private InvBaseDocNo, sPath, strSelectedFilepath, strSelectedFolderPath As String
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
        oForm = oApplication.Utilities.LoadForm(xml_PropertyData, frm_PropertyData)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            
            oForm.DataBrowser.BrowseBy = "5"
            oForm.EnableMenu(mnu_DELETE_ROW, True)
            'oForm.PaneLevel = 1
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
            For count = 1 To oDataSrc_Line.Size - 1
                oDataSrc_Line.SetValue("LineId", count - 1, count)
            Next
            oMatrix = oForm.Items.Item("25").Specific
            oMatrix.AutoResizeColumns()
            databind(oForm)
            AddChooseFromList(oForm)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
            oForm.Items.Item("19").DisplayDesc = True
            oForm.Items.Item("30").DisplayDesc = True
            oForm.PaneLevel = 1
            oForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            oForm.Freeze(False)
        End Try
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

            oCFL = oCFLs.Item("CFL_2")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_3")

            ' Adding Conditions to CFL1
            oCons = oCFL.GetConditions()

            oCon = oCons.Add()
            oCon.Alias = "Postable"
            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCon.CondVal = "Y"
            oCFL.SetConditions(oCons)

            oCFL = oCFLs.Item("CFL_4")

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
            oCombo = aForm.Items.Item("15").Specific
            oApplication.Utilities.FillComboBox(oCombo, "Select Code,Name from [@Z_OLOC] order by Code ")
            aForm.Items.Item("15").DisplayDesc = True
            oCombo = aForm.Items.Item("30").Specific
            oApplication.Utilities.FillComboBox(oCombo, "Select U_Z_CODE,U_Z_NAME from [@Z_OPROTYPE] order by Convert(numeric,Code)")
            aForm.Items.Item("30").DisplayDesc = True
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
                oMatrix = aForm.Items.Item("25").Specific
                oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
        End Select
        Try
            aForm.Freeze(True)
            oMatrix = aForm.Items.Item("25").Specific
            oDataSrc_Line = oForm.DataSources.DBDataSources.Item("@Z_PROP1")
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

    Private Sub LoadFiles(ByVal aform As SAPbouiCOM.Form)
        oMatrix = aform.Items.Item("25").Specific
        For intRow As Integer = 1 To oMatrix.RowCount
            If oMatrix.IsRowSelected(intRow) Then
                Dim strFilename As String
                strFilename = oMatrix.Columns.Item("V_1").Cells.Item(intRow).Specific.value
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
    Public Function AddtoUDT(ByVal aForm As SAPbouiCOM.Form) As Boolean
        Dim oTemp As SAPbobsCOM.Recordset
        Dim strDocEntry As String
        '  Return True
        If validation(aForm) = False Then
            Return False
        End If
        AssignLineNo(aForm)
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
            If aForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        End If
        
    End Sub
#End Region
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

#Region "validations"
    Private Function validation(ByVal aform As SAPbouiCOM.Form) As Boolean
        Dim strCode, strcode1, strProject, strActivity, strActivity1 As String
        Dim oTemp As SAPbobsCOM.Recordset

        strProject = oApplication.Utilities.getEdittextvalue(aform, "4")
        oTemp = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        If aform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
            If strProject = "" Then
                oApplication.Utilities.Message("Project code missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            If oApplication.Utilities.getEdittextvalue(aform, "32") = """" Then
                oApplication.Utilities.Message("Project Name missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oCombo = aform.Items.Item("30").Specific
            If oCombo.Selected.Value = "" Then
                oApplication.Utilities.Message("Property type is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oCombo = aform.Items.Item("15").Specific
            If oCombo.Selected.Value = "" Then
                oApplication.Utilities.Message("Location is missing...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            oTemp.DoQuery("Select * from [@Z_PROP] where U_Z_Code='" & strProject & "'")
            If oTemp.RecordCount > 0 Then
                oApplication.Utilities.Message("Property code already exists...", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
        End If

        Dim strAnnualRent As String
        Dim dblAnnualRent As Double
        strAnnualRent = oApplication.Utilities.getEdittextvalue(aform, "29")
        If strAnnualRent = "" Then
            ' oApplication.Utilities.Message("Account Code is missing....", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ' Return False
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

#Region "Update / Add Project Details in SAP"
    Private Function ProjectDetailstoSAP(ByVal strProjectCode As String, ByVal strChoice As String) As Boolean
        Dim oTempRS As SAPbobsCOM.Recordset
        Dim intDocEntry As Integer
        Dim strProject, strProjectName As String
        oApplication.Utilities.ExecuteSQL(oTempRS, "Select * from [@Z_PROP] where U_Z_Code='" & strProjectCode & "'")
        intDocEntry = 0
        strProject = ""
        strProjectName = ""
        If oTempRS.RecordCount > 0 Then
            intDocEntry = oTempRS.Fields.Item("DocEntry").Value
            strProject = oTempRS.Fields.Item("U_Z_Code").Value
            strProjectName = oTempRS.Fields.Item("U_Z_Desc").Value
        End If

        oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oTempRS.DoQuery("Select * from OPRJ where PrjCode='" & strProject & "'")
        '  MsgBox(oTempRS.RecordCount)
        If oTempRS.RecordCount <= 0 Then
            If strChoice <> "Delete" Then
                strChoice = "Add"
            End If
        End If
        If strChoice = "Add" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            oCmpSrv = oApplication.Company.GetCompanyService
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery("Select * from [@Z_PROP] where DocEntry=" & intDocEntry)

            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            project = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProject)
            project.Code = strProject
            project.Name = strProjectName
            projectService.AddProject(project)
        ElseIf strChoice = "Update" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            Dim projectParams As SAPbobsCOM.IProjectParams
            oTempRS = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oTempRS.DoQuery("Select * from [@Z_PROP] where DocEntry=" & intDocEntry)
            oCmpSrv = oApplication.Company.GetCompanyService
            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            'Get a project
            projectParams = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)
            projectParams.Code = strProject
            project = projectService.GetProject(projectParams)
            'Update the project
            project.Name = strProjectName
            projectService.UpdateProject(project)
        ElseIf strChoice = "Delete" Then
            Dim oCmpSrv As SAPbobsCOM.CompanyService
            Dim projectService As SAPbobsCOM.IProjectsService
            Dim project As SAPbobsCOM.IProject
            Dim projectParams As SAPbobsCOM.IProjectParams
            oCmpSrv = oApplication.Company.GetCompanyService
            projectService = oCmpSrv.GetBusinessService(SAPbobsCOM.ServiceTypes.ProjectsService)
            'Get a project
            projectParams = projectService.GetDataInterface(SAPbobsCOM.ProjectsServiceDataInterfaces.psProjectParams)
            projectParams.Code = strProject
            project = projectService.GetProject(projectParams)
            'delete the project
            Try
                projectService.DeleteProject(projectParams)
            Catch ex As Exception
                oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End Try
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
                Case mnu_Propertydata
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
                        Dim intNextnumber As Integer
                        Dim projectCode As String
                        intNextnumber = oApplication.Utilities.getMaxCode("@Z_PROP", "DocEntry")
                        projectCode = "P" & intNextnumber.ToString("000")


                        oApplication.Utilities.setEdittextvalue(oForm, "5", oApplication.Utilities.getMaxCode("@Z_PROP", "DocEntry"))
                        oApplication.Utilities.setEdittextvalue(oForm, "4", projectCode)
                        oForm.Items.Item("5").Enabled = False
                        oForm.Items.Item("4").Enabled = False

                        'Dim oTest As SAPbobsCOM.Recordset
                        'oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'oTest.DoQuery("select CompnyName ,CompnyAddr,Phone1,Phone2,Fax,E_Mail  from OADM")
                        'Dim str As String
                        'str = oTest.Fields.Item(0).Value & "," & oTest.Fields.Item(1).Value & "," & oTest.Fields.Item(2).Value & "," & oTest.Fields.Item(3).Value & "," & oTest.Fields.Item(4).Value & "," & oTest.Fields.Item(5).Value
                        'oApplication.Utilities.setEdittextvalue(oForm, "89", oTest.Fields.Item(0).Value)
                        'oApplication.Utilities.setEdittextvalue(oForm, "91", str)
                    End If

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

                Case mnu_Remove
                    oForm = oApplication.SBO_Application.Forms.ActiveForm()
                    If pVal.BeforeAction = True Then
                        Dim strProject As String
                        Dim oTest As SAPbobsCOM.Recordset
                        oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        strProject = oApplication.Utilities.getEdittextvalue(oForm, "4")
                        If strProject <> "" Then
                            oTest.DoQuery("Select * from [@Z_PROPUNIT] where  U_Z_PropCode='" & strProject & "'")
                            If oTest.RecordCount > 0 Then
                                oApplication.Utilities.Message("Property already mapped to Property Unit. You can not remove", SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                BubbleEvent = False
                                Exit Sub
                            End If

                            If oApplication.SBO_Application.MessageBox("Do you want to remove this property ?", , "Yes", "No") = 2 Then
                                BubbleEvent = False
                                Exit Sub
                            End If

                            Dim strDocNum As String
                            Dim objedittext As SAPbouiCOM.EditText
                            objedittext = oForm.Items.Item("4").Specific
                            strDocNum = objedittext.String
                            If ProjectDetailstoSAP(strDocNum, "Delete") = False Then
                                BubbleEvent = False
                                Exit Sub
                            End If

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


    Public Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        Try

            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    oGrid = oForm.Items.Item("100").Specific
                    oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry] 'Evaluation ID', T0.[U_Z_PROCODE] 'Project Code', T0.[U_Z_PRONAME] 'Project Name', T0.[U_Z_EVL_ID] 'Evaluation Serial No', T0.[U_Z_EV_DATE] 'Evaluation Date', T0.[U_Z_OWNER] 'Owner' FROM [dbo].[@Z_AQEVAL]  T0 where U_Z_ProCode='" & oApplication.Utilities.getEdittextvalue(oForm, "4") & "'")
                    oEditTextColumn = oGrid.Columns.Item(0)
                    oEditTextColumn.LinkedObjectType = "@Z_AQEVAL"
                    oGrid.AutoResizeColumns()
                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD) Then
                Dim strDocNum, strDocType As String
                Dim objedittext As SAPbouiCOM.EditText
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                objedittext = oForm.Items.Item("4").Specific
                strDocNum = objedittext.String
                ProjectDetailstoSAP(strDocNum, "Add")
            End If

            If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD Then
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    oForm.Items.Item("5").Enabled = False
                    oForm.Items.Item("4").Enabled = False 'oForm.Items.Item("71").Enabled = False
                    oGrid = oForm.Items.Item("100").Specific
                    oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry] 'Evaluation ID', T0.[U_Z_PROCODE] 'Project Code', T0.[U_Z_PRONAME] 'Project Name', T0.[U_Z_EVL_ID] 'Evaluation Serial No', T0.[U_Z_EV_DATE] 'Evaluation Date', T0.[U_Z_OWNER] 'Owner' FROM [dbo].[@Z_AQEVAL]  T0 where U_Z_ProCode='" & oApplication.Utilities.getEdittextvalue(oForm, "4") & "'")
                    oEditTextColumn = oGrid.Columns.Item(0)
                    oEditTextColumn.LinkedObjectType = "Z_AQEVAL"
                    oGrid.AutoResizeColumns()

                Else
                    oForm.Items.Item("5").Enabled = True
                    oForm.Items.Item("4").Enabled = True
                End If
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '  oApplication.Utilities.setEdittextvalue(oForm, "89", oApplication.Company.CompanyName)
                    Dim oTest As SAPbobsCOM.Recordset
                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    oTest.DoQuery("select CompnyName ,CompnyAddr,Phone1,Phone2,Fax,E_Mail  from OADM")
                    Dim str As String
                    str = oTest.Fields.Item(0).Value & "," & oTest.Fields.Item(1).Value & "," & oTest.Fields.Item(2).Value & "," & oTest.Fields.Item(3).Value & "," & oTest.Fields.Item(4).Value & "," & oTest.Fields.Item(5).Value
                    oApplication.Utilities.setEdittextvalue(oForm, "89", oTest.Fields.Item(0).Value)
                    oApplication.Utilities.setEdittextvalue(oForm, "91", str)
                    If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        '   oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                    End If


                End If
            End If
            If BusinessObjectInfo.BeforeAction = False And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE) Then
                Dim strDocNum, strDocType As String
                Dim objedittext As SAPbouiCOM.EditText
                oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                objedittext = oForm.Items.Item("4").Specific
                strDocNum = objedittext.String
                ProjectDetailstoSAP(strDocNum, "Update")
            End If

            If BusinessObjectInfo.BeforeAction = True And (BusinessObjectInfo.EventType = SAPbouiCOM.BoEventTypes.et_FORM_DATA_DELETE) Then
                'Dim strDocNum, strDocType As String
                'Dim objedittext As SAPbouiCOM.EditText
                'oForm = oApplication.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID)
                'objedittext = oForm.Items.Item("5").Specific
                'strDocNum = objedittext.String
                'If ProjectDetailstoSAP(strDocNum, "Delete") = False Then
                '    BubbleEvent = False
                '    Exit Sub
                'End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_PropertyData Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
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
                                oMatrix = oForm.Items.Item("25").Specific
                                If pVal.ItemUID = "25" And pVal.Row > 0 Then
                                    Me.RowtoDelete = pVal.Row
                                    intSelectedMatrixrow = pVal.Row
                                    Me.MatrixId = "12"
                                    frmSourceMatrix = oMatrix
                                End If
                        End Select
                    Case False
                        Select Case pVal.EventType

                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                If pVal.ItemUID = "93" And pVal.CharPressed = 9 Then
                                    Dim oTest As SAPbobsCOM.Recordset
                                    oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                    If oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) <> "" Then
                                        oTest.DoQuery("Select * from OCRD where CardCode='" & oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) & "'")
                                        If oTest.RecordCount <= 0 Then
                                            Exit Sub
                                        End If
                                    Else
                                        Exit Sub
                                    End If
                                    oTest.DoQuery("select CntctPrsn , MailAddres ,Phone1,isnull(Phone2,'') + ',' + isnull(Cellular,''), Fax,E_Mail ,* from OCRD where CardCode='" & oApplication.Utilities.getEdittextvalue(oForm, pVal.ItemUID) & "'")
                                    Dim str As String
                                    str = oTest.Fields.Item(0).Value & "," & oTest.Fields.Item(1).Value & "," & oTest.Fields.Item(2).Value & "," & oTest.Fields.Item(3).Value & "," & oTest.Fields.Item(4).Value & "," & oTest.Fields.Item(5).Value
                                    Try
                                        oApplication.Utilities.setEdittextvalue(oForm, "96", str)
                                    Catch ex As Exception
                                    End Try

                                End If
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                'oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Try
                                    oForm.Items.Item("33").Width = oForm.Items.Item("26").Left + oForm.Items.Item("26").Width + 5
                                Catch ex As Exception

                                End Try

                                Try
                                    '  oForm.Items.Item("33").Height = oForm.Items.Item("73").Top + oForm.Items.Item("73").Height + 1
                                Catch ex As Exception

                                End Try
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oGrid = oForm.Items.Item("100").Specific
                                Dim oobj As New clsPropertyEvalution
                                oobj.LoadForm_View(oGrid.DataTable.GetValue("Evaluation ID", pVal.Row), oGrid.DataTable.GetValue("Project Name", pVal.Row), oGrid.DataTable.GetValue("Project Code", pVal.Row))

                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "35"
                                        oForm.PaneLevel = 2
                                    Case "135"
                                        oForm.PaneLevel = 1
                                    Case "99"
                                        oForm.PaneLevel = 3
                                        'If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        '    oGrid = oForm.Items.Item("100").Specific
                                        '    oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry] 'Evaluation ID', T0.[U_Z_PROCODE] 'Project Code', T0.[U_Z_PRONAME] 'Project Name', T0.[U_Z_EVL_ID] 'Evaluation Serial No', T0.[U_Z_EV_DATE] 'Evaluation Date', T0.[U_Z_OWNER] 'Owner' FROM [dbo].[@Z_AQEVAL]  T0 where U_Z_ProCode='" & oApplication.Utilities.getEdittextvalue(oForm, "4") & "'")
                                        '    oEditTextColumn = oGrid.Columns.Item(0)
                                        '    oEditTextColumn.LinkedObjectType = "@Z_AQEVAL"
                                        '    oGrid.AutoResizeColumns()
                                        'End If
                                    Case "100"
                                        'If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                        '    oGrid = oForm.Items.Item("100").Specific
                                        '    oGrid.DataTable.ExecuteQuery("SELECT T0.[DocEntry] 'Evaluation ID', T0.[U_Z_PROCODE] 'Project Code', T0.[U_Z_PRONAME] 'Project Name', T0.[U_Z_EVL_ID] 'Evaluation Serial No', T0.[U_Z_EV_DATE] 'Evaluation Date', T0.[U_Z_OWNER] 'Owner' FROM [dbo].[@Z_AQEVAL]  T0 where U_Z_ProCode='" & oApplication.Utilities.getEdittextvalue(oForm, "4") & "'")
                                        '    oEditTextColumn = oGrid.Columns.Item(0)

                                        '    oEditTextColumn.LinkedObjectType = "@Z_AQEVAL"
                                        '    oGrid.AutoResizeColumns()
                                        'End If

                                    Case "74"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            Dim strCode, strname As String
                                            strCode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                            strname = oApplication.Utilities.getEdittextvalue(oForm, "32")
                                            If strCode <> "" Then
                                                Dim oObj As New clsPropertyEvalution
                                                oObj.LoadForm(strCode, strname)
                                            End If
                                        Else
                                            oApplication.Utilities.Message("Evalution Details are accessed in OK Mode", SAPbouiCOM.BoStatusBarMessageType.smt_Error)

                                        End If
                                    Case "98"
                                        If oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                            Dim strcode As String
                                            strcode = oApplication.Utilities.getEdittextvalue(oForm, "4")
                                            Dim oObj As New clsPropertyUnitDetails
                                            oObj.CreateUnitFromProperty(strcode)
                                        End If
                                    Case "26"
                                        fillopen()
                                        oMatrix = oForm.Items.Item("25").Specific
                                        AddRow(oForm)
                                        Try
                                            oForm.Freeze(True)
                                            oApplication.Utilities.SetMatrixValues(oMatrix, "V_1", oMatrix.RowCount, strSelectedFilepath)
                                            Dim strDate As String
                                            Dim dtdate As Date
                                            dtdate = Now.Date
                                            strDate = Now.Date.Today().ToString
                                            ''  strdate=
                                            Dim oColumn As SAPbouiCOM.Column
                                            oColumn = oMatrix.Columns.Item("V_0")
                                            oColumn.Editable = True
                                            'oColumn.Cells.Item(oMatrix.RowCount).Specific.value = "t"
                                            'oApplication.SBO_Application.SendKeys("{TAB}")

                                            oEditText = oMatrix.Columns.Item("V_0").Cells.Item(oMatrix.RowCount).Specific
                                            oEditText.String = "t"
                                            oApplication.SBO_Application.SendKeys("{TAB}")
                                            oForm.Items.Item("21").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
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
                                    Case "27"
                                        LoadFiles(oForm)
                                    Case "34"
                                        oMatrix = oForm.Items.Item("25").Specific
                                        oApplication.SBO_Application.ActivateMenuItem(mnu_DELETE_ROW)
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
                                        If pVal.ItemUID = "29" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If

                                        If pVal.ItemUID = "86" Then
                                            val = oDataTable.GetValue("FormatCode", 0)
                                            oApplication.Utilities.setEdittextvalue(oForm, pVal.ItemUID, val)
                                        End If

                                        If pVal.ItemUID = "93" Then
                                            oApplication.Utilities.setEdittextvalue(oForm, "94", oDataTable.GetValue("CardName", 0))
                                            'Dim oTest As SAPbobsCOM.Recordset
                                            'oTest = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                            'oTest.DoQuery("select CntctPrsn , MailAddres ,Phone1,isnull(Phone2,'') + ',' + isnull(Cellular,''), Fax,E_Mail ,* from OCRD where CardCode='" & oDataTable.GetValue("CardCode", 0) & "'")
                                            'Dim str As String
                                            'str = oTest.Fields.Item(0).Value & "," & oTest.Fields.Item(1).Value & "," & oTest.Fields.Item(2).Value & "," & oTest.Fields.Item(3).Value & "," & oTest.Fields.Item(4).Value & "," & oTest.Fields.Item(5).Value
                                            '' oApplication.Utilities.setEdittextvalue(oForm, "89", oTest.Fields.Item(0).Value)
                                            ''oForm.Items.Item("11").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            '' oApplication.Utilities.setEdittextvalue(oForm, "96", str)
                                            oApplication.Utilities.setEdittextvalue(oForm, "93", oDataTable.GetValue("CardCode", 0))
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
