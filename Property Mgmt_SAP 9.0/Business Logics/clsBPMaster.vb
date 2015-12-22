Public Class clsBPMaster
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
    Private Sub LoadForm(ByVal aForm As SAPbouiCOM.Form)
        Dim strCardCode, strCardtype As String
        If (aForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Or aForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
            oCombo = aForm.Items.Item("40").Specific
            If oCombo.Selected.Value = "C" Then
                strCardCode = oApplication.Utilities.getEdittextvalue(aForm, "5")
                oForm = oApplication.Utilities.LoadForm(xml_ItemGroup, frm_Itemgroup)
                oForm = oApplication.SBO_Application.Forms.ActiveForm()
                oForm.Freeze(True)
                databind(oForm, strCardCode)
                oForm.Freeze(False)
            End If
        End If
    End Sub
#Region "DataBind"
    Private Sub databind(ByVal aForm As SAPbouiCOM.Form, ByVal acode As String)
        Try

            aForm.Freeze(True)
            oGrid = aForm.Items.Item("6").Specific
            oGrid.DataTable.ExecuteQuery("Select * from [@DABT_OITB] where U_CardCode='" & acode & "' order by Convert(Numeric,Code)")
            oApplication.Utilities.setEdittextvalue(aForm, "4", acode)
            FormatGrids(aForm)
            aForm.Freeze(False)
        Catch ex As Exception
            oApplication.Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            aForm.Freeze(False)
        End Try
    End Sub

#Region "FormatGrids"
    Private Sub FormatGrids(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("6").Specific
        oGrid.Columns.Item(0).Visible = False
        oGrid.Columns.Item(0).TitleObject.Caption = "Code"
        oGrid.Columns.Item(1).TitleObject.Caption = "Name"
        oGrid.Columns.Item(1).Visible = False
        oGrid.Columns.Item(2).TitleObject.Caption = "Card Code"
        oGrid.Columns.Item(2).Visible = False
        oGrid.Columns.Item(3).TitleObject.Caption = "Item Group"
        oGrid.Columns.Item(3).Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
        oRecset = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecset.DoQuery("Select ItmsGrpCod,ItmsGrpnam from OITB order by ItmsGrpCod")
        oCombobox = oGrid.Columns.Item(3)
        For intRow As Integer = 0 To oRecset.RecordCount - 1
            oCombobox.ValidValues.Add(oRecset.Fields.Item(0).Value, oRecset.Fields.Item(1).Value)
            oRecset.MoveNext()
        Next
        oCombobox.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_both

        oGrid.AutoResizeColumns()
        oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single


    End Sub
#End Region

#Region "AddRow /Delete Row"
    Private Sub AddRow(ByVal aForm As SAPbouiCOM.Form)

        oGrid = aForm.Items.Item("9").Specific
    End Sub


#End Region





#End Region

#End Region

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
                        If oForm.TypeEx = frm_BPMaster Then
                            If oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                                oForm.Items.Item("btnDisplay").Visible = False
                            Else
                                oCombo = oForm.Items.Item("40").Specific
                                If oCombo.Selected.Value = "C" Then
                                    oForm.Items.Item("btnDisplay").Visible = True
                                Else
                                    oForm.Items.Item("btnDisplay").Visible = False

                                End If

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
            If pVal.FormTypeEx = frm_BPMaster Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                oApplication.Utilities.AddControls(oForm, "btnDisplay", "2", SAPbouiCOM.BoFormItemTypes.it_BUTTON, "RIGHT", 1, 1, "2", "Item Group Mapping")
                            Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                                Select Case pVal.ItemUID
                                    Case "btnDisplay"
                                        LoadForm(oForm)
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
