Imports System.IO
Public Class clsAppHisDetails
    Inherits clsBase
    Private oCFLEvent As SAPbouiCOM.IChooseFromListEvent
    Private oEditText As SAPbouiCOM.EditText
    Private oCombobox As SAPbouiCOM.ComboBox
    Private oCombo As SAPbouiCOM.ComboBoxColumn
    Private oEditTextColumn As SAPbouiCOM.EditTextColumn
    Private oGrid As SAPbouiCOM.Grid
    Private dtDocumentList As SAPbouiCOM.DataTable
    Private dtHistoryList As SAPbouiCOM.DataTable
    Private InvForConsumedItems As Integer
    Private blnFlag As Boolean = False

    Public Sub New()
        MyBase.New()
        InvForConsumedItems = 0
    End Sub


    Public Sub LoadViewHistory(ByVal enDocType As String, ByVal strDocEntry As String)
        oForm = oApplication.Utilities.LoadForm(xml_AppHisDetails, frm_AppHisDetails)
        oForm = oApplication.SBO_Application.Forms.ActiveForm()
        Try
            oForm.Freeze(True)
            If enDocType = "TEA" Then
                oForm.Title = "Contract Approval History"
            ElseIf enDocType = "TER" Then
                oForm.Title = "Termination Approval History"
            End If
            Dim sQuery As String
            oGrid = oForm.Items.Item("3").Specific
            sQuery = " Select DocEntry,U_Z_DOCENTRY,U_Z_DOCTYPE,U_Z_EMPID,U_Z_EMPNAME,U_Z_APPROVEBY,CreateDate ,CreateTime,UpdateDate,UpdateTime,U_Z_APPSTATUS,U_Z_REMARKS From [@Z_APHIS] "
            sQuery += " Where U_Z_DOCTYPE = '" + enDocType.ToString() + "'"
            sQuery += " And U_Z_DOCENTRY = '" + strDocEntry + "'"
            oGrid.DataTable.ExecuteQuery(sQuery)
            SummaryformatHistory(oForm)
            oApplication.Utilities.assignMatrixLineno(oGrid, oForm)
            oForm.Freeze(False)
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        End Try
    End Sub
     Private Sub SummaryformatHistory(ByVal aForm As SAPbouiCOM.Form)
        Try
            aForm.Freeze(True)
            Dim oGrid As SAPbouiCOM.Grid
            Dim oGridCombo As SAPbouiCOM.ComboBoxColumn
            Dim oEditTextColumn As SAPbouiCOM.EditTextColumn
            oGrid = aForm.Items.Item("3").Specific
            oGrid.Columns.Item("DocEntry").Visible = False
            oGrid.Columns.Item("U_Z_DOCENTRY").TitleObject.Caption = "Reference No."
            oGrid.Columns.Item("U_Z_DOCENTRY").Visible = False
            oGrid.Columns.Item("U_Z_DOCTYPE").Visible = False
            oGrid.Columns.Item("U_Z_EMPID").TitleObject.Caption = "Employee ID"
            oEditTextColumn = oGrid.Columns.Item("U_Z_EMPID")
            oEditTextColumn.LinkedObjectType = "171"
            oGrid.Columns.Item("U_Z_EMPNAME").TitleObject.Caption = "Employee Name"
            oGrid.Columns.Item("U_Z_APPROVEBY").TitleObject.Caption = "Approved By"
            oGrid.Columns.Item("U_Z_APPSTATUS").TitleObject.Caption = "Approved Status"
            oGrid.Columns.Item("U_Z_APPSTATUS").Type = SAPbouiCOM.BoGridColumnType.gct_ComboBox
            oGridCombo = oGrid.Columns.Item("U_Z_APPSTATUS")
            oGridCombo.ValidValues.Add("P", "Pending")
            oGridCombo.ValidValues.Add("A", "Approved")
            oGridCombo.ValidValues.Add("R", "Rejected")
            oGridCombo.DisplayType = SAPbouiCOM.BoComboDisplayType.cdt_Description
            oGrid.Columns.Item("U_Z_REMARKS").TitleObject.Caption = "Remarks"
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single
            oGrid.AutoResizeColumns()
            aForm.Freeze(False)
        Catch ex As Exception
            aForm.Freeze(False)
            Throw ex
        End Try
    End Sub

#Region "Item Event"
    Public Overrides Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = frm_AppHisDetails Then
                Select Case pVal.BeforeAction
                    Case True
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD

                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                        End Select
                    Case False
                        Select Case pVal.EventType
                            Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                oForm = oApplication.SBO_Application.Forms.Item(FormUID)
                            Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED
                            
                            Case SAPbouiCOM.BoEventTypes.et_CLICK
                             
                            Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN
                            Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                                'If oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore Or oForm.State = SAPbouiCOM.BoFormStateEnum.fs_Maximized Then
                                '    oApplication.Utilities.Resize(oForm)
                                'End If
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
                Case mnu_FIRST, mnu_LAST, mnu_NEXT, mnu_PREVIOUS
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
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

End Class
