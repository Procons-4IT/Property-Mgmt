Public NotInheritable Class clsTable

#Region "Private Functions"
    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Tables in DB. This function shall be called by 
    '                     public functions to create a table
    '**************************************************************************************************************
    Private Sub AddTables(ByVal strTab As String, ByVal strDesc As String, ByVal nType As SAPbobsCOM.BoUTBTableType)
        Dim oUserTablesMD As SAPbobsCOM.UserTablesMD
        Try
            oUserTablesMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
            'Adding Table
            If Not oUserTablesMD.GetByKey(strTab) Then
                oUserTablesMD.TableName = strTab.ToUpper()
                oUserTablesMD.TableDescription = strDesc.ToUpper()
                oUserTablesMD.TableType = nType
                If oUserTablesMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD)
            oUserTablesMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : AddFields
    'Parameter          : SstrTab As String,strCol As String,
    '                     strDesc As String,nType As Integer,i,nEditSize,nSubType As Integer
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Generic Function for adding all Fields in DB Tables. This function shall be called by 
    '                     public functions to create a Field
    '**************************************************************************************************************
    Private Sub AddFields(ByVal strTab As String, _
                            ByVal strCol As String, _
                                ByVal strDesc As String, _
                                    ByVal nType As SAPbobsCOM.BoFieldTypes, _
                                        Optional ByVal i As Integer = 0, _
                                            Optional ByVal nEditSize As Integer = 10, _
                                                Optional ByVal nSubType As SAPbobsCOM.BoFldSubTypes = 0, _
                                                    Optional ByVal Mandatory As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO, Optional ByVal linktable As String = "")
        Dim oUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try
            If Not (strTab = "JDT1" Or strTab = "OCRD" Or strTab = "RCT1" Or strTab = "OITB" Or strTab = "ODPS" Or strTab = "ORCT" Or strTab = "OITM" Or strTab = "ODPI" Or strTab = "INV1" Or strTab = "OWTR" Or strTab = "OINV" Or strTab = "QUT1" Or strTab = "OPRJ") Then
                strTab = "@" + strTab
            End If

            If Not IsColumnExists(strTab, strCol) Then
                oUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                oUserFieldMD.Description = strDesc.ToUpper()
                oUserFieldMD.Name = strCol.ToUpper
                oUserFieldMD.Type = nType
                oUserFieldMD.SubType = nSubType
                oUserFieldMD.TableName = strTab
                oUserFieldMD.EditSize = nEditSize
                oUserFieldMD.Mandatory = Mandatory
                If linktable <> "" Then
                    oUserFieldMD.LinkedTable = linktable
                End If
                If oUserFieldMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD)

            End If

        Catch ex As Exception
            Throw ex
        Finally
            oUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If
            If TableName.StartsWith("@") Or TableName = "ODPI" Or TableName = "OINV" Or TableName = "OCRD" Or TableName = "OITM" Or TableName = "OITB" Or TableName = "ORDR" Or TableName = "OUSR" Or TableName = "OACT" Then
            Else
                TableName = "@" & TableName
            End If
            If (Not IsColumnExists(TableName, ColumnName)) Then
                objUserFieldMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName.ToUpper()
                objUserFieldMD.Description = ColDescription.ToUpper()
                objUserFieldMD.Type = FieldType
                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                If SetValidValue <> "" Then
                    objUserFieldMD.DefaultValue = SetValidValue
                End If
                For intLoop = 0 To strValue.GetLength(0) - 1
                    objUserFieldMD.ValidValues.Value = strValue(intLoop)
                    objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                    objUserFieldMD.ValidValues.Add()
                Next
                If (objUserFieldMD.Add() <> 0) Then
                    MsgBox(oApplication.Company.GetLastErrorDescription)
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            Else

            End If

        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            objUserFieldMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try


    End Sub

    '*************************************************************************************************************
    'Type               : Private Function
    'Name               : IsColumnExists
    'Parameter          : ByVal Table As String, ByVal Column As String
    'Return Value       : Boolean
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Function to check if the Column already exists in Table
    '**************************************************************************************************************
    Private Function IsColumnExists(ByVal Table As String, ByVal Column As String) As Boolean
        Dim oRecordSet As SAPbobsCOM.Recordset

        Try
            strSQL = "SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & Table & "' AND ""AliasID"" = '" & Column & "'"
            oRecordSet = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRecordSet.DoQuery(strSQL)

            If oRecordSet.Fields.Item(0).Value = 0 Then
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet)
            oRecordSet = Nothing
            GC.Collect()
        End Try
    End Function

    Private Sub AddKey(ByVal strTab As String, ByVal strColumn As String, ByVal strKey As String, ByVal i As Integer)
        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        Try
            '// The meta-data object must be initialized with a
            '// regular UserKeys object
            oUserKeysMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

            If Not oUserKeysMD.GetByKey("@" & strTab, i) Then

                '// Set the table name and the key name
                oUserKeysMD.TableName = strTab
                oUserKeysMD.KeyName = strKey

                '// Set the column's alias
                oUserKeysMD.Elements.ColumnAlias = strColumn
                oUserKeysMD.Elements.Add()
                oUserKeysMD.Elements.ColumnAlias = "RentFac"

                '// Determine whether the key is unique or not
                oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES

                '// Add the key
                If oUserKeysMD.Add <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If

            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)
            oUserKeysMD = Nothing
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try

    End Sub

    '********************************************************************
    'Type		            :   Function    
    'Name               	:	AddUDO
    'Parameter          	:   
    'Return Value       	:	Boolean
    'Author             	:	
    'Created Date       	:	
    'Last Modified By	    :	
    'Modified Date        	:	
    'Purpose             	:	To Add a UDO for Transaction Tables
    '********************************************************************
    Private Sub AddUDO(ByVal strUDO As String, ByVal strDesc As String, ByVal strTable As String, _
                                Optional ByVal sFind1 As String = "", Optional ByVal sFind2 As String = "", _
                                        Optional ByVal strChildTbl As String = "", Optional ByVal strChildTb2 As String = "", Optional ByVal nObjectType As SAPbobsCOM.BoUDOObjType = SAPbobsCOM.BoUDOObjType.boud_Document, Optional ByVal strChildTb3 As String = "")

        Dim oUserObjectMD As SAPbobsCOM.UserObjectsMD
        Try
            oUserObjectMD = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjectMD.GetByKey(strUDO) = 0 Then
                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                If strTable = "Z_PROPUNIT" Or strTable = "Z_PROP" Then
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                End If
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                If sFind1 <> "" And sFind2 <> "" Then
                    oUserObjectMD.FindColumns.ColumnAlias = sFind1
                    oUserObjectMD.FindColumns.Add()
                    oUserObjectMD.FindColumns.SetCurrentLine(1)
                    oUserObjectMD.FindColumns.ColumnAlias = sFind2
                    oUserObjectMD.FindColumns.Add()
                End If

                oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.LogTableName = ""
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjectMD.ExtensionName = ""

                If strChildTbl <> "" Then
                    oUserObjectMD.ChildTables.TableName = strChildTbl
                End If
                If strChildTb2 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                    End If
                    oUserObjectMD.ChildTables.TableName = strChildTb2
                End If

                If strChildTb3 <> "" Then
                    If strChildTbl <> "" Then
                        oUserObjectMD.ChildTables.Add()
                    End If
                    oUserObjectMD.ChildTables.TableName = strChildTb3
                End If

                oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjectMD.Code = strUDO
                oUserObjectMD.Name = strDesc
                oUserObjectMD.ObjectType = nObjectType
                oUserObjectMD.TableName = strTable

                If oUserObjectMD.Add() <> 0 Then
                    Throw New Exception(oApplication.Company.GetLastErrorDescription)
                End If
            End If

        Catch ex As Exception
            Throw ex

        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD)
            oUserObjectMD = Nothing
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try

    End Sub

#End Region

#Region "Public Functions"
    '*************************************************************************************************************
    'Type               : Public Function
    'Name               : CreateTables
    'Parameter          : 
    'Return Value       : none
    'Author             : Manu
    'Created Dt         : 
    'Last Modified By   : 
    'Modified Dt        : 
    'Purpose            : Creating Tables by calling the AddTables & AddFields Functions
    '**************************************************************************************************************
    Public Sub CreateTables()
        ' Dim oProgressBar As SAPbouiCOM.ProgressBar
        Try
            oApplication.Utilities.Message("Initializing Database...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

            AddTables("Z_OAPPT", "Approval Template", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddTables("Z_APPT2", "Approval Authorizer", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_OAPPT", "Z_Code", "Approval Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OAPPT", "Z_Name", "Approval Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OAPPT", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_OAPPT", "Z_DocDesc", "Document Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OAPPT", "Z_Active", "Active Template", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_APPT2", "Z_AUser", "Authorizer Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APPT2", "Z_AName", "Authorizer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_APPT2", "Z_AMan", "Mandatory", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_APPT2", "Z_AFinal", "Final Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)

            AddTables("Z_APHIS", "Approval History", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_APHIS", "Z_DocEntry", "Document Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_DocType", "Document Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_EmpId", "Employee Id", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            addField("@Z_APHIS", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            AddFields("Z_APHIS", "Z_Remarks", "Comments", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_APHIS", "Z_ApproveBy", "Approved By", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_APHIS", "Z_Approvedt", "Approver Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_APHIS", "Z_ADocEntry", "Template DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_APHIS", "Z_ALineId", "Template LineId", SAPbobsCOM.BoFieldTypes.db_Numeric)





            addField("OITM", "Z_ProFlg", "Property Type Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("OITM", "Z_PrjCode", "Project Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OITM", "Z_CostCenter", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("ORCT", "Z_ContID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ORCT", "Z_ContNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 19)
            AddFields("ORCT", "Z_CntNumber", "Contract Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("ORCT", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddFields("ORCT", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("ORCT", "Z_ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("RCT1", "Z_ContNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 19)
            AddFields("RCT1", "Z_ContID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("RCT1", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("RCT1", "Z_CntNumber", "Contract Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("RCT1", "Z_FromDate", "From Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("RCT1", "Z_ToDate", "To Date", SAPbobsCOM.BoFieldTypes.db_Date)

            'addField("OINV", "Z_InvType", "Invoice Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,T", "Owner,Tenant", "T")


            AddFields("ODPS", "Z_ContID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODPS", "Z_ContNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("ODPS", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODPS", "Z_CntNumber", "Contract Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("JDT1", "Z_ContID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("JDT1", "Z_ContNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("JDT1", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("JDT1", "Z_CntNumber", "Contract Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("ODPI", "Z_ContID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODPI", "Z_ContNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("ODPI", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("ODPI", "Z_CntNumber", "Contract Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("ODPI", "Z_StartDate", "Contract Start date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("ODPI", "Z_EndDate", "Contract End date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("ODPI", "Z_DPType", "Down payment Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "N,A,S", "Normal,Annual rental,Security Deposit", "N")
            addField("ODPI", "Z_InvType", "Invoice Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,T", "Owner,Tenant", "O")

            AddFields("OCRD", "Nationality", "Nationality", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("OCRD", "Occupation", "Occupation", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("OCRD", "MaritalStatus", "Marital Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("OCRD", "RegDate", "Registration Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddTables("Z_UnitType", "Property Unit Type", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_UnitType", "Z_UnitType", "Property Unit Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_UnitType", "Z_Status", "Property Unit Status", SAPbobsCOM.BoFieldTypes.db_Alpha, , 1)
            AddFields("Z_UnitTYpe", "Z_FrgnName", "Second Lanugage Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_PROP", "Property Data", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_PROP", "Z_Code", "Property Code ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 7)
            AddFields("Z_PROP", "Z_Type", "Property Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROP", "Z_Desc", "Property Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROP", "Z_Location", "Property Location", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROP", "Z_Address", "Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROP", "Z_NoofFloor", "Number of Floors", SAPbobsCOM.BoFieldTypes.db_Numeric, , )
            AddFields("Z_PROP", "Z_NoofUnits", "Number of Property Units", SAPbobsCOM.BoFieldTypes.db_Numeric, , )
            AddFields("Z_PROP", "Z_Noofoutlets", "Number of Outlets", SAPbobsCOM.BoFieldTypes.db_Numeric, , )
            addField("@Z_PROP", "Z_Area", "Area/Unit", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "L,N", "Leasable,Non-Leasable", "L")
            AddFields("Z_PROP", "Z_Total", "Total Leasable Area", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PROP", "Z_TotalArea", "TotalArea in the building", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PROP", "Z_ActCode", "Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROP", "Z_ActName", "Account Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROP", "Z_FrgnName", "Second Lanuage Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'New field additions 
            AddFields("Z_PROP", "Z_Parking", "Number of Parking", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Pools", "Number of Swimming Pools", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Intercom", "Number of Intercom", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Garage", "Garage", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_No", "Building Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PROP", "Z_Name", "Building Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            '  AddFields("Z_PROP", "Z_Description", "Full Description", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROP", "Z_AreaSqMe", "Building Area Sq.Meter", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PROP", "Z_AreaFeet", "Building Area Sq.Feet", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PROP", "Z_Lifts", "Lifts", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_ClubHouse", "Club House", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddFields("Z_PROP", "Z_Gardens", "Gardens", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_SuperMarket", "SuperMarket", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Security", "Security", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_EleNo", "Shared electicity Number", SAPbobsCOM.BoFieldTypes.db_Alpha)
            AddFields("Z_PROP", "Z_WaterNo", "Shared Water Number", SAPbobsCOM.BoFieldTypes.db_Alpha)
            AddFields("Z_PROP", "Z_Arial", "Arial", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Outer", "Outer Lighting", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROP", "Z_Quality", "Building Quality", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROP", "Z_Notes", "Building Notes", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddFields("Z_PROP", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROP", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROP", "Z_SigDate", "Signature Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROP", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            addField("@Z_PROP", "Z_MType", "Property Type Management", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,T", "Owned,Third Party", "T")
            AddFields("Z_PROP", "Z_ComGL", "Commission Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_PROP", "Z_Firstparty", "First Party Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_PROP", "Z_Address2", "First Party Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            'AddFields("Z_PROP", "Z_Phone1", "Company Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_PROP", "Z_Fax1", "Company Fax", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_PROP", "Z_Email", "Company Email", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_PROP", "Z_CardCode", "Second Party Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROP", "Z_CardName", "Second Party Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROP", "Z_Address1", "Second Party Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROP", "Z_TitleNo", "Title Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROP", "Z_RegFor", "Register For", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            'AddFields("Z_PROP", "Z_Phone2", "Second Party Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_PROP", "Z_Fax2", "Second Party Fax", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            'AddFields("Z_PROP", "Z_Emai2", "Second Party Email", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)


            'end New field additions
            AddTables("Z_PROP1", "Property Data Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROP1", "Z_Filename", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROP1", "Z_AttachDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROP1", "Z_AttName", "Attachment Detals", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)




            AddTables("Z_PROPUNIT", "Property Unit Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_PROPUNIT", "Z_Code", "Property Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 12)
            AddFields("Z_PROPUNIT", "Z_Desc", "Unit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)


            AddFields("Z_PROPUNIT", "Z_PropCode", "Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 7)
            AddFields("Z_PROPUNIT", "Z_PropDesc", "Property Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROPUNIT", "Z_ProItemCode", "Property ItemCode", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("Z_PROPUNIT", "Z_UnitTYpe", "Unit Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_PROPUNIT", "Z_UnitStatus", "Unit Statu", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,R,L,S,A,N,M", "Offered,Reserved(Before Signature),Leased(After Signature),Sold,Available,NotAvailable,Under Maintenance ", "O")
            AddFields("Z_PROPUNIT", "Z_FurType", "Furniture Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT", "Z_NoofFur", "Number of Furniture", SAPbobsCOM.BoFieldTypes.db_Numeric, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PROPUNIT", "Z_Space", "Space Provided by Sqm", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            addField("@Z_PROPUNIT", "Z_FSType", "Floor Space Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "W,I,C,R", "Windows,Inside Office,Corridor,Receiption Area", "W")
            AddFields("Z_PROPUNIT", "Z_Factor", "Mulitplication factor", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PROPUNIT", "Z_Price", "Sales price per square meter", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_PROPUNIT", "Z_TotalArea", "Total Area", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_PROPUNIT", "Z_TotalPrice", "Total Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_PROPUNIT", "Z_PriceList", "Price List", SAPbobsCOM.BoFieldTypes.db_Alpha, , )
            AddFields("Z_PROPUNIT", "Z_Rules", "Rules and Regulations", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROPUNIT", "Z_CostCenter", "Cost Center", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            'New Fields
            '  AddFields("Z_PROPUNIT", "Z_AreaFloor", "Area Floor", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            '  AddFields("Z_PROPUNIT", "Z_FootPrice", "Foot Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            '  AddFields("Z_PROPUNIT", "Z_AreaPrice", "Area Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_PROPUNIT", "Z_Notes", "Location Notes", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_PROPUNIT", "Z_OwnerCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROPUNIT", "Z_OwnerName", "Owner Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROPUNIT", "Z_Color", "Color", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PROPUNIT", "Z_Floor", "Floor", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Electricity", "Electricity", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Water", "Water", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Saloon", "Saloon", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Dinning", "Dinning Rooms", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Master", "Master Rooms", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Rooms", "Rooms", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Kitchens", "Kitchens", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_RestRoom", "Rest Rooms", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Balcones", "Balcones", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_AcType", "AC Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT", "Z_NoofAcs", "Number of Ac's", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Parking", "Parking", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Pools", "Swimming Pools", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_InterCome", "InterComes", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_Garage", "Garage", SAPbobsCOM.BoFieldTypes.db_Numeric)
            ' AddFields("Z_PROPUNIT", "Z_Furniture", "Furniture", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_PROPUNIT", "Z_UnitLevel", "Unit Levels", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROPUNIT", "Z_RegDate", "Title Number", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROPUNIT", "Z_TitleNo", "Title Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROPUNIT", "Z_PreViewer", "PreViewer", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROPUNIT", "Z_UnitCity", "Unit City", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PROPUNIT", "Z_UnitArea", "Unit Area", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PROPUNIT", "Z_UnitStreet", "Unit Street", SAPbobsCOM.BoFieldTypes.db_Alpha, , 25)
            AddFields("Z_PROPUNIT", "Z_RefPoint", "Reference Point", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT", "Z_UnitClass", "Unit Class", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_PROPUNIT", "Z_UnitNo", "Unit Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_GuestRoom", "Guest Room", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_Majles", "Majles", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_Attachments", "Attached Units", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_AttContent", "Attached Unit Content", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROPUNIT", "Z_Stores", "Stores", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_Quality", "Quality", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_BuilName", "Building name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_PROPUNIT", "Z_UnitsInFloor", "Units In floor", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_PROPUNIT", "Z_Lifts", "Lifts", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_PROPUNIT", "Z_ClubHouse", "Club House", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_PROPUNIT", "Z_Gardens", "Gardens", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_PROPUNIT", "Z_SuperMarket", "Super Markets", SAPbobsCOM.BoFieldTypes.db_Alpha, , 5)
            AddFields("Z_PROPUNIT", "Z_Security", "Security", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_SharedElec", "Shared Electricity", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_SharedWater", "Shared Water", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_PROPUNIT", "Z_Lights", "Lighting", SAPbobsCOM.BoFieldTypes.db_Alpha, , 3)
            AddFields("Z_PROPUNIT", "Z_OwnerRefNo", "Owener Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT", "Z_ContractNo", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PROPUNIT", "Z_SaleValue", "Sales Value", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PROPUNIT", "Z_RenValue", "Rental Value", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_PROPUNIT", "Z_OwnerForName", "Owner For name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_PROPUNIT", "Z_Broker", "Broker", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PROPUNIT", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)

            AddFields("Z_PROPUNIT", "Z_PrvName", "Previwer Details", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROPUNIT", "Z_FrgnName", "Second Lanuage Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_PROPUNIT", "Z_UsedFor1", "Used For Details", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)


            'End New Fields
        
            AddTables("Z_PROPUNIT1", "Property Unit Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROPUNIT1", "Z_Filename", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROPUNIT1", "Z_AttachDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_PROPUNIT1", "Z_AttName", "Attachment Detals", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)



            AddTables("Z_RESER", "Reservation", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_RESER", "Z_PropCode", "Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_RESER", "Z_PropDesc", "Property Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_UnitCode", "Property Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_RESER", "Z_UnitDesc", "Property Unit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_RESER", "Z_CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_Address", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 253)
            AddFields("Z_RESER", "Z_AgentCode", "Sales Agent Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_RESER", "Z_AgentName", "Sale Agent Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_Empid", "Employee responsible", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_RESER", "Z_EmpName", "Employee Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_RESER", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_RESER", "Z_IssueDate", "Issue Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_RESER", "Z_PayTerms", "Payment Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_RESER", "Z_Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Memo)
            addField("@Z_RESER", "Z_DownPay", "Down Payment", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_RESER", "Z_DownPayRef", "Down Payment Reference Entry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_RESER", "Z_DownNumber", "Down Payment Reference Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_RESER", "Z_DownAmount", "Down Payment Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("Z_RESER", "Z_AcctCode", "Receiable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_RESER", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            addField("@Z_RESER", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,C", "Pending for confirmation,Confirmed", "P")



            AddTables("Z_CONTRACT", "Maintain Tenant Contracts", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_CONTRACT", "Z_UnitCode", "Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CONTRACT", "Z_Desc", "Unit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_CONTRACT", "Z_ContDate", "Contract date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_CONTRACT", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_CONTRACT", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_Address, "PED,APP,AGR,TER,CAN,RED", "Pending for Approval,Approved,Agreed,Terminated,Cancelled,Renewed", "PED")
            AddFields("Z_CONTRACT", "Z_TenCode", "Tenant Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CONTRACT", "Z_TenName", "Tenant Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_OffAddress", "Office Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT", "Z_AlJAddress", "AI Jassim Group Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT", "Z_Annualrent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT", "Z_AcctCode", "Receiable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_Deposit", "Security Deposit", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT", "Z_PayTrms", "Payment Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("@Z_CONTRACT", "Z_Insurance", "Insurance", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_CONTRACT", "Z_PolicyNumber", "Insurance PolicyNumber ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_ChgMonth", "Number of Months", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT", "Z_ChgAmt", "Termination Charges", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT", "Z_Period", "Termination Application Period", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT", "Z_TermDate", "Termination Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT", "Z_Rules", "Rules and Regulations", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT", "Z_DPEntry", "Down Payment DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT", "Z_DPNumber", "Down Payment Number", SAPbobsCOM.BoFieldTypes.db_Alpha)
            addField("@Z_CONTRACT", "Z_Type", "Contract Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,T", "Owner,Tenant", "O")
            AddFields("Z_CONTRACT", "Z_AcctCode1", "Tenant Receiable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_OwnerCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CONTRACT", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT", "Z_ConNo", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_LiaAc", "Liability Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_CONTRACT", "Z_ProType", "Property Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,T", "Owned,Third Party", "T")
            AddFields("Z_CONTRACT", "Z_CommAc", "Commission Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_CONTRACT", "Z_Master", "Main Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_BaseConNo", "Base Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_RenewalNo", "Renewal Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_BaseStartDate", "Base Contract Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT", "Z_BaseEndDate", "Base Contract End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("Z_CONTRACT", "Z_RenStatus", "Renewal Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_CONTRACT", "Z_UsedFor1", "Used For Details", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_BaseEntry", "Base COntract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT", "Z_BaseSeq", "Base Contract Sequence No", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddFields("Z_CONTRACT", "Z_SeqNo", "Renewal Sequence Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT", "Z_CntNo", "Contract Number with Sequance", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)

            AddFields("Z_CONTRACT", "Z_Monthly", "Monthly Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_CONTRACT", "Z_IsCommission", "Commission Paid", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            addField("@Z_CONTRACT", "Z_AppStatus", "Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_CONTRACT", "Z_TerStatus", "Termination Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_CONTRACT", "Z_DPAcctCode", "Down Payment Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT", "Z_SDAcctCode", "Security Deposit Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            '  addfields("Z_CONTACT","Z_PolicyName","Insurance PolicyName",
           

            AddTables("Z_CONTRACT1", "Contract Expenses Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CONTRACT1", "Z_EXPREFCODEREF", "Expense Code reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_CONTRACT1", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_CONTRACT1", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT1", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("@Z_CONTRACT1", "Z_TYPE", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,F,P", "Sq.Met,Fixed,Percentage", "F")
            addField("@Z_CONTRACT1", "Z_FREQUENCY", "Frequency", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "M,Q,H,Y,O", "Monthly,Quarterly,Half yearly,Yearly,One Time", "M")
            AddFields("Z_CONTRACT1", "Z_RATE", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONTRACT1", "Z_AMOUNT1", "Expense Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONTRACT1", "Z_MONTHS", "Posting Months", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            addField("@Z_CONTRACT1", "Z_RENEWAL", "Exclude in Renewal", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")

            AddTables("Z_CONTRACT2", "Property Unit Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CONTRACT2", "Z_Filename", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_CONTRACT2", "Z_AttachDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT2", "Z_AttName", "Attachment Detals", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)



            AddTables("Z_CONTRACT_OWN", "Maintain Owner Contracts", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_CONTRACT_OWN", "Z_UnitCode", "Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CONTRACT_OWN", "Z_Desc", "Unit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_CONTRACT_OWN", "Z_ContDate", "Contract date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_CONTRACT_OWN", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT_OWN", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_CONTRACT_OWN", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_Address, "PED,APP,AGR,TER,CAN", "Pending for Approval,Approved,Agreed,Terminated,Cancelled", "PED")
            AddFields("Z_CONTRACT_OWN", "Z_TenCode", "Tenant Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CONTRACT_OWN", "Z_TenName", "Tenant Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_OffAddress", "Office Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT_OWN", "Z_AlJAddress", "AI Jassim Group Address", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT_OWN", "Z_Annualrent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT_OWN", "Z_AcctCode", "Receiable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_Deposit", "Security Deposit", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT_OWN", "Z_PayTrms", "Payment Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            addField("@Z_CONTRACT_OWN", "Z_Insurance", "Insurance", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_CONTRACT_OWN", "Z_PolicyNumber", "Insurance PolicyNumber ", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_ChgMonth", "Number of Months", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT_OWN", "Z_ChgAmt", "Termination Charges", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT_OWN", "Z_Period", "Termination Application Period", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT_OWN", "Z_TermDate", "Termination Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT_OWN", "Z_Rules", "Rules and Regulations", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_CONTRACT_OWN", "Z_DPEntry", "Down Payment DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT_OWN", "Z_DPNumber", "Down Payment Number", SAPbobsCOM.BoFieldTypes.db_Alpha)
            addField("@Z_CONTRACT_OWN", "Z_Type", "Contract Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "O,T", "Owner,Tenant", "O")
            AddFields("Z_CONTRACT_OWN", "Z_AcctCode1", "Tenant Receiable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_OwnerCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_CONTRACT_OWN", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_CONTRACT_OWN", "Z_ConNo", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_LiaAc", "Liability Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_UsedFor1", "Used For Details", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN", "Z_SeqNo", "Renewal Sequence Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONTRACT_OWN", "Z_CntNo", "Contract Number with Sequance", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            '  addfields("Z_CONTACT","Z_PolicyName","Insurance PolicyName",


            AddTables("Z_CONTRACT_OWN1", "OwnerContract Expenses Details", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CONTRACT_OWN1", "Z_EXPREFCODEREF", "Expense Code reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_CONTRACT_OWN1", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_CONTRACT_OWN1", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_CONTRACT_OWN1", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("@Z_CONTRACT_OWN1", "Z_TYPE", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,F,P", "Sq.Met,Fixed,Percentage", "F")
            addField("@Z_CONTRACT_OWN1", "Z_FREQUENCY", "Frequency", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "M,Q,H,Y,O", "Monthly,Quarterly,Half yearly,Yearly,One Time", "M")
            AddFields("Z_CONTRACT_OWN1", "Z_RATE", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONTRACT_OWN1", "Z_AMOUNT1", "Expense Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONTRACT_OWN1", "Z_MONTHS", "Posting Months", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)

            AddTables("Z_CONTRACT_OWN2", "OwnerProperty Unit Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_CONTRACT_OWN2", "Z_Filename", "Attachment File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_CONTRACT_OWN2", "Z_AttachDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONTRACT_OWN2", "Z_AttName", "Attachment Detals", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddTables("Z_INSURANCE", "Insurance Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_INSURANCE", "Z_CompName", "Insurance Comapny Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_INSURANCE", "Z_PolicyNumber", "Policy Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_INSURANCE", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_INSURANCE", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_INSURANCE", "Z_TenCode", "Tenant Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_INSURANCE", "Z_UnitCode", "Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)


            AddTables("Z_PRICELIST", "Price List Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_PRICELIST", "Z_PrjCode", "Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_PRICELIST", "Z_PrlNam", "Pricelist Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PRICELIST", "Z_Price", "Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            addField("@Z_PRICELIST", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,P,R", "Approved,Pending,Rejected", "P")

            AddTables("Z_OEXP", "Expenses Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OEXP", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OEXP", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OEXP", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("@Z_OEXP", "Z_TYPE", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,F,P", "Sq.Met,Fixed,Percentage", "F")
            AddFields("Z_OEXP", "Z_RATE", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_OEXP", "Z_RENEWAL", "Exclude in Renewal", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            addField("@Z_OEXP", "Z_FREQUENCY", "Frequency", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "M,Q,H,Y,O", "Monthly,Quarterly,Half yearly,Yearly,One Time", "M")

            AddTables("Z_OPROTYPE", "Property Type Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OPROTYPE", "Z_CODE", "Property Type Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OPROTYPE", "Z_NAME", "Property Type Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_OPROTYPE", "Z_FRGNNAME", "Second Lanugage Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_OLOC", "Location Master", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OLOC", "Z_FrgnName", "Second Lanuage Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddTables("Z_OBILL", "Contract Bill Generation", SAPbobsCOM.BoUTBTableType.bott_NoObject)

            AddFields("Z_OBILL", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "ContractID", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "UnitCode", "Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OBILL", "Space", "Space in Sq.Mtrs", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Quantity)
            AddFields("Z_OBILL", "CardCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OBILL", "Annualrent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL", "PayTrms", "Payment Terms", SAPbobsCOM.BoFieldTypes.db_Alpha, , 10)
            AddFields("Z_OBILL", "ChgMonth", "Number of Months", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "MonthRent", "Monthly Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL", "RentGL", "Reciable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OBILL", "Expenses", "Total Expenses", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL", "Total", "Grand total", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_OBILL", "Invoiced", "Invoiced", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_OBILL", "InvEntry", "Invoice DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "InvNumber", "Invoice DocNumber", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "ContractNumber", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_OBILL", "Z_ProType", "Property Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,T", "Owned,Third Party", "T")
            AddFields("Z_OBILL", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OBILL", "ComPer", "Commission Percentage", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Percentage)
            AddFields("Z_OBILL", "Commission", "Commission Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL", "CommGL", "Reciable Account Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OBILL", "OwnerCode", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OBILL", "Seq", "Sequane number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL", "DocDate", "Posting Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OBILL", "CntNo", "Contract Number with Seq", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OBILL", "ExtUnitCode", "Ext.Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OBILL", "ProName", "Property Unit Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_OBILL", "CardName", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)



            AddTables("Z_OBILL1", "Contract Expenses Details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OBILL1", "Z_RefNo", "Bill Reference Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_OBILL1", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OBILL1", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OBILL1", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            addField("@Z_OBILL1", "Z_TYPE", "Expense Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,F,P", "Sq.Met,Fixed,Percentage", "F")
            AddFields("Z_OBILL1", "Z_RATE", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL1", "Z_AMOUNT", "Expense Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL1", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL1", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)


            AddTables("Z_OBILL2", "Expenses allocation ", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_OBILL2", "Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL2", "Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_OBILL2", "Z_CODE", "Expenses Code", SAPbobsCOM.BoFieldTypes.db_Alpha, )
            AddFields("Z_OBILL2", "Z_NAME", "Expenses Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_OBILL2", "Z_GLACC", "Deduction G/L Account", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_OBILL2", "Z_AMOUNT", "Expense Amount", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL2", "Z_TotalSq", "Total Sq.Meter", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL2", "Z_Rate", "Expense Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            AddTables("Z_AQEVAL", "Property Evaluation Details", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_AQEVAL", "Z_ProCode", "Property Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AQEVAL", "Z_ProName", "Property Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_EVL_ID", "Evaluation Serial Number", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_AQEVAL", "Z_EV_DATE", "Evaluation Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_AQEVAL", "Z_OWNERCODE", "Owner Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_AQEVAL", "Z_OWNER", "Owner", SAPbobsCOM.BoFieldTypes.db_Alpha, , 250)
            AddFields("Z_AQEVAL", "Z_OWNRTEL", "Owner Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_OWNRGSM", "Owner Mobile", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_OWNRPOBOX", "Owner P.O.Box", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_OWNRADD", "Owner Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            ' AddTables("Z_AQEVAL", "Property Management Type1", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_AQEVAL", "Z_AQTYP", "Property Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQBOND", "Property Title No", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQCITY", "Property City", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQZONE", "Pro.Zone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQSTREET", "Pro.Street", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQLSNC", "Pro.License No.", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQLSNC_DATE", "Pro.License Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_AQEVAL", "Z_AQNEARST", "Property Nearest Place", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQAGE", "Pro.Age", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_AQEVAL", "Z_AQBUILD", "Pro.Building Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQQUALITY", "Pro.Quality", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_USRENTNAME", "Pro.Tenant Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_USRMNTHRNT", "Pro.Monthly Rent Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_AQEVAL", "Z_RENTTYP", "Rent Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_MRKTLAND", "Land Market Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_AQEVAL", "Z_MRKTZFOOT", "Market Price of Foot", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_AQEVAL", "Z_BUILTCOST", "Building Cost", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_AQEVAL", "Z_MOREVALS", "Additional Cost the Building", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Rate)
            AddFields("Z_AQEVAL", "Z_AQAREA", "Property Area(Sqr Foot)", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_AQAREAMT", "Property Area(Sqr.Mtr)", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_BLDAREAFT", "Built Area (Sqr Foot)", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_FORCPER", "Force Selling Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_MRKTBLDMT", "Land's Market Price", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_BUILDAREA", "Building Area(Sqr Mtr)", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)

            'AddTables("Z_AQEVAL", "Property Management Type2", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_AQEVAL", "Z_REQ", "Person Requested Evaluation", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AQEVAL", "Z_RQTEL", "Person Telephone", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_RQGSM", "Person Mobile", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_RQADD", "Person Address", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AQEVAL", "Z_RQPOBOX", "Person P.O.Box", SAPbobsCOM.BoFieldTypes.db_Alpha, , 15)
            AddFields("Z_AQEVAL", "Z_REQEMAIL", "Person Email", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_REQFAX", "Person Fax", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_AQDESC", "Property Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_AQNOTES", "Notes", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_AQFACILITY", "Pro.Facilities", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_MOREFAC", "More Facilities", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_USEINFO", "Useful Information", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_AQEVAL", "Z_RENTERNAME", "Tenant Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_CHKRNAME", "Inspector Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AQEVAL", "Z_USERNAME", "User Name(of the system)", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_RCHKR", "Re-Checker Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 50)
            AddFields("Z_AQEVAL", "Z_EVCHRG", "Eval.Value(To Inv.Customer)", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_INVNO", "Invoice No", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_AQEVAL", "Z_DBRCPER", "Depreciation Ratio", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_AQEVAL", "Z_AC_ID", "Account No of Person", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_AQEVAL", "Z_NOTES", "Notes", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_AQEVAL", "Z_AQZNOTES", "Other Notes", SAPbobsCOM.BoFieldTypes.db_Memo)

            AddTables("Z_AQEVAL1", "Evaluation Attachments", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_AQEVAL1", "FileName", "File Name", SAPbobsCOM.BoFieldTypes.db_Memo)
            AddFields("Z_AQEVAL1", "AttDate", "Attachment Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_AQEVAL1", "AttName", "File Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_CONINS", "Contract Installments", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_CONINS", "Z_ConId", "Contract ID", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONINS", "Z_ConNo", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_CONINS", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONINS", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_CONINS", "Z_StartDate1", "Installment Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONINS", "Z_EndDate1", " Installment End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_CONINS", "Z_NoofMonths", "Number of Months", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONINS", "Z_AnnualRent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONINS", "Z_MonthRent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_CONINS", "Z_Month", "Month", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONINS", "Z_Year", "Year", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_CONINS", "Z_Amount", "Monthly Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_CONINS", "Z_Status", "Paid Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")


            addField("OUSR", "Z_Install", "Installment Authorization", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            '---- User Defined Object's

            'Phase II Changes 02-04-2014

            AddFields("Z_PROPUNIT", "Z_ExtProNo", "Extended Property Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_CONTRACT", "Z_ExtProNo", "Extended Property Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_CONTRACT_OWN", "Z_ExtProNo", "Extended Property Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 8)
            AddFields("Z_CONTRACT", "Z_DPDate", "DownPayment Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddFields("Z_CONTRACT", "Z_SDPDate", "Security DownPayment Date", SAPbobsCOM.BoFieldTypes.db_Date)

            AddTables("Z_OPRFA", "Property Facilitiy Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_OPRFA", "Z_CODE", "Facility Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPRFA", "Z_DESC", "Facility Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_OPRUFA", "Property Unit Facilitiy Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)
            AddFields("Z_OPRUFA", "Z_CODE", "Facility Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OPRUFA", "Z_DESC", "Facility Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)


            AddTables("Z_PROP2", "Property Facilities", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROP2", "Z_CODE", "Facility Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROP2", "Z_DESC", "Facility Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROP2", "Z_VALUE", "Value", SAPbobsCOM.BoFieldTypes.db_Memo)


          

            AddTables("Z_PROPUNIT2", "Property Unit Facilities", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROPUNIT2", "Z_CODE", "Facility Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT2", "Z_DESC", "Facility Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)
            AddFields("Z_PROPUNIT2", "Z_VALUE", "Value", SAPbobsCOM.BoFieldTypes.db_Memo)


            AddTables("Z_PROP3", "Property Fixed Asset", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROP3", "Z_CODE", "Asset Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROP3", "Z_DESC", "Asset Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_PROPUNIT3", "Property Unit Fixed Asset", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
            AddFields("Z_PROPUNIT3", "Z_CODE", "Asset Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_PROPUNIT3", "Z_DESC", "Asset Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            AddTables("Z_LogDetail", "Log file details", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            AddFields("Z_LOGDETAIL", "Z_LOG_PATH", "Log File Path", SAPbobsCOM.BoFieldTypes.db_Alpha, , 200)

            addField("@Z_CONTRACT", "Z_SPLIT", "Expenses in Separte Invoice", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")

            AddFields("Z_CONTRACT", "Z_CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_CONTRACT", "Z_NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)

            AddFields("Z_CONTRACT", "Z_CurApprover1", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_CONTRACT", "Z_NxtApprover1", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)

            addField("@Z_CONTRACT", "Z_AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_CONTRACT", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)

            addField("@Z_CONTRACT", "Z_AppRequired1", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_CONTRACT", "Z_AppReqDate1", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)

            addField("@Z_CONTRACT", "Z_RentType", "Rental Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "D,W,M,Q,S,A", "Daily,Weekly,Monthly,Quarterly,Semi Annual,Annual", "M")
            AddFields("Z_CONTRACT", "Z_NoofDays", "Number of Days", SAPbobsCOM.BoFieldTypes.db_Numeric)

            AddTables("Z_TCONTRACT", "Termination Contracts", SAPbobsCOM.BoUTBTableType.bott_Document)
            AddFields("Z_TCONTRACT", "Z_UnitCode", "Unit Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_TCONTRACT", "Z_Desc", "Unit Description", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TCONTRACT", "Z_ContDate", "Contract date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_TCONTRACT", "Z_StartDate", "Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_TCONTRACT", "Z_EndDate", "End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            addField("@Z_TCONTRACT", "Z_Status", "Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 3, SAPbobsCOM.BoFldSubTypes.st_Address, "O,A,R", "Open,Approved,Rejected", "O")
            AddFields("Z_TCONTRACT", "Z_TenCode", "Tenant Code", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_TCONTRACT", "Z_TenName", "Tenant Name", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TCONTRACT", "Z_Annualrent", "Annual Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_TCONTRACT", "Z_Deposit", "Security Deposit", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_TCONTRACT", "Z_Monthly", "Monthly Rent", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            addField("@Z_TCONTRACT", "Z_IsCommission", "Commission Paid", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            AddFields("Z_TCONTRACT", "Z_ChgMonth", "Number of Months", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_TCONTRACT", "Z_ChgAmt", "Termination Charges", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_TCONTRACT", "Z_Period", "Termination Application Period", SAPbobsCOM.BoFieldTypes.db_Numeric)
            AddFields("Z_TCONTRACT", "Z_TermDate", "Termination Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_TCONTRACT", "Z_Comm", "Commission Rate", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Price)
            AddFields("Z_TCONTRACT", "Z_ConNo", "Contract Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            addField("@Z_TCONTRACT", "Z_ProType", "Property Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "A,T", "Owned,Third Party", "T")
            AddFields("Z_TCONTRACT", "Z_DocNo", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)
            AddFields("Z_TCONTRACT", "Z_CntNo", "Contract Number with Sequance", SAPbobsCOM.BoFieldTypes.db_Alpha, , 20)
            AddFields("Z_TCONTRACT", "Z_CurApprover", "Current Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            AddFields("Z_TCONTRACT", "Z_NxtApprover", "Next Approver", SAPbobsCOM.BoFieldTypes.db_Alpha, , 40)
            addField("@Z_TCONTRACT", "Z_AppRequired", "Approval Required", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "Y")
            AddFields("Z_TCONTRACT", "Z_AppReqDate", "Required Date", SAPbobsCOM.BoFieldTypes.db_Date)

            addField("@Z_CONTRACT", "Z_ConAppStatus", "Contract Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")
            addField("@Z_CONTRACT", "Z_TerAppStatus", "Termination Approved Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "P,A,R", "Pending,Approved,Rejected", "P")

            addField("OACT", "Z_PostType", "Postable Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "S,D,N", "Security Deposit,DownPayment,Normal", "N")

            addField("@Z_CONINS", "Z_Manual", "Manual Change", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "Y,N", "Yes,No", "N")
            addField("@Z_CONINS", "Z_RentType", "Rental Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, SAPbobsCOM.BoFldSubTypes.st_Address, "D,W,M,Q,S,A", "Daily,Weekly,Monthly,Quarterly,Semi Annual,Annual", "M")
            AddFields("Z_OBILL", "Z_StartDate1", "Installment Start Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OBILL", "Z_EndDate1", " Installment End Date", SAPbobsCOM.BoFieldTypes.db_Date)
            AddFields("Z_OBILL", "Z_InsRef", " Installment Reference", SAPbobsCOM.BoFieldTypes.db_Alpha, , 30)
            AddFields("Z_OBILL", "RenType", "Rental Type", SAPbobsCOM.BoFieldTypes.db_Alpha, , 100)

            AddFields("Z_CONTRACT", "Z_MRent", "Monthly Rental", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            AddFields("Z_OBILL", "Z_MRent", "Monthly Rental", SAPbobsCOM.BoFieldTypes.db_Float, , , SAPbobsCOM.BoFldSubTypes.st_Sum)
            CreateUDO()

        Catch ex As Exception
            Throw ex
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub
    Public Sub CreateUDO()
        Try
            AddUDO("Z_TCONTRACT", "Termination Contracts", "Z_TCONTRACT", "DocEntry", "U_Z_UnitCode", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_APHIS", "Approval History", "Z_APHIS", "DocEntry", "U_Z_DocEntry", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_OAPPT", "Approval Template", "Z_OAPPT", "DocEntry", "U_Z_Code", "Z_APPT2", , SAPbobsCOM.BoUDOObjType.boud_Document)
            UDOPropertyUnitType("Z_UnitType", "Property Unit Type", "Z_UnitType", 1, "U_Z_UnitType")
            UDOPropertyFacility("Z_OPRFA", "Property Facility ", "Z_OPRFA", 1, "U_Z_CODE")
            UDOPropertyFacility("Z_OPRUFA", "Property Unit Facility ", "Z_OPRUFA", 1, "U_Z_CODE")
            UDOPriceList("Z_PRICELIST", "Pricelist master", "Z_PRICELIST", 1, "U_Z_PrlNam")
            AddUDO("Z_PROP", "Property Data", "Z_PROP", "DocEntry", "U_Z_Code", "Z_PROP1", "Z_PROP2", SAPbobsCOM.BoUDOObjType.boud_Document, "Z_PROP3")
            AddUDO("Z_PROPUNIT", "Property Unit Details", "Z_PROPUNIT", "DocEntry", "U_Z_Desc", "Z_PROPUNIT1", "Z_PROPUNIT2", SAPbobsCOM.BoUDOObjType.boud_Document, "Z_PROPUNIT3")
            AddUDO("Z_RESER", "Reservation", "Z_RESER", "DocEntry", "U_Z_CardCode", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_CONTRACT", "Contracts", "Z_CONTRACT", "DocEntry", "U_Z_TenCode", "Z_CONTRACT1", "Z_CONTRACT2", SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_INSURANCE", "Insurance Details", "Z_INSURANCE", "DocEntry", "U_Z_TenCode", , , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_AQEVAL", "Property Evalution", "Z_AQEVAL", "DocEntry", "U_Z_ProCode", "Z_AQEVAL1", , SAPbobsCOM.BoUDOObjType.boud_Document)
            AddUDO("Z_CONTRACT_OWN", "OwnerContracts", "Z_CONTRACT_OWN", "DocEntry", "U_Z_TenCode", "Z_CONTRACT_OWN1", "Z_CONTRACT_OWN2", SAPbobsCOM.BoUDOObjType.boud_Document)


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

    Public Function UDOPropertyFacility(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_CODE"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_CODE"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_DESC"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_DESC"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Function UDOPropertyUnitType(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_UnitType"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_UnitType"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Status"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Status"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_FrgnName"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_FrgnName"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function

    Public Function UDOPriceList(ByVal strUDO As String, _
                            ByVal strDesc As String, _
                                ByVal strTable As String, _
                                    ByVal intFind As Integer, _
                                        Optional ByVal strCode As String = "", _
                                            Optional ByVal strName As String = "") As Boolean
        Dim oUserObjects As SAPbobsCOM.UserObjectsMD
        Dim lngRet As Long
        Try
            oUserObjects = oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            If oUserObjects.GetByKey(strUDO) = 0 Then
                oUserObjects.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
                oUserObjects.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.CanFind = SAPbobsCOM.BoYesNoEnum.tYES


                oUserObjects.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.LogTableName = ""
                oUserObjects.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.ExtensionName = ""

                oUserObjects.FormColumns.FormColumnAlias = "Code"
                oUserObjects.FormColumns.FormColumnDescription = "Code"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "Name"
                oUserObjects.FormColumns.FormColumnDescription = "Name"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "DocEntry"
                oUserObjects.FormColumns.FormColumnDescription = "DocEntry"
                oUserObjects.FormColumns.Add()

                oUserObjects.FormColumns.FormColumnAlias = "U_Z_PrjCode"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_PrjCode"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_PrlNam"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_PrlNam"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Price"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Price"
                oUserObjects.FormColumns.Add()
                oUserObjects.FormColumns.FormColumnAlias = "U_Z_Status"
                oUserObjects.FormColumns.FormColumnDescription = "U_Z_Status"
                oUserObjects.FormColumns.Add()
                oUserObjects.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO
                oUserObjects.Code = strUDO
                oUserObjects.Name = strDesc
                oUserObjects.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData
                oUserObjects.TableName = strTable

                If oUserObjects.CanFind = 1 Then
                    oUserObjects.FindColumns.ColumnAlias = strCode
                End If

                If oUserObjects.Add() <> 0 Then
                    oApplication.Utilities.Message(oApplication.Company.GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                    oUserObjects = Nothing
                    Return False
                End If
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjects)
                oUserObjects = Nothing
                Return True
            End If
        Catch ex As Exception
            Throw ex
        Finally
            oUserObjects = Nothing
            GC.Collect()
        End Try
    End Function
End Class
