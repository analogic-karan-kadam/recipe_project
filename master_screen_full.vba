
Private Function InitRecipeTagGroup() As TagGroup
    Dim recipeTagGroup As TagGroup
    If recipeTagGroup Is Nothing Then
        Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)

        ' --- SCADA Data ---
        recipeTagGroup.Add "recipe\SCADA_DATA_Batch_No"
        recipeTagGroup.Add "recipe\status"
        recipeTagGroup.Add "recipe\disable"
        recipeTagGroup.Add "recipe\recipeName"
         recipeTagGroup.Add "recipe\comment"
         recipeTagGroup.Add "recipe\Display"

        ' --- Wax Heat SP TSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_0"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_1"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_2"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_3"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_4"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_5"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_6"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_7"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_8"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_9"

        ' --- Wax Heat ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_0"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_1"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_2"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_3"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_4"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_5"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_6"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_7"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_8"
        recipeTagGroup.Add "recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_9"

        ' --- Product Heating MFG Heat TSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_0"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_1"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_2"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_3"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_4"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_5"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_6"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_7"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_8"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_TSP_Cyc_9"

        ' --- Product Heating MFG Heat ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_0"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_1"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_2"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_3"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_4"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_5"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_6"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_7"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_8"
        recipeTagGroup.Add "recipe\Product_Heating_MFG_Heat_ONSP_Cyc_9"

        ' --- Product Cooling MFG Cool TSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_0"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_1"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_2"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_3"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_4"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_5"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_6"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_7"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_8"
        recipeTagGroup.Add "recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_9"

        ' --- Water Heating Water Heat TSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_0"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_1"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_2"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_3"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_4"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_5"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_6"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_7"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_8"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_TSP_Cyc_9"

        ' --- Water Heating Water Heat ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_0"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_1"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_2"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_3"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_4"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_5"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_6"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_7"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_8"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Heat_ONSP_Cyc_9"

        ' --- Anchor Motor RPM SP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_0"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_1"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_2"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_3"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_4"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_5"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_6"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_7"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_8"
        recipeTagGroup.Add "recipe\Motor_Anchor_Mtr_RPM_SP_CYC_9"

        ' --- Anchor Motor ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_0"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_1"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_2"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_3"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_4"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_5"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_6"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_7"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_8"
        recipeTagGroup.Add "recipe\Anchor_Motor_ONSP_Cyc_9"

        ' --- Homogeniser Motor RPM SP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_0"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_1"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_2"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_3"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_4"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_5"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_6"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_7"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_8"
        recipeTagGroup.Add "recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_9"

        ' --- Homogeniser Motor ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_0"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_1"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_2"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_3"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_4"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_5"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_6"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_7"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_8"
        recipeTagGroup.Add "recipe\Homogeniser_Motor_ONSP_Cyc_9"

        ' --- Water Stirrer Motor RPM SP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_0"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_1"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_2"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_3"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_4"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_5"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_6"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_7"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_8"
        recipeTagGroup.Add "recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_9"

        ' --- Water Stirrer ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_0"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_1"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_2"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_3"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_4"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_5"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_6"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_7"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_8"
        recipeTagGroup.Add "recipe\Water_Heating_Water_Stirr_ONSP_CyC_9"

        ' --- Wax Stirrer Motor RPM SP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_0"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_1"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_2"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_3"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_4"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_5"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_6"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_7"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_8"
        recipeTagGroup.Add "recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_9"

        ' --- Wax Stirrer ONSP (Cycles 0-9) ---
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_0"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_1"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_2"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_3"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_4"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_5"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_6"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_7"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_8"
        recipeTagGroup.Add "recipe\Wax_Stirrer_ONSP_CYC_9"

    End If
    Set InitRecipeTagGroup = recipeTagGroup
End Function
Private Sub CreateRecipeBtn_Press()
Dim recipeTagGroup As TagGroup
If recipeTagGroup Is Nothing Then
    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)
         
        recipeTagGroup.Add "recipe\recipeName"
        recipeTagGroup.Add "recipe\SCADA_DATA_Batch_No"
        
         
   End If
   
    recipeTagGroup.Item("recipe\recipeName").Value = ""
    recipeTagGroup.Item("recipe\SCADA_DATA_Batch_No").Value = ""
End Sub

 

Private Sub Display_AnimationStart()
'----------------load recipe table data on start
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    LoadRecipeData
    
End Sub
Private Function FormatColumnHeader(ByVal colName As String) As String

    Dim i As Integer
    Dim result As String
    Dim ch As String

    result = ""

    For i = 1 To Len(colName)

        ch = Mid(colName, i, 1)

        If i > 1 Then
            If ch Like "[A-Z]" Then
                result = result & " "
            End If
        End If

        result = result & ch

    Next i

    FormatColumnHeader = UCase(result)

End Function
'------------------- save sql to text file
Public Sub SaveTextToFile(ByVal filePath As String, ByVal textData As String)

    On Error GoTo ErrHandler

    Dim fileNum As Integer

    fileNum = FreeFile

    Open filePath For Output As #fileNum
        Print #fileNum, textData
    Close #fileNum

    Exit Sub

ErrHandler:
    MsgBox "Error saving file: " & Err.Description
End Sub
'----------------load recipe data funtion
Public Sub LoadRecipeData()
    'On Error GoTo ErrHandler
    Dim conn As ADODB.Connection
    Dim Rs As ADODB.Recordset
    Dim sql As String
    Dim r As Long
    Dim c As Long
    Set conn = New ADODB.Connection
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"
        sql = "SELECT recipeName,status , createdBy , createdDate ,approvedBy , approvedDate , comment , updatedBy, updatedDate FROM recipeMaster2"
        Set Rs = conn.Execute(sql)
        recipeMasterTable.Clear
    
       
        recipeMasterTable.Rows = 1
        recipeMasterTable.Cols = Rs.Fields.Count
        recipeMasterTable.ColWidth(0) = 2500   ' Recipe Name
        recipeMasterTable.ColWidth(1) = 4500   ' status
        recipeMasterTable.ColWidth(2) = 2500   ' createdBy
        recipeMasterTable.ColWidth(3) = 2500   ' createdDate
        recipeMasterTable.ColWidth(4) = 2500   ' approvedBy
        recipeMasterTable.ColWidth(5) = 2500   ' approvedDate
        recipeMasterTable.ColWidth(6) = 2500   ' comment
        recipeMasterTable.ColWidth(7) = 2500   ' updatedBy
        recipeMasterTable.ColWidth(8) = 2500   ' updatedDate
    ' Column Names
        For c = 0 To recipeMasterTable.Cols - 1
            recipeMasterTable.TextMatrix(0, c) = _
            FormatColumnHeader(Rs.Fields(c).Name)
        Next c
        
        '==========================
        ' Grid Formatting
        '==========================
        
        With recipeMasterTable
        
        ' Entire grid font size
        .Font.Name = "Arial"
        .Font.Size = 13          ' Change this for all rows
        .Font.Bold = False
        
        End With
        
        ' Header formatting
        For c = 0 To recipeMasterTable.Cols - 1
        
        recipeMasterTable.Row = 0
        recipeMasterTable.Col = c
        
        recipeMasterTable.CellFontBold = True
        
        ' Header alignment
        recipeMasterTable.CellAlignment = flexAlignCenterCenter
        
        Next c
        
        ' Data alignment (left)
        Dim t As Long
        
        For t = 1 To recipeMasterTable.Rows - 1
        
        For c = 0 To recipeMasterTable.Cols - 1
        
         recipeMasterTable.Row = t
         recipeMasterTable.Col = c
        
         recipeMasterTable.CellAlignment = flexAlignLeftCenter
        
        Next c
        
        Next t
        
    If Rs.EOF Then
        LogDiagnosticsMessage _
            "No records found in RecipeMaster2Screen"
        Rs.Close
        conn.Close
    Exit Sub

    End If
    
    ' Data Rows
    Do Until Rs.EOF
        recipeMasterTable.Rows = recipeMasterTable.Rows + 1

        r = recipeMasterTable.Rows - 1
        For c = 0 To recipeMasterTable.Cols - 1
            '-- RecipeMasterTable.TextMatrix(r, c) = _
                rs.Fields(c).Value & ""
             recipeMasterTable.TextMatrix(r, c) = GetDisplayValue(Rs.Fields(c).Name, Rs.Fields(c).Value & "")
        Next c
        Rs.MoveNext
    Loop
    Rs.Close
    conn.Close
    Set Rs = Nothing
    Set conn = Nothing
    LogDiagnosticsMessage _
        "Recipe data loaded successfully"
    Exit Sub

ErrHandler:
    LogDiagnosticsMessage _
        "LoadRecipe Error : " & Err.Description
    MsgBox _
        "LoadRecipe Error : " & Err.Description

End Sub


Private Sub recipeMasterTable_Click()

End Sub

Private Sub RefreshTableBtn_Released()
    LoadRecipeData
End Sub
'-------------------------------------get display value for 0 and 1
Private Function GetDisplayValue( _
    ByVal columnName As String, _
    ByVal rawValue As Variant) As String

    Select Case LCase(columnName)

        Case "status"

            Select Case CLng(rawValue)
                Case 0
                    GetDisplayValue = "Not Approved"
                Case 1
                    GetDisplayValue = "Approved"
                Case 2
                    GetDisplayValue = "Pending Approval after Recipe Update"
                Case 3
                    GetDisplayValue = "Disabled"
                Case Else
                    GetDisplayValue = "Unknown"
            End Select
            
        Case Else

            GetDisplayValue = CStr(rawValue)

    End Select

End Function
Private Sub UpdateRecipeVersion()
    On Error GoTo ErrHandler
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim recipeName As String

    If recipeMasterTable.Row <= 0 Then

        MsgBox "Please select a recipe"

        Exit Sub

    End If

    recipeName = recipeMasterTable.TextMatrix( _
                    recipeMasterTable.Row, 0)

    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "UPDATE RecipeMaster2 " & _
          "SET version = version + 1 " & _
          "WHERE recipeName = '" & recipeName & "'"

    LogDiagnosticsMessage sql

    conn.Execute sql

    MsgBox "Recipe version updated successfully"
    LoadRecipeData
    conn.Close
    Set conn = Nothing

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _
        "Version update failed : " & Err.Description

    MsgBox _
        "Version update failed : " & Err.Description

End Sub

Private Sub ViewRecipeBtn1_Press()
   Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    
    recipeTagGroup.Item("recipe\recipeName").Value = recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)
   'MsgBox "selected recipe name is : " & recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)
   
   If (StrComp(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0), "recipeName", vbBinaryCompare) = 0) Or (StrComp(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0), "RECIPE NAME", vbBinaryCompare) = 0) Then
        MsgBox ("Please select the recipe")
   Else
    
    SaveQueryLog "Before Recipe View"
    
    Call renderRecipeParameters(recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0))
     LogDiagnosticsMessage (" getting Display RECIPE_PARAMETERS_WAXopt")
     
      SaveQueryLog "after Recipe View"
      
    ExecuteCommand "Display RECIPE_PARAMETERS_WAXopt"
    LogDiagnosticsMessage (" getting Display RECIPE_PARAMETERS_WAX")
    'ShowDisplay "RECIPE_PARAMETERS_WAX"
   End If
   
End Sub

Private Sub ViewRecipeBtn1_Released()
  
   
End Sub

Public Sub SaveQueryLog(ByVal logText As String)

    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim sqlQuery As String

    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sqlQuery = "INSERT INTO checkLogTime ([condition], [datetime]) " & _
               "VALUES ('" & Replace(logText, "'", "''") & "', GETDATE())"

    conn.Execute sqlQuery

CleanUp:
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error saving log: " & Err.Description, vbExclamation
    Resume CleanUp

End Sub
Function renderRecipeParameters(ByVal recipeName As String)
   ' MsgBox ("Render function for master screen " & recipeName)
    Dim conn
    Dim Rs
    Dim Field, xyz
    Dim ServerName, sqlUsername, sqlPassword, Table, Database As String
    '--------------------------tag group loaded
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    
    If IsNull(recipeName) Or Trim(recipeName) = "" Or recipeName = "RECIPE NAME" Or recipeName = "recipeName" Then
        MsgBox "Please select a recipe", vbExclamation, "Missing Recipe"
    Exit Function
    End If
    
    Set conn = New ADODB.Connection
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    ' ==============================================================
    ' FETCH RECIPE ROW
    ' ==============================================================
    StrQuery = "SELECT * from recipeMaster2 " & _
               "WHERE recipeName = '" & recipeName & "'"
               
               
    'SaveQueryLog "Before Recipe View"
    
    Set Rs = conn.Execute(StrQuery)
    
    'SaveQueryLog "after Recipe view"
    
    
    'LogDiagnosticsMessage "Execution Time: " & _
                Format((Timer - StartTime) * 1000, "0.00") & " ms"
                
    
    'MsgBox "Recipe tag found: " & Rs, vbExclamation
    ' ==============================================================
    ' SET TAG VALUES FROM RECORDSET
    ' ==============================================================
    Dim displayValue As Integer
    Dim i As Integer
    
    displayValue = recipeTagGroup.Item("recipe\Display").Value


    
    If Rs.EOF Then
        MsgBox "Recipe not found: " & recipeName, vbExclamation
    Else
    Do While Not Rs.EOF
    
    Select Case displayValue
    
    '==========================
    ' Display = 0 (Wax Heating)
    '==========================
    Case 0
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i).Value = _
        Rs("PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_" & i).Value = _
        Rs("PID_Parameter_Wax_Heat_ONSP_Cyc_" & i)
    Next i
    
    '==========================
    ' Display = 1 (Water Heating)
    '==========================
    Case 1
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_TSP_Cyc_" & i).Value = _
        Rs("Water_Heating_Water_Heat_TSP_Cyc_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_ONSP_Cyc_" & i).Value = _
        Rs("Water_Heating_Water_Heat_ONSP_Cyc_" & i)
    Next i
    
    '==========================
    ' Display = 2 (Product Heating/Cooling)
    '==========================
    Case 2
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_TSP_Cyc_" & i).Value = _
        Rs("Product_Heating_MFG_Heat_TSP_Cyc_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_" & i).Value = _
        Rs("Product_Heating_MFG_Heat_ONSP_Cyc_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i).Value = _
        Rs("Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i)
    Next i
    
    '==========================
    ' Display = 3 (Anchor Motor)
    '==========================
    Case 3
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_" & i).Value = _
        Rs("Motor_Anchor_Mtr_RPM_SP_CYC_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Anchor_Motor_ONSP_Cyc_" & i).Value = _
        Rs("Anchor_Motor_ONSP_Cyc_" & i)
    Next i
    
    '==========================
    ' Display = 4 (Homogeniser)
    '==========================
    Case 4
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i).Value = _
        Rs("Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Homogeniser_Motor_ONSP_Cyc_" & i).Value = _
        Rs("Homogeniser_Motor_ONSP_Cyc_" & i)
    Next i
    
    '==========================
    ' Display = 5 (Water Stirrer)
    '==========================
    Case 5
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i).Value = _
        Rs("Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Water_Heating_Water_Stirr_ONSP_CyC_" & i).Value = _
        Rs("Water_Heating_Water_Stirr_ONSP_CyC_" & i)
    Next i
    
    '==========================
    ' Display = 6 (Wax Stirrer)
    '==========================
    Case 6
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i).Value = _
        Rs("Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i)
    Next i
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\Wax_Stirrer_ONSP_CYC_" & i).Value = _
        Rs("Wax_Stirrer_ONSP_CYC_" & i)
    Next i
    
    End Select
    
    ' Always update control tags
    recipeTagGroup.Item("recipe\status").Value = Rs("status").Value
    recipeTagGroup.Item("recipe\disable").Value = Rs("disable").Value
    
    Rs.MoveNext
    
    Loop
    End If

    ' --- Recipe name written to tag regardless of DB result ---
    recipeTagGroup.Item("recipe\recipeName").Value = recipeName

    ' ==============================================================
    ' CLEANUP
    ' ==============================================================
    Rs.Close
    conn.Close
    Set Rs = Nothing
    Set conn = Nothing

End Function

Private Sub DeleteRecipeBtn_Released()
   'MsgBox "delete is clicked "
    Dim recipeTagGroup As TagGroup
   
   If recipeTagGroup Is Nothing Then
    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)
        recipeTagGroup.Add "recipe\recipeName"
   End If
   'MsgBox "selected recipe name is : " & RecipeMasterTable.TextMatrix(RecipeMasterTable.Row, 1)
   recipeTagGroup.Item("recipe\recipeName").Value = recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)
    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim sql As String
    Dim recipeName As String

    
    recipeName = recipeTagGroup.Item("recipe\recipeName").Value
    
    If IsNull(recipeName) Or Trim(recipeName) = "" Or recipeName = "recipeName" Then
        MsgBox "Please select a recipe", vbExclamation, "Missing Recipe"
        Exit Sub
    End If

    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "DELETE FROM RecipeMaster2 WHERE recipeName = '" & recipeName & "'"

    conn.Execute sql
    
    MsgBox "Recipe Deleted Successfully"
    Call LoadRecipeData
    conn.Close

    Set conn = Nothing

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _
        "Delete failed: " & Err.Description

    MsgBox _
        "Delete failed: " & Err.Description

End Sub



 
