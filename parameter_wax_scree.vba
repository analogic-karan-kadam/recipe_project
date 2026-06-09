
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
'getting tag value by name
Private Function GetTagVal(tagName As String) As Variant
     Dim recipeTagGroup As TagGroup
     Set recipeTagGroup = InitRecipeTagGroup
    Dim t As Tag
    Set t = recipeTagGroup.Item(tagName)
    GetTagVal = t.Value
End Function

Private Sub approveBtn_Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)
    Dim conn As ADODB.Connection
    Dim sql As String
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    
       Dim strRecipeName   As String
    strRecipeName = GetTagVal("recipe\recipeName")
     MsgBox "you are approving " & strRecipeName & " recipe"
    
    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "UPDATE RecipeMaster2 " & _
          "SET Status = 1, " & _
          "[Comment] = '" & Replace(PerformerComment, "'", "''") & "', " & _
          "approvedBy = '" & Replace(PerformerUser, "'", "''") & "', " & _
          "approvedDate = GETDATE() " & _
          "WHERE recipeName = '" & Replace(recipeNameInput.Value, "'", "''") & "'"

    conn.Execute sql
    
    conn.Close
    Set conn = Nothing

End Sub


 

Private Function CheckDuplicate() As Boolean

    Dim conn As ADODB.Connection
    Dim sql As String

    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "UPDATE RecipeMaster2 " & _
          "SET Status = 0, " & _
          "Comment = '" & Replace(PerformerComment, "'", "''") & "' " & _
          "WHERE recipeName = '" & Replace(recipeNameInput.Value, "'", "''") & "'"

    conn.Execute sql

    conn.Close
    Set conn = Nothing

End Function
Function renderRecipeParameters(ByVal recipeName As String)
    MsgBox ("Render function for " & recipeName)
    Dim conn
    Dim Rs
    Dim Field, xyz
    Dim ServerName, sqlUsername, sqlPassword, Table, Database As String
    '--------------------------tag group loaded
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
   
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

    Set Rs = conn.Execute(StrQuery)
    'MsgBox "Recipe tag found: " & Rs, vbExclamation
    ' ==============================================================
    ' SET TAG VALUES FROM RECORDSET
    ' ==============================================================
    If Rs.EOF Then
        MsgBox "Recipe not found: " & recipeName, vbExclamation
    Else
        Do While Not Rs.EOF
            ' --- Wax Heat SP TSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i).Value = _
                    Rs("PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i)
            Next i

            ' --- Wax Heat ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_" & i).Value = _
                    Rs("PID_Parameter_Wax_Heat_ONSP_Cyc_" & i)
            Next i

            ' --- Product Heating TSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_TSP_Cyc_" & i).Value = _
                    Rs("Product_Heating_MFG_Heat_TSP_Cyc_" & i)
            Next i

            ' --- Product Heating ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_" & i).Value = _
                    Rs("Product_Heating_MFG_Heat_ONSP_Cyc_" & i)
            Next i

            ' --- Product Cooling TSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i).Value = _
                    Rs("Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i)
            Next i

            ' --- Water Heating TSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_TSP_Cyc_" & i).Value = _
                    Rs("Water_Heating_Water_Heat_TSP_Cyc_" & i)
            Next i

            ' --- Water Heating ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_ONSP_Cyc_" & i).Value = _
                    Rs("Water_Heating_Water_Heat_ONSP_Cyc_" & i)
            Next i

            ' --- Anchor Motor RPM SP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_" & i).Value = _
                    Rs("Motor_Anchor_Mtr_RPM_SP_CYC_" & i)
            Next i

            ' --- Anchor Motor ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Anchor_Motor_ONSP_Cyc_" & i).Value = _
                    Rs("Anchor_Motor_ONSP_Cyc_" & i)
            Next i

            ' --- Homogeniser Motor RPM SP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i).Value = _
                    Rs("Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i)
            Next i

            ' --- Homogeniser Motor ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Homogeniser_Motor_ONSP_Cyc_" & i).Value = _
                    Rs("Homogeniser_Motor_ONSP_Cyc_" & i)
            Next i

            ' --- Water Stirrer RPM SP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i).Value = _
                    Rs("Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i)
            Next i

            ' --- Water Stirrer ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Water_Heating_Water_Stirr_ONSP_CyC_" & i).Value = _
                    Rs("Water_Heating_Water_Stirr_ONSP_CyC_" & i)
            Next i

            ' --- Wax Stirrer RPM SP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i).Value = _
                    Rs("Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i)
            Next i

            ' --- Wax Stirrer ONSP ---
            For i = 0 To 9
                recipeTagGroup.Item("recipe\Wax_Stirrer_ONSP_CYC_" & i).Value = _
                    Rs("Wax_Stirrer_ONSP_CYC_" & i)
            Next i

            ' --- Control Tags ---
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

Private Sub Display_AnimationStart()
     Dim recipeTagGroup As TagGroup
     Set recipeTagGroup = InitRecipeTagGroup
     'MsgBox ("Recipe value is " & recipeTagGroup.Item("recipe\recipeName").Value)
End Sub

Private Sub rejectBtn__Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)

End Sub


Private Sub rejectBtn_Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)
       Dim conn As ADODB.Connection
    Dim sql As String
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    
       Dim strRecipeName   As String
    strRecipeName = GetTagVal("recipe\recipeName")
     MsgBox "you are rejecting " & strRecipeName & " recipe"
    
    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "UPDATE RecipeMaster2 " & _
      "SET Status = 0, " & _
      "[Comment] = '" & Replace(PerformerComment, "'", "''") & "', " & _
      "approvedBy = 'Null', " & _
      "approvedDate = GETDATE() " & _
      "WHERE recipeName = '" & Replace(recipeNameInput.Value, "'", "''") & "'"

    conn.Execute sql
    
    conn.Close
    Set conn = Nothing
End Sub

Private Sub saveToSqlBtn_Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)
    
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    
    If CheckDuplicate Then
        MsgBox "Recipe Already Exists Choose Differen Name"
        recipeNameInput.Value = ""
        Exit Sub
    End If
    
    'On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim sqlQuery As String
    
    ' --- 1. VALIDATION: Check if objects exist and have text ---
    ' Using Controls ensures VBA binds to the current display items
    If Trim(Me.recipeNameInput.Value) = "" Then
        MsgBox "Error: Please Enter Recipe Name", vbCritical, "Missing Data"
        Exit Sub
    End If

    ' --- 3. DATABASE CONNECTION ---
    Set conn = New ADODB.Connection
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"
        
        'Dim createdByName As String
        'createdByName = PerformerUser
     
           
    ' --- 4. BUILD THE SQL QUERY ---
    ' Explicitly added 'productName' back to match your original column structure if needed
    ' Added 'Me.' prefix to force FactoryTalk to look at the screen elements
    'sqlQuery = "INSERT INTO RecipeMaster2 (recipeName, " & _
        "TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_0, TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_1, " & _
        "TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_2, TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_3, " & _
        "TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_4, TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_5, " & _
        "TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_6, TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_7, " & _
        "TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_8, TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_9, " & _
        "TESTPID_Parameter_Wax_Heat_ONSP_Cyc_0, TESTPID_Parameter_Wax_Heat_ONSP_Cyc_1, " & _
        "TESTPID_Parameter_Wax_Heat_ONSP_Cyc_2, TESTPID_Parameter_Wax_Heat_ONSP_Cyc_3, " & _
        "TESTPID_Parameter_Wax_Heat_ONSP_Cyc_4, TESTPID_Parameter_Wax_Heat_ONSP_Cyc_5, " & _
        "TESTPID_Parameter_Wax_Heat_ONSP_Cyc_6, TESTPID_Parameter_Wax_Heat_ONSP_Cyc_7, " & _
        "TESTPID_Parameter_Wax_Heat_ONSP_Cyc_8, TESTPID_Parameter_Wax_Heat_ONSP_Cyc_9) " & _
        "VALUES (" & _
        "'" & recipeNameInput.Value & "', " & _
        NumericInput1.Value & ", " & NumericInput2.Value & ", " & _
       NumericInput3.Value & ", " & NumericInput4.Value & ", " & _
        NumericInput5.Value & ", " & NumericInput6.Value & ", " & _
        NumericInput7.Value & ", " & NumericInput8.Value & ", " & _
        NumericInput9.Value & ", " & NumericInput10.Value & ", " & _
       NumericInput11.Value & ", " & NumericInput11.Value & ", " & _
       NumericInput12.Value & ", " & NumericInput13.Value & ", " & _
       NumericInput14.Value & ", " & NumericInput15.Value & ", " & _
       NumericInput16.Value & ", " & NumericInput17.Value & ", " & _
       NumericInput18.Value & ", " & NumericInput19.Value & " " & _
        ")"
    'save to sql by tag values or hmi tags
    
    
    ' --- Dummy variables — replace with your actual sources ---
    Dim strRecipeName   As String
    strRecipeName = GetTagVal("recipe\recipeName")
     If strRecipeName = "" Then
        MsgBox "Please Enter Recipe Name or Press enter after entering name"
        recipeNameInput.Value = ""
        Exit Sub
    End If
    
    MsgBox "recipe name is from tag : " & strRecipeName
    Dim strBatchNumber  As String
    strBatchNumber = GetTagVal("recipe\SCADA_DATA_Batch_No")
    Dim strCreatedBy    As String
    strCreatedBy = PerformerUser

    ' ==============================================================
    ' COLUMN GROUPS
    ' ==============================================================

    Dim colsMeta As String
    colsMeta = "recipeName, batchNumber, createdBy, createdDate"

    Dim colsWaxHeatTSP As String
    colsWaxHeatTSP = "PID_Parameter_Wax_Heat_SP_TSP_Cyc_0, PID_Parameter_Wax_Heat_SP_TSP_Cyc_1, " & _
                     "PID_Parameter_Wax_Heat_SP_TSP_Cyc_2, PID_Parameter_Wax_Heat_SP_TSP_Cyc_3, " & _
                     "PID_Parameter_Wax_Heat_SP_TSP_Cyc_4, PID_Parameter_Wax_Heat_SP_TSP_Cyc_5, " & _
                     "PID_Parameter_Wax_Heat_SP_TSP_Cyc_6, PID_Parameter_Wax_Heat_SP_TSP_Cyc_7, " & _
                     "PID_Parameter_Wax_Heat_SP_TSP_Cyc_8, PID_Parameter_Wax_Heat_SP_TSP_Cyc_9"

    Dim colsWaxHeatONSP As String
    colsWaxHeatONSP = "PID_Parameter_Wax_Heat_ONSP_Cyc_0, PID_Parameter_Wax_Heat_ONSP_Cyc_1, " & _
                      "PID_Parameter_Wax_Heat_ONSP_Cyc_2, PID_Parameter_Wax_Heat_ONSP_Cyc_3, " & _
                      "PID_Parameter_Wax_Heat_ONSP_Cyc_4, PID_Parameter_Wax_Heat_ONSP_Cyc_5, " & _
                      "PID_Parameter_Wax_Heat_ONSP_Cyc_6, PID_Parameter_Wax_Heat_ONSP_Cyc_7, " & _
                      "PID_Parameter_Wax_Heat_ONSP_Cyc_8, PID_Parameter_Wax_Heat_ONSP_Cyc_9"

    Dim colsProdHeatTSP As String
    colsProdHeatTSP = "Product_Heating_MFG_Heat_TSP_Cyc_0, Product_Heating_MFG_Heat_TSP_Cyc_1, " & _
                      "Product_Heating_MFG_Heat_TSP_Cyc_2, Product_Heating_MFG_Heat_TSP_Cyc_3, " & _
                      "Product_Heating_MFG_Heat_TSP_Cyc_4, Product_Heating_MFG_Heat_TSP_Cyc_5, " & _
                      "Product_Heating_MFG_Heat_TSP_Cyc_6, Product_Heating_MFG_Heat_TSP_Cyc_7, " & _
                      "Product_Heating_MFG_Heat_TSP_Cyc_8, Product_Heating_MFG_Heat_TSP_Cyc_9"

    Dim colsProdHeatONSP As String
    colsProdHeatONSP = "Product_Heating_MFG_Heat_ONSP_Cyc_0, Product_Heating_MFG_Heat_ONSP_Cyc_1, " & _
                       "Product_Heating_MFG_Heat_ONSP_Cyc_2, Product_Heating_MFG_Heat_ONSP_Cyc_3, " & _
                       "Product_Heating_MFG_Heat_ONSP_Cyc_4, Product_Heating_MFG_Heat_ONSP_Cyc_5, " & _
                       "Product_Heating_MFG_Heat_ONSP_Cyc_6, Product_Heating_MFG_Heat_ONSP_Cyc_7, " & _
                       "Product_Heating_MFG_Heat_ONSP_Cyc_8, Product_Heating_MFG_Heat_ONSP_Cyc_9"

    Dim colsProdCoolTSP As String
    colsProdCoolTSP = "Product_Cooling_MFG_Cool_TSP_CycMW1168_0, Product_Cooling_MFG_Cool_TSP_CycMW1168_1, " & _
                      "Product_Cooling_MFG_Cool_TSP_CycMW1168_2, Product_Cooling_MFG_Cool_TSP_CycMW1168_3, " & _
                      "Product_Cooling_MFG_Cool_TSP_CycMW1168_4, Product_Cooling_MFG_Cool_TSP_CycMW1168_5, " & _
                      "Product_Cooling_MFG_Cool_TSP_CycMW1168_6, Product_Cooling_MFG_Cool_TSP_CycMW1168_7, " & _
                      "Product_Cooling_MFG_Cool_TSP_CycMW1168_8, Product_Cooling_MFG_Cool_TSP_CycMW1168_9"

    Dim colsWaterHeatTSP As String
    colsWaterHeatTSP = "Water_Heating_Water_Heat_TSP_Cyc_0, Water_Heating_Water_Heat_TSP_Cyc_1, " & _
                       "Water_Heating_Water_Heat_TSP_Cyc_2, Water_Heating_Water_Heat_TSP_Cyc_3, " & _
                       "Water_Heating_Water_Heat_TSP_Cyc_4, Water_Heating_Water_Heat_TSP_Cyc_5, " & _
                       "Water_Heating_Water_Heat_TSP_Cyc_6, Water_Heating_Water_Heat_TSP_Cyc_7, " & _
                       "Water_Heating_Water_Heat_TSP_Cyc_8, Water_Heating_Water_Heat_TSP_Cyc_9"

    Dim colsWaterHeatONSP As String
    colsWaterHeatONSP = "Water_Heating_Water_Heat_ONSP_Cyc_0, Water_Heating_Water_Heat_ONSP_Cyc_1, " & _
                        "Water_Heating_Water_Heat_ONSP_Cyc_2, Water_Heating_Water_Heat_ONSP_Cyc_3, " & _
                        "Water_Heating_Water_Heat_ONSP_Cyc_4, Water_Heating_Water_Heat_ONSP_Cyc_5, " & _
                        "Water_Heating_Water_Heat_ONSP_Cyc_6, Water_Heating_Water_Heat_ONSP_Cyc_7, " & _
                        "Water_Heating_Water_Heat_ONSP_Cyc_8, Water_Heating_Water_Heat_ONSP_Cyc_9"

    Dim colsAnchorRPM As String
    colsAnchorRPM = "Motor_Anchor_Mtr_RPM_SP_CYC_0, Motor_Anchor_Mtr_RPM_SP_CYC_1, " & _
                    "Motor_Anchor_Mtr_RPM_SP_CYC_2, Motor_Anchor_Mtr_RPM_SP_CYC_3, " & _
                    "Motor_Anchor_Mtr_RPM_SP_CYC_4, Motor_Anchor_Mtr_RPM_SP_CYC_5, " & _
                    "Motor_Anchor_Mtr_RPM_SP_CYC_6, Motor_Anchor_Mtr_RPM_SP_CYC_7, " & _
                    "Motor_Anchor_Mtr_RPM_SP_CYC_8, Motor_Anchor_Mtr_RPM_SP_CYC_9"

    Dim colsAnchorONSP As String
    colsAnchorONSP = "Anchor_Motor_ONSP_Cyc_0, Anchor_Motor_ONSP_Cyc_1, " & _
                     "Anchor_Motor_ONSP_Cyc_2, Anchor_Motor_ONSP_Cyc_3, " & _
                     "Anchor_Motor_ONSP_Cyc_4, Anchor_Motor_ONSP_Cyc_5, " & _
                     "Anchor_Motor_ONSP_Cyc_6, Anchor_Motor_ONSP_Cyc_7, " & _
                     "Anchor_Motor_ONSP_Cyc_8, Anchor_Motor_ONSP_Cyc_9"

    Dim colsHomogRPM As String
    colsHomogRPM = "Motor_Homoginiser_Mtr_RPM_SP_CYC_0, Motor_Homoginiser_Mtr_RPM_SP_CYC_1, " & _
                   "Motor_Homoginiser_Mtr_RPM_SP_CYC_2, Motor_Homoginiser_Mtr_RPM_SP_CYC_3, " & _
                   "Motor_Homoginiser_Mtr_RPM_SP_CYC_4, Motor_Homoginiser_Mtr_RPM_SP_CYC_5, " & _
                   "Motor_Homoginiser_Mtr_RPM_SP_CYC_6, Motor_Homoginiser_Mtr_RPM_SP_CYC_7, " & _
                   "Motor_Homoginiser_Mtr_RPM_SP_CYC_8, Motor_Homoginiser_Mtr_RPM_SP_CYC_9"

    Dim colsHomogONSP As String
    colsHomogONSP = "Homogeniser_Motor_ONSP_Cyc_0, Homogeniser_Motor_ONSP_Cyc_1, " & _
                    "Homogeniser_Motor_ONSP_Cyc_2, Homogeniser_Motor_ONSP_Cyc_3, " & _
                    "Homogeniser_Motor_ONSP_Cyc_4, Homogeniser_Motor_ONSP_Cyc_5, " & _
                    "Homogeniser_Motor_ONSP_Cyc_6, Homogeniser_Motor_ONSP_Cyc_7, " & _
                    "Homogeniser_Motor_ONSP_Cyc_8, Homogeniser_Motor_ONSP_Cyc_9"

    Dim colsWaterStirRPM As String
    colsWaterStirRPM = "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_0, Motor_Water_Stirrer_Mtr_RPM_SP_CYC_1, " & _
                       "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_2, Motor_Water_Stirrer_Mtr_RPM_SP_CYC_3, " & _
                       "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_4, Motor_Water_Stirrer_Mtr_RPM_SP_CYC_5, " & _
                       "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_6, Motor_Water_Stirrer_Mtr_RPM_SP_CYC_7, " & _
                       "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_8, Motor_Water_Stirrer_Mtr_RPM_SP_CYC_9"

    Dim colsWaterStirONSP As String
    colsWaterStirONSP = "Water_Heating_Water_Stirr_ONSP_CyC_0, Water_Heating_Water_Stirr_ONSP_CyC_1, " & _
                        "Water_Heating_Water_Stirr_ONSP_CyC_2, Water_Heating_Water_Stirr_ONSP_CyC_3, " & _
                        "Water_Heating_Water_Stirr_ONSP_CyC_4, Water_Heating_Water_Stirr_ONSP_CyC_5, " & _
                        "Water_Heating_Water_Stirr_ONSP_CyC_6, Water_Heating_Water_Stirr_ONSP_CyC_7, " & _
                        "Water_Heating_Water_Stirr_ONSP_CyC_8, Water_Heating_Water_Stirr_ONSP_CyC_9"

    Dim colsWaxStirRPM As String
    colsWaxStirRPM = "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_0, Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_1, " & _
                     "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_2, Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_3, " & _
                     "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_4, Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_5, " & _
                     "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_6, Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_7, " & _
                     "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_8, Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_9"

    Dim colsWaxStirONSP As String
    colsWaxStirONSP = "Wax_Stirrer_ONSP_CYC_0, Wax_Stirrer_ONSP_CYC_1, " & _
                      "Wax_Stirrer_ONSP_CYC_2, Wax_Stirrer_ONSP_CYC_3, " & _
                      "Wax_Stirrer_ONSP_CYC_4, Wax_Stirrer_ONSP_CYC_5, " & _
                      "Wax_Stirrer_ONSP_CYC_6, Wax_Stirrer_ONSP_CYC_7, " & _
                      "Wax_Stirrer_ONSP_CYC_8, Wax_Stirrer_ONSP_CYC_9"

    ' --- Final COLUMN concatenation ---
    Dim sqlCols As String
    sqlCols = "INSERT INTO RecipeMaster2 (" & _
                  colsMeta & ", " & _
                  colsWaxHeatTSP & ", " & _
                  colsWaxHeatONSP & ", " & _
                  colsProdHeatTSP & ", " & _
                  colsProdHeatONSP & ", " & _
                  colsProdCoolTSP & ", " & _
                  colsWaterHeatTSP & ", " & _
                  colsWaterHeatONSP & ", " & _
                  colsAnchorRPM & ", " & _
                  colsAnchorONSP & ", " & _
                  colsHomogRPM & ", " & _
                  colsHomogONSP & ", " & _
                  colsWaterStirRPM & ", " & _
                  colsWaterStirONSP & ", " & _
                  colsWaxStirRPM & ", " & _
                  colsWaxStirONSP & ") "

    ' ==============================================================
    ' VALUE GROUPS  — GetTagVal reads live from HMI TagGroup

    Dim valsMeta As String
    valsMeta = "'" & strRecipeName & "', " & _
               "'" & strBatchNumber & "', " & _
               "'" & strCreatedBy & "', " & _
                "GETDATE()" '

    Dim valsWaxHeatTSP As String
    valsWaxHeatTSP = GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_0") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_1") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_2") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_3") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_4") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_5") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_6") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_7") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_8") & ", " & _
                     GetTagVal("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_9")

    Dim valsWaxHeatONSP As String
    valsWaxHeatONSP = GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_0") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_1") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_2") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_3") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_4") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_5") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_6") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_7") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_8") & ", " & _
                      GetTagVal("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_9")

    Dim valsProdHeatTSP As String
    valsProdHeatTSP = GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_0") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_1") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_2") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_3") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_4") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_5") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_6") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_7") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_8") & ", " & _
                      GetTagVal("recipe\Product_Heating_MFG_Heat_TSP_Cyc_9")

    Dim valsProdHeatONSP As String
    valsProdHeatONSP = GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_0") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_1") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_2") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_3") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_4") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_5") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_6") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_7") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_8") & ", " & _
                       GetTagVal("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_9")

    Dim valsProdCoolTSP As String
    valsProdCoolTSP = GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_0") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_1") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_2") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_3") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_4") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_5") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_6") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_7") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_8") & ", " & _
                      GetTagVal("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_9")

    Dim valsWaterHeatTSP As String
    valsWaterHeatTSP = GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_0") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_1") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_2") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_3") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_4") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_5") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_6") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_7") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_8") & ", " & _
                       GetTagVal("recipe\Water_Heating_Water_Heat_TSP_Cyc_9")

    Dim valsWaterHeatONSP As String
    valsWaterHeatONSP = GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_0") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_1") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_2") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_3") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_4") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_5") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_6") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_7") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_8") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Heat_ONSP_Cyc_9")

    Dim valsAnchorRPM As String
    valsAnchorRPM = GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_0") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_1") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_2") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_3") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_4") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_5") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_6") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_7") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_8") & ", " & _
                    GetTagVal("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_9")

    Dim valsAnchorONSP As String
    valsAnchorONSP = GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_0") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_1") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_2") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_3") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_4") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_5") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_6") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_7") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_8") & ", " & _
                     GetTagVal("recipe\Anchor_Motor_ONSP_Cyc_9")

    Dim valsHomogRPM As String
    valsHomogRPM = GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_0") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_1") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_2") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_3") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_4") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_5") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_6") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_7") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_8") & ", " & _
                   GetTagVal("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_9")

    Dim valsHomogONSP As String
    valsHomogONSP = GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_0") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_1") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_2") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_3") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_4") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_5") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_6") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_7") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_8") & ", " & _
                    GetTagVal("recipe\Homogeniser_Motor_ONSP_Cyc_9")

    Dim valsWaterStirRPM As String
    valsWaterStirRPM = GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_0") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_1") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_2") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_3") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_4") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_5") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_6") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_7") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_8") & ", " & _
                       GetTagVal("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_9")

    Dim valsWaterStirONSP As String
    valsWaterStirONSP = GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_0") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_1") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_2") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_3") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_4") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_5") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_6") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_7") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_8") & ", " & _
                        GetTagVal("recipe\Water_Heating_Water_Stirr_ONSP_CyC_9")

    Dim valsWaxStirRPM As String
    valsWaxStirRPM = GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_0") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_1") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_2") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_3") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_4") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_5") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_6") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_7") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_8") & ", " & _
                     GetTagVal("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_9")

    Dim valsWaxStirONSP As String
    valsWaxStirONSP = GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_0") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_1") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_2") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_3") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_4") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_5") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_6") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_7") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_8") & ", " & _
                      GetTagVal("recipe\Wax_Stirrer_ONSP_CYC_9")

    ' --- Final VALUES concatenation ---
    Dim sqlVals As String
    sqlVals = "VALUES (" & _
                  valsMeta & ", " & _
                  valsWaxHeatTSP & ", " & _
                  valsWaxHeatONSP & ", " & _
                  valsProdHeatTSP & ", " & _
                  valsProdHeatONSP & ", " & _
                  valsProdCoolTSP & ", " & _
                  valsWaterHeatTSP & ", " & _
                  valsWaterHeatONSP & ", " & _
                  valsAnchorRPM & ", " & _
                  valsAnchorONSP & ", " & _
                  valsHomogRPM & ", " & _
                  valsHomogONSP & ", " & _
                  valsWaterStirRPM & ", " & _
                  valsWaterStirONSP & ", " & _
                  valsWaxStirRPM & ", " & _
                  valsWaxStirONSP & ")"

    ' ==============================================================
    ' EXECUTE

    sqlQuery = sqlCols & sqlVals


          'MsgBox (sqlQuery)
    LogDiagnosticsMessage (sqlQuery)
    

    ' --- 5. EXECUTE AND CLEANUP ---
    conn.Execute sqlQuery
    
    MsgBox "Recipe saved successfully", vbInformation, "Success"
    LogDiagnosticsMessage "Recipe saved successfully"

    conn.Close
    Set conn = Nothing
    Exit Sub

ErrHandler:
    LogDiagnosticsMessage "Save Recipe Error : " & Err.Description
    MsgBox "Save Recipe Error : " & Err.Description, vbCritical, "Database Error"
    
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub


Private Sub updateBtn_Released()
       updateToSql
End Sub

Private Sub updateToSql()
       
    Dim conn       As ADODB.Connection
    Dim sqlQuery   As String
    Dim i          As Integer
    
    Dim recipeTagGroup As TagGroup
    Set recipeTagGroup = InitRecipeTagGroup
    ' ==============================================================
    ' 1. VALIDATION
    ' ==============================================================
    If Trim(Me.recipeNameInput.Value) = "" Then
        MsgBox "Error: Please Enter Recipe Name", vbCritical, "Missing Data"
        Exit Sub
    End If
    
    
       Dim strRecipeName   As String
    strRecipeName = GetTagVal("recipe\recipeName")
     MsgBox "you are updating " & strRecipeName & "recipe"
    ' ==============================================================
    ' 3. DATABASE CONNECTION
    ' ==============================================================
    Set conn = New ADODB.Connection
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    ' ==============================================================
    ' 4. BUILD UPDATE QUERY — grouped SET strings
    ' ==============================================================

    ' --- Group: Wax Heat SP TSP ---
    Dim setWaxHeatTSP As String
    setWaxHeatTSP = ""
    For i = 0 To 9
        setWaxHeatTSP = setWaxHeatTSP & _
            "PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i).Value
        If i < 9 Then setWaxHeatTSP = setWaxHeatTSP & ", "
    Next i

    ' --- Group: Wax Heat ONSP ---
    Dim setWaxHeatONSP As String
    setWaxHeatONSP = ""
    For i = 0 To 9
        setWaxHeatONSP = setWaxHeatONSP & _
            "PID_Parameter_Wax_Heat_ONSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_" & i).Value
        If i < 9 Then setWaxHeatONSP = setWaxHeatONSP & ", "
    Next i

    ' --- Group: Product Heating TSP ---
    Dim setProdHeatTSP As String
    setProdHeatTSP = ""
    For i = 0 To 9
        setProdHeatTSP = setProdHeatTSP & _
            "Product_Heating_MFG_Heat_TSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_TSP_Cyc_" & i).Value
        If i < 9 Then setProdHeatTSP = setProdHeatTSP & ", "
    Next i

    ' --- Group: Product Heating ONSP ---
    Dim setProdHeatONSP As String
    setProdHeatONSP = ""
    For i = 0 To 9
        setProdHeatONSP = setProdHeatONSP & _
            "Product_Heating_MFG_Heat_ONSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Product_Heating_MFG_Heat_ONSP_Cyc_" & i).Value
        If i < 9 Then setProdHeatONSP = setProdHeatONSP & ", "
    Next i

    ' --- Group: Product Cooling TSP ---
    Dim setProdCoolTSP As String
    setProdCoolTSP = ""
    For i = 0 To 9
        setProdCoolTSP = setProdCoolTSP & _
            "Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i & " = " & _
            recipeTagGroup.Item("recipe\Product_Cooling_MFG_Cool_TSP_CycMW1168_" & i).Value
        If i < 9 Then setProdCoolTSP = setProdCoolTSP & ", "
    Next i

    ' --- Group: Water Heating TSP ---
    Dim setWaterHeatTSP As String
    setWaterHeatTSP = ""
    For i = 0 To 9
        setWaterHeatTSP = setWaterHeatTSP & _
            "Water_Heating_Water_Heat_TSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_TSP_Cyc_" & i).Value
        If i < 9 Then setWaterHeatTSP = setWaterHeatTSP & ", "
    Next i

    ' --- Group: Water Heating ONSP ---
    Dim setWaterHeatONSP As String
    setWaterHeatONSP = ""
    For i = 0 To 9
        setWaterHeatONSP = setWaterHeatONSP & _
            "Water_Heating_Water_Heat_ONSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Water_Heating_Water_Heat_ONSP_Cyc_" & i).Value
        If i < 9 Then setWaterHeatONSP = setWaterHeatONSP & ", "
    Next i

    ' --- Group: Anchor Motor RPM SP ---
    Dim setAnchorRPM As String
    setAnchorRPM = ""
    For i = 0 To 9
        setAnchorRPM = setAnchorRPM & _
            "Motor_Anchor_Mtr_RPM_SP_CYC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Motor_Anchor_Mtr_RPM_SP_CYC_" & i).Value
        If i < 9 Then setAnchorRPM = setAnchorRPM & ", "
    Next i

    ' --- Group: Anchor Motor ONSP ---
    Dim setAnchorONSP As String
    setAnchorONSP = ""
    For i = 0 To 9
        setAnchorONSP = setAnchorONSP & _
            "Anchor_Motor_ONSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Anchor_Motor_ONSP_Cyc_" & i).Value
        If i < 9 Then setAnchorONSP = setAnchorONSP & ", "
    Next i

    ' --- Group: Homogeniser Motor RPM SP ---
    Dim setHomogRPM As String
    setHomogRPM = ""
    For i = 0 To 9
        setHomogRPM = setHomogRPM & _
            "Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Motor_Homoginiser_Mtr_RPM_SP_CYC_" & i).Value
        If i < 9 Then setHomogRPM = setHomogRPM & ", "
    Next i

    ' --- Group: Homogeniser Motor ONSP ---
    Dim setHomogONSP As String
    setHomogONSP = ""
    For i = 0 To 9
        setHomogONSP = setHomogONSP & _
            "Homogeniser_Motor_ONSP_Cyc_" & i & " = " & _
            recipeTagGroup.Item("recipe\Homogeniser_Motor_ONSP_Cyc_" & i).Value
        If i < 9 Then setHomogONSP = setHomogONSP & ", "
    Next i

    ' --- Group: Water Stirrer RPM SP ---
    Dim setWaterStirRPM As String
    setWaterStirRPM = ""
    For i = 0 To 9
        setWaterStirRPM = setWaterStirRPM & _
            "Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Motor_Water_Stirrer_Mtr_RPM_SP_CYC_" & i).Value
        If i < 9 Then setWaterStirRPM = setWaterStirRPM & ", "
    Next i

    ' --- Group: Water Stirrer ONSP ---
    Dim setWaterStirONSP As String
    setWaterStirONSP = ""
    For i = 0 To 9
        setWaterStirONSP = setWaterStirONSP & _
            "Water_Heating_Water_Stirr_ONSP_CyC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Water_Heating_Water_Stirr_ONSP_CyC_" & i).Value
        If i < 9 Then setWaterStirONSP = setWaterStirONSP & ", "
    Next i

    ' --- Group: Wax Stirrer RPM SP ---
    Dim setWaxStirRPM As String
    setWaxStirRPM = ""
    For i = 0 To 9
        setWaxStirRPM = setWaxStirRPM & _
            "Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Motor_Wax_Stirrer_Mtr_RPM_SP_CYC_" & i).Value
        If i < 9 Then setWaxStirRPM = setWaxStirRPM & ", "
    Next i

    ' --- Group: Wax Stirrer ONSP ---
    Dim setWaxStirONSP As String
    setWaxStirONSP = ""
    For i = 0 To 9
        setWaxStirONSP = setWaxStirONSP & _
            "Wax_Stirrer_ONSP_CYC_" & i & " = " & _
            recipeTagGroup.Item("recipe\Wax_Stirrer_ONSP_CYC_" & i).Value
        If i < 9 Then setWaxStirONSP = setWaxStirONSP & ", "
    Next i

    ' ==============================================================
    ' 5. FINAL QUERY CONCATENATION
    ' ==============================================================
    sqlQuery = "UPDATE RecipeMaster2 SET " & _
                    setWaxHeatTSP & ", " & _
                    setWaxHeatONSP & ", " & _
                    setProdHeatTSP & ", " & _
                    setProdHeatONSP & ", " & _
                    setProdCoolTSP & ", " & _
                    setWaterHeatTSP & ", " & _
                    setWaterHeatONSP & ", " & _
                    setAnchorRPM & ", " & _
                    setAnchorONSP & ", " & _
                    setHomogRPM & ", " & _
                    setHomogONSP & ", " & _
                    setWaterStirRPM & ", " & _
                    setWaterStirONSP & ", " & _
                    setWaxStirRPM & ", " & _
                    setWaxStirONSP & _
               " WHERE recipeName = '" & recipeTagGroup.Item("recipe\recipeName").Value & "'"

    ' ==============================================================
    ' 6. EXECUTE AND CLEANUP
    ' ==============================================================
    LogDiagnosticsMessage (sqlQuery)
    conn.Execute sqlQuery

    MsgBox "Recipe updated successfully!", vbInformation, "Success"
    LogDiagnosticsMessage "Recipe updated successfully"

    conn.Close
    Set conn = Nothing
    Exit Sub

ErrHandler:
    LogDiagnosticsMessage "Update Recipe Error: " & Err.Description
    MsgBox "Update Recipe Error: " & Err.Description, vbCritical, "Database Error"
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If

End Sub

