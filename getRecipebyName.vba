Function renderRecipeParameters(ByVal recipeName As String)
    
    MsgBox ("Render function for " & recipeName)

    Dim conn
    Dim Rs
    Dim Field, xyz
    Dim ServerName, sqlUsername, sqlPassword, Table, Database As String
    Dim recipeTagGroup As TagGroup
   
   If recipeTagGroup Is Nothing Then
    Set recipeTagGroup = Application.CreateTagGroup(Me.AreaName)
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
         
   End If
   
    Set conn = New ADODB.Connection
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"
   
 
     
    StrQuery = "SELECT  *  FROM [vbaLiveProjects].[dbo].[RecipeMaster] where recipeName = '" & recipeName & "'"
     Set Rs = conn.Execute(StrQuery)
   
   
    
If Rs.EOF Then
    'do nothing
Else
   'set tags of the group and username
   Do While Not Rs.EOF
   
    For i = 0 To 9
    recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i).Value = _
        Rs("TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i)
        
        'MsgBox ("Resultset data for tag 1 value " & Rs("TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_" & i))
    Next i
    
    'recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_0").Value = Rs("TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_0")
    'recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_1").Value = Rs("TESTPID_Parameter_Wax_Heat_SP_TSP_Cyc_1")
    
    For i = 0 To 9
    recipeTagGroup.Item("recipe\PID_Parameter_Wax_Heat_ONSP_Cyc_" & i).Value = _
        Rs("TESTPID_Parameter_Wax_Heat_ONSP_Cyc_" & i)
    Next i

        Rs.MoveNext
    Loop
End If
    conn.Close
   
End Function