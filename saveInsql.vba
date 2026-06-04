Private Sub saveToSqlBtn_Accepted(PerformerUser As String, PerformerComment As String, ApproveUser As String, ApproveComment As String)
     'MsgBox "performer : " & PerformerUser

    'On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim recipeData As TagGroup
    Dim sql As String

    'Set recipeData = Application.CreateTagGroup(Me.AreaName)
    If recipeData Is Nothing Then
        Set recipeData = Application.CreateTagGroup(Me.AreaName)
        recipeData.Add "recipe\recipeName"
        recipeData.Add "recipe\productName"
        recipeData.Add "recipe\createdBy"
        recipeData.Add "recipe\status"
        recipeData.Add "recipe\disable"
        recipeData.Active = True
    End If

    recipeData.Active = True

    Set conn = New ADODB.Connection
     
    If Len(Trim(Me.recipeNameInput.Value)) = 0 Or _
       Len(Trim(Me.productNameInput.Value)) = 0 Then
        MsgBox "Error: Please Enter Data ", vbCritical, "Missing Data"
        Exit Sub
        
    End If
    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "INSERT INTO recipeMasterScreen " & _
      "(recipeName, createdBy, productName) VALUES (" & _
      "'" & recipeNameInput.Value & "'," & _
      "" & _
      "'" & PerformerUser & "'," & _
      "'" & productNameInput.Value & "')"
      'sql = "INSERT INTO recipeMasterScreen " & _
      "(recipeName, productName) VALUES (" & _
      "'" & recipeNameInput.Value & "'," & _
      "" & _
      "'" & productNameInput.Value & "')"
     
    LogDiagnosticsMessage sql

    conn.Execute sql

    MsgBox "Recipe saved successfully"

    LogDiagnosticsMessage "Recipe saved successfully"

    conn.Close

    Set conn = Nothing

    Exit Sub

ErrHandler:

    LogDiagnosticsMessage _
        "Save Recipe Error : " & Err.Description

    MsgBox _
        "Save Recipe Error : " & Err.Description

End Sub