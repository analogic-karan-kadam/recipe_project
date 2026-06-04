Private Sub GoToMainScreenBtn_Released()
    Dim CurrentUser As String
    CurrentUser = Application.CurrentUserName
    MsgBox CurrentUser

End Sub

Private Sub SaveRecipeBtn_Released()

    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim recipe As TagGroup
    Dim sql As String

    Set recipeData = Application.CreateTagGroup(Me.AreaName)

    recipeData.Add "recipe\recipeName"
    recipeData.Add "recipe\productName"
    recipeData.Add "recipe\createdBy"
    recipeData.Add "recipe\status"
    recipeData.Add "recipe\disable"

    recipeData.Active = True

    Set conn = New ADODB.Connection
     
    If Len(Trim(Me.recipeNameInput.Value)) = 0 Or _
       Len(Trim(Me.createdByInput.Value)) = 0 Or _
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
      "'" & createdByInput.Value & "'," & _
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
