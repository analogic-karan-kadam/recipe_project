

Private Sub DeleteRecipeBtn_Released()
    On Error GoTo ErrHandler

    Dim conn As ADODB.Connection
    Dim sql As String
    Dim recipeName As String

    If recipeMasterTable.Row <= 0 Then

        MsgBox "Please select a recipe"

        Exit Sub

    End If
    recipeName = recipeMasterTable.TextMatrix(recipeMasterTable.Row, 0)


    Set conn = New ADODB.Connection

    conn.Open _
        "Provider=MSOLEDBSQL;" & _
        "Data Source=AAPL-L170\SQLEXPRESS;" & _
        "Initial Catalog=vbaLiveProjects;" & _
        "User ID=sa;" & _
        "Password=sa;" & _
        "TrustServerCertificate=yes;"

    sql = "DELETE FROM recipeMasterScreen WHERE recipeName = '" & recipeName & "'"

    conn.Execute sql
    
    MsgBox "Recipe deleted successfully"
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

Private Sub Display_AnimationStart()
'**---------------load recipe table data on start
    LoadRecipeData
End Sub
'**---------------load recipe data funtion
Public Sub LoadRecipeData()
    On Error GoTo ErrHandler
    Dim conn As ADODB.Connection
    Dim rs As ADODB.Recordset
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
    sql = "SELECT * FROM recipeMasterScreen"
    Set rs = conn.Execute(sql)
    recipeMasterTable.Clear
    
    recipeMasterTable.Rows = 1
    recipeMasterTable.Cols = rs.Fields.Count
    recipeMasterTable.ColWidth(0) = 2500   ' Recipe Name
    recipeMasterTable.ColWidth(1) = 1500   ' status
    recipeMasterTable.ColWidth(2) = 1500   ' disable
    recipeMasterTable.ColWidth(3) = 2500   ' createdBy
    recipeMasterTable.ColWidth(4) = 2500   ' Product name
    ' Column Names
    For c = 0 To rs.Fields.Count - 1
        recipeMasterTable.TextMatrix(0, c) = _
            rs.Fields(c).Name
    Next c
    If rs.EOF Then
        LogDiagnosticsMessage _
            "No records found in recipeMasterScreen"
        rs.Close
        conn.Close
    Exit Sub

    End If
    
    ' Data Rows
    Do Until rs.EOF
        recipeMasterTable.Rows = recipeMasterTable.Rows + 1

        r = recipeMasterTable.Rows - 1
        For c = 0 To rs.Fields.Count - 1
            '**- recipeMasterTable.TextMatrix(r, c) = _
                rs.Fields(c).Value & ""
             recipeMasterTable.TextMatrix(r, c) = GetDisplayValue(rs.Fields(c).Name, rs.Fields(c).Value & "")
        Next c
        rs.MoveNext
    Loop
    rs.Close
    conn.Close
    Set rs = Nothing
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
'**------------------------------------get display value for 0 and 1
Private Function GetDisplayValue( _
    ByVal columnName As String, _
    ByVal rawValue As Variant) As String

    Select Case LCase(columnName)

        Case "status"

            If CBool(rawValue) Then
                GetDisplayValue = "Approved"
            Else
                GetDisplayValue = "Not Approved"
            End If

        Case "disable"

            If CBool(rawValue) Then
                GetDisplayValue = "Yes"
            Else
                GetDisplayValue = "No"
            End If

        Case Else

            GetDisplayValue = CStr(rawValue)

    End Select

End Function
Private Sub updateBtn_Released()
    UpdateRecipeVersion
End Sub
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

    sql = "UPDATE recipeMasterScreen " & _
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

Private Sub ViewRecipeBtn1_Released()
    If recipeMasterTable.Row <= 0 Then
        MsgBox "Please select a recipe"
        Exit Sub
    End If
    ExecuteCommand "Display RecipeView"

End Sub
