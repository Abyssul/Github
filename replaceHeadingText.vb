Sub replaceHeadingText(ctrl As IRibbonControl)
    'Function parameter: 'ctrl As IRibbonControl' (Must be removed if debugging)

    Dim userName As String
    Dim fieldEmptyError As Boolean
    
    fieldEmptyError = False
    
    userName = Application.userName
    
    Dim replaceArray(0 To 5) As String
    replaceArray(0) = "Full_Name"
    replaceArray(1) = "Position_Variable"
    replaceArray(2) = "Location_Variable"
    replaceArray(3) = "Direct_Phone"
    replaceArray(4) = "Direct_Fax"
    replaceArray(5) = "Email_Variable"
    
    Dim firstName As String

    ' Connect to worksheet as database
    Set conExcel = New ADODB.Connection
    
    'Connect to database
    conExcel.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\AMICUSSVR\DocumentAssemblyTemplates\Template Header\contactHeaderDatabase.xlsx;Extended Properties='Excel 12.0 Xml;HDR=YES';"
    conExcel.Open
    Set rsExcel = New ADODB.Recordset
                 
    ' Compare to Excel DB
    strQuery = "SELECT [Full Name] FROM [Sheet1$] ORDER BY [Full Name]"
    
    ' Find records
    rsExcel.Open strQuery, conExcel
    
    'Set pointer to first record in recordset (currently empty)
    rsExcel.MoveFirst
    
    'Load the database into a recordset for parsing
    With FromTheDeskOf.FullName
        .Clear
        Do
        
           .AddItem rsExcel![Full Name]
            rsExcel.MoveNext
            
        Loop Until rsExcel.EOF
    End With
            
    FromTheDeskOf.FullName.Value = userName
    
    'Show form
    FromTheDeskOf.Show
    
    'Select row that contains the selection from the form
    strQuery = "SELECT * FROM [Sheet1$] WHERE [Full Name] = '" & FromTheDeskOf.FullName.Value & " '"
    
    rsExcel.Close 'Close record set so we can requery
    rsExcel.Open strQuery, conExcel 'Requery with selection
    
    'Loop through all values
    For i = 0 To 5
        If IsNull(rsExcel.Fields.Item(i)) Then
                    
              fieldEmptyError = True
              
        Else
        
        ' Find and Replace document with record
            Selection.Find.ClearFormatting
            Selection.Find.Replacement.ClearFormatting
            With Selection.Find
                .Text = replaceArray(i)
                .Replacement.Text = rsExcel.Fields.Item(i)
                .Forward = True
                .Wrap = wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            Selection.Find.Execute Replace:=wdReplaceAll
                  
        End If
        
    Next
        
    'Display message box if the selection contained a blank/null value
    If fieldEmptyError = True Then
        MsgBox ("One or more fields are empty for the person selected. Please manually fill them out.")
    End If
    
    conExcel.Close
    
End Sub

