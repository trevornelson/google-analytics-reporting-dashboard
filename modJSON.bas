Attribute VB_Name = "modJSON"
Option Compare Database

Public Sub ParseJSON(json As String, m As String, Med As Date, Optional ByVal Api As String)

Dim sD As Long 'start of data point to be extracted
Dim eD As Long 'end of data point to be extracted
Dim ElemTag As String 'the string the InStr function searches for within json
Dim CharDiff As Long
Dim dP1 As Variant

'Below code finds data type.
'Dim dTypeTag As String 'The string the InStr function searches within json
'Dim dTypeCharDiff As Long
'Dim dType As String
'Dim dT1 As Long 'start of datatype json element
'Dim dT2 As Long 'end of datatype json element

'dTypeTag = "dataType"
'dT1 = InStr(1, json, dTypeTag)
'dT1 = dT1 + 11
'dT2 = InStr(dT1, json, "]")
'dT2 = dT2 - 2
'dTypeCharDiff = dT2 - dT1

'dType = Mid(json, dT1, dTypeCharDiff)
'If dType = "INTEGER" Then
'    Dim dP1 As Long
'    Debug.Print dType
'ElseIf dType = "DOUBLE" Then
'    Dim dP1 As Double
'    Debug.Print dType
'Else
'    Dim dP1 As String
'    Debug.Print dType
'End If



ElemTag = "rows"

sD = InStrRev(json, ElemTag)
'(1, json, ElemTag)
sD = sD + 9
eD = InStr(sD, json, "]]")
eD = eD - 1
CharDiff = eD - sD
    
    Debug.Print sD
    Debug.Print eD
    Debug.Print CharDiff
    
If CharDiff > 0 Then
    dP1 = Mid(json, sD, CharDiff)
Else
    Debug.Print "Looks like this datapoint is zero. How datapointing!"
    dP1 = 0
End If
    
    
    Dim MetName As String
    Dim endD As Date
    
    MetName = m
    endD = Med
        
    Call GAimportToTable(MetName, dP1, endD, Api)
    


'The below code block is for later- it will be used to find multiple "rows" in json string
'While InStr(eD, json, ElemTag) Is Not Nothing
'    sD = InStr(eD, json, ElemTag)
'    sD = sD + 9
'    eD = InStr(sD, json, "]]")
'    eD = eD - 1
'    CharDiff = eD - sD


End Sub


Public Sub GAimportToTable(ByVal m As String, ByVal data As Variant, ByVal eD As Date, ByVal Api As String)
    'input "m" is the Query Name
    'input "data" is the parsed datapoint
    'input "ed" is the end date
    'input "API" is the API query table being used
    

    Dim ImpTable As String
    Dim ImpField As String

    ImpTable = DLookup("[TableName]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    ImpField = DLookup("[FieldName]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    Debug.Print "Import to Table" & ImpTable
    Debug.Print "Import to Field" & ImpField
    
    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    'rs = dB.OpenRecordset(ImpTable)
    
    Debug.Print "Data to import:" & data
    
    Dim ImpSQL As String
    ImpSQL = "UPDATE " & ImpTable & " SET " & ImpField & " = '" & data & "' WHERE MonthEndDate = #" & eD & "#"
    
    Debug.Print ImpSQL
    
    dB.Execute ImpSQL
    
    
End Sub

Public Sub JSONtoXML(json As String)

    Dim ie As Object
    Dim frm As Variant
    Dim element As Variant
    Dim btn As Variant
    Dim Response As Variant
    Dim xml As String
        
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate "http://www.utilities-online.info/xmltojson/"
    While ie.ReadyState <> 4: DoEvents: Wend
    
    Set frm = ie.Document.getelementbyid("json")
    Set btn = ie.Document.getelementbyid("toxml")
    
    ie.Visible = True
    frm.Value = json
    btn.Click
    
    Set Response = ie.Document.getelementbyid("xml")
    
    xml = Response.Value
    
    Debug.Print xml
    
    ie.Quit
    Set ie = Nothing

    
End Sub


Public Sub NumberOfJSONRows(ByVal QueryResult As String, ByVal ImportTable As String, Field1 As String, Field2 As String, Field3 As String, Optional ByVal Dimension As String)

'This sub returns the number of rows in a 2 column JSON response.
'
'Inputs:
'   QueryResult is the variable that passes the JSON response. It should be unparsed.
'   ImportTable is the table the number of responses should be added to
'   Field1 and Field 2 are the field names in the ImportTable where data should be added.
'   Dimension is the identifier string of a number of results. This is optional. i.e. you can pass "LastMonth" into this and it will add "LastMonth to the import table.
'   **special note: This doesn't clear the importtable in this sub. It needs to be done before this sub is called.


    'Starts VBs javascript engine
    Dim objJson As MSScriptControl.ScriptControl
    Set objJson = New MSScriptControl.ScriptControl
    objJson.Language = "JScript"
    
    Debug.Print "Printed Response From JSON Parsing"
    Debug.Print QueryResult
    
    'Creates a javascript object and assigns it the JSON query result input
    Dim objResp As Object
    Set objResp = objJson.Eval("(" & QueryResult & ")")
    
    Dim dB As Database
    Set dB = CurrentDb
    Dim rs As Recordset
    Set rs = dB.OpenRecordset(ImportTable)
    
    Dim NumberOfCommas As Integer
    Dim NumberOfRows As Integer
    
    If InStr(QueryResult, "rows") = 0 Then
        NumberOfRows = 0
    Else
        Dim Response As String
        Response = objResp.rows
        Debug.Print Response
        
        NumberOfCommas = Len(Response) - Len(Replace(Response, ",", ""))
        
        Debug.Print NumberOfCommas
        
        NumberOfRows = (NumberOfCommas - 1) / 2
        
        Debug.Print NumberOfRows
        
        Dim ReportingMonthLPCount
        
        If Dimension = "Month1" Then
            Debug.Print ""
        Else
            ReportingMonthLPCount = DLookup("[LandingPageResults]", "[tblGALandingPageTotals]", "MonthIdentifier = 'Month1'")
        End If
    End If
    
    rs.AddNew
        rs(Field1).Value = Dimension
        rs(Field2).Value = NumberOfRows
        If Dimension = "Month1" Then
            rs(Field3).Value = NumberOfRows
        Else
            rs(Field3).Value = ReportingMonthLPCount
    End If
    rs.Update
    
    rs.Close
    dB.Close
    Set rs = Nothing
    Set dB = Nothing
    
End Sub
