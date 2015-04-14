Attribute VB_Name = "modGoogleAPIQueries"
Option Compare Database
Public Sub testGAQueries()
'This sub is for manual metrics querying.
'it will loop through the dates for a metric and update them all.


Dim c As String
Dim m As String
Dim aT As String

c = "70742771"
m = "notprovNonBranded"
aT = "ya29.AHES6ZRNdcMRdHh6xhcPO3mdrCbI8fuMNaV7x6gIUFG4a3k"

Call GAqueryBuild(c, m, aT)
'Call GoogAPIClientSpecific(c, m, "tblLiquidSpaceApiQueries", aT)

End Sub

Public Sub manualloopGAQueries()

Dim c As String
Dim m As String
Dim aT As String
Dim Api As String
Dim dB As Database
Dim rsQ2 As Recordset



c = "42625605" 'standard profile view
'c = "70695004" 'novpv profile view
aT = "ya29.AHES6ZRNdcMRdHh6xhcPO3mdrCbI8fuMNaV7x6gIUFG4a3k"
Api = "tblLiquidSpaceApiQueries"

Set dB = CurrentDb
Set rsQ2 = dB.OpenRecordset(Api)

If Not (rsQ2.EOF And rsQ2.BOF) Then
        Do Until rsQ2.EOF = True
            
            m = rsQ2("QueryName").Value
            Call GoogAPIClientSpecific(c, m, Api, aT)
            
            rsQ2.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If

End Sub


Public Sub GoogleAnalyticsAPIHandler(ByVal c As String, ByVal Api As String, ByVal aT As String, ByVal bT As String)


Dim m As String

Dim dB As Database
Dim rsQ As Recordset
Dim rsQ2 As Recordset

Set dB = CurrentDb
If Api = "" Then
    Debug.Print "No Client Specific Queries input in the Clients table."
Else
    Set rsQ2 = dB.OpenRecordset(Api)
End If

Set rsQ = dB.OpenRecordset("tblAPIQueries")

If Not (rsQ.EOF And rsQ.BOF) Then
    Do Until rsQ.EOF = True
        
        m = rsQ("QueryName").Value
        Call GAqueryBuild(c, m, aT, bT)
        
        rsQ.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If

Dim rsD As Recordset
Set rsD = dB.OpenRecordset("tblDimAPIQueries")

If Not (rsD.EOF And rsD.BOF) Then
    Do Until rsD.EOF = True
        m = rsD("QueryName").Value
        Call DimDataQueries("tblDimAPIQueries", m, c, aT)
        rsD.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If

        
If Api = "" Then  'if statement checks if API is null, ie if the client doesn't have a specific api table. If not, it skips googapiclientspecific
    Debug.Print "No Client specific querying instructions"
Else
    If Not (rsQ2.EOF And rsQ2.BOF) Then
        Do Until rsQ2.EOF = True
            
            m = rsQ2("QueryName").Value
            Call GoogAPIClientSpecific(c, m, Api, aT)
            
            rsQ2.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
End If
            
rsD.Close
rsQ.Close
If Api = "" Then
    Debug.Print "no client specific api instructions"
Else
    rsQ2.Close
End If
dB.Close
Set rsD = Nothing
Set rsQ = Nothing
Set rsQ2 = Nothing
Set dB = Nothing

End Sub


Public Sub GAqueryBuild(ByVal c As String, ByVal m As String, ByVal aT As String, ByVal bT As String)
    'Builds a Google Analytics Query URL and queries the Google Analytics server.
    'Oauth flow required previously, with a fresh Access Token issued and implemented into the query urls.
    'the c input is the profile ID (found in the settings of a specified GA profile
    'the m input is the query name.
    
    Dim URL As String
    Dim bodySend As String
    Dim MetricName As String
    Dim Metric As String
    Dim MetricParam As String
    Dim ProfID As String
    Dim MainAPI As String
    
    MainAPI = "tblApiQueries"
    
    MetricParam = DLookup("[MetricID]", "tblApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    Metric = DLookup("[Parameter2]", "tblApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    Dim TabN As String
    Dim FieldN As String

    TabN = DLookup("[TableName]", "tblApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    FieldN = DLookup("[FieldName]", "tblApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    
    Debug.Print m
    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset(TabN)
    
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF = True
        
            Dim Sdate As Date
            Dim Edate As Date
            Dim GAStartDate As String
            Dim GAEndDate As String
            
            Debug.Print TabN
            
            
            Sdate = rs("MonthStartDate").Value
            Edate = rs("MonthEndDate").Value
    
            GAStartDate = Format(Sdate, "yyyy-mm-dd") 'Converts Sdate to a string in the format of GA query
            GAEndDate = Format(Edate, "yyyy-mm-dd") 'Converts Edate to a string in the format of GA query
    
            Debug.Print GAEndDate
    
            ProfID = c
    
            Debug.Print MetricParam
            Debug.Print Metric
            Debug.Print GAStartDate
            Debug.Print GAEndDate
            Debug.Print ProfID
    
            
            ' The Following IF statements add nonbranded and branded filtering to relevant queries
            If m = "notprovBranded" Then
                MetricParam = MetricParam & "&filters=ga:medium%3D%3Dorganic;ga:keyword%3D@" & bT
            End If
    
            If m = "notprovNonBranded" Then
                MetricParam = MetricParam & "&filters=ga:medium%3D%3Dorganic;ga:keyword!@%28not%20provided%29;ga:keyword!@" & bT
            End If
            
            If m = "notprovRevBranded" Then
                MetricParam = MetricParam & "&filters=ga:medium%3D%3Dorganic;ga:keyword%3D@" & bT
            End If
            
            If m = "notprovRevNonBranded" Then
                MetricParam = MetricParam & "&filters=ga:medium%3D%3Dorganic;ga:keyword!@%28not%20provided%29;ga:keyword!@" & bT
            End If
            'End of nonBranded and branded filters
                       
            If c = "24294295" Then
                MetricParam = MetricParam & "&segment=gaid%3A%3AmyvdFr97RnWuty5z5nNL5Q"
            End If
                       
            URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & ProfID & "&metrics=" & MetricParam & "&start-date=" & GAStartDate & "&end-date=" & GAEndDate
            
            Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            Dim QueryResult As String
            QueryResult = objhttp.responseText
            
            Debug.Print QueryResult
                        
            Call ParseJSON(QueryResult, m, Edate, MainAPI)
            
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
            
                
    rs.Close
    dB.Close
    Set rs = Nothing
    Set dB = Nothing
            
           
End Sub

Public Sub GoogAPIClientSpecific(ByVal c As String, ByVal m As String, ByVal Api As String, ByVal aT As String)
    'Builds a Google Analytics Query URL and queries the Google Analytics server.
    'Oauth flow required previously, with a fresh Access Token issued and implemented into the query urls.
    'the c input is the profile ID (found in the settings of a specified GA profile
    'the m input is the query name.
    'the API input is the API query table name
    
    Dim URL As String
    Dim bodySend As String
    Dim MetricName As String
    Dim Metric As String
    Dim MetricParam As String
    Dim ProfID As String
    Dim ParamModifier As String
    
    MetricParam = DLookup("[MetricID]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    Metric = DLookup("[Parameter2]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    If Not IsNull(DLookup("[ParamModifier]", Api, "[QueryName] = " & "'" & m & "'")) Then
        ParamModifier = DLookup("[ParamModifier]", Api, "[QueryName] = " & "'" & m & "'")
    Else
        ParamModifier = ""
    End If
    
    Dim TabN As String
    Dim FieldN As String

    TabN = DLookup("[TableName]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    FieldN = DLookup("[FieldName]", Api, "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    Debug.Print m
    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset(TabN)
    
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF = True
        
            Dim Sdate As Date
            Dim Edate As Date
            Dim GAStartDate As String
            Dim GAEndDate As String
            
            Debug.Print TabN
            
            
            Sdate = rs("MonthStartDate").Value
            Edate = rs("MonthEndDate").Value
    
            GAStartDate = Format(Sdate, "yyyy-mm-dd") 'Converts Sdate to a string in the format of GA query
            GAEndDate = Format(Edate, "yyyy-mm-dd") 'Converts Edate to a string in the format of GA query
    
            Debug.Print GAEndDate
    
            ProfID = c
    
            Debug.Print MetricParam
            Debug.Print Metric
            Debug.Print GAStartDate
            Debug.Print GAEndDate
            Debug.Print ProfID
    
    
            
                       
            URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & ProfID & "&metrics=" & MetricParam & ParamModifier & "&start-date=" & GAStartDate & "&end-date=" & GAEndDate
            
            Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            Dim QueryResult As String
            QueryResult = objhttp.responseText
            
            Debug.Print QueryResult

                        
            Call ParseJSON(QueryResult, m, Edate, Api)
            
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
            
                
    rs.Close
    dB.Close
    Set rs = Nothing
    Set dB = Nothing



End Sub

Public Sub NonDateAPIHandler() 'this sub is borked at the moment
Dim c As String
c = "3199261"

Dim m As String
Dim dB As Database
Dim rsND As Recordset
Dim aT As String 'GA access token from oauth

aT = "ya29.AHES6ZSksvlSYLtBxIPp4Nyq4emmEeKk-2-Ee1aHtlCNBteVCYZ8tA"

Set dB = CurrentDb
Set rsND = dB.OpenRecordset("tblNonDateApiQueries")

If Not (rsND.EOF And rsND.BOF) Then
    Do Until rsND.EOF = True
        
        m = rsND("QueryName").Value
        Call NonDateQueries(c, m, aT)
        
        rsND.MoveNext
    Loop
Else
    MsgBox "There are no records in the recordset."
End If
                  
rsND.Close
dB.Close
Set rsND = Nothing
Set dB = Nothing

End Sub


Public Sub NonDateQueries(ByVal c As String, ByVal m As String, ByVal aT As String)

Dim pK As String 'primary key for the metric listed in the api query table.

    Dim URL As String
    Dim bodySend As String
    Dim MetricName As String
    Dim Metric As String
    Dim MetricParam As String
    Dim ProfID As String
    Dim ParamModifier As String
    
    MetricParam = DLookup("[MetricID]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    Metric = DLookup("[Parameter2]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    
    
    Dim TabN As String
    Dim FieldN As String

    TabN = DLookup("[TableName]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    FieldN = DLookup("[FieldName]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.

    ProfID = c

    Debug.Print MetricParam
    Debug.Print Metric
    Debug.Print ProfID


    
               
    URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & ProfID & "&metrics=" & MetricParam & "&start-date=" & "2013-08-01" & "&end-date=" & "2013-08-31"
    
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
         
    objhttp.Open "GET", URL, False
    objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
    objhttp.setRequestHeader "Host", "googleapis.com"
    objhttp.setRequestHeader "Authorization", "Bearer " & aT
    objhttp.send
        
    Dim QueryResult As String
    QueryResult = objhttp.responseText
    
    Debug.Print QueryResult
    
    Dim sD As Long 'start of data point to be extracted
    Dim eD As Long 'end of data point to be extracted
    Dim ElemTag As String 'the string the InStr function searches for within json
    Dim CharDiff As Long
    Dim dP1 As Variant
    Dim ElemTag2 As String
    
    ElemTag = "rows"
    ElemTag2 = Metric

    sD = InStrRev(QueryResult, ElemTag)
    '(1, json, ElemTag)
    sD = sD + 9
    If sD = 9 Then
        sD = InStrRev(QueryResult, ElemTag2)
    End If
    
    
    eD = InStr(sD, QueryResult, "]]")
    eD = eD - 1
    CharDiff = eD - sD
    
    Debug.Print sD
    Debug.Print eD
    Debug.Print CharDiff
    
    dP1 = Mid(QueryResult, sD, CharDiff)
    
    Debug.Print dP1
    Debug.Print m
    
    Dim ImpTable As String
    Dim ImpField As String

    ImpTable = DLookup("[TableName]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    ImpField = DLookup("[FieldName]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'") 'Looks up the Parameter2 for current query. This variable is used to handle the response, not in the actual query.
    pK = DLookup("[PrimKey]", "tblNonDateApiQueries", "[QueryName] = " & "'" & m & "'")
    
    Debug.Print "Import to Table" & ImpTable
    Debug.Print "Import to Field" & ImpField
    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    'rs = dB.OpenRecordset(ImpTable)
    
    Debug.Print "Data to import:" & data
    
    Dim ImpSQL As String
    ImpSQL = "UPDATE " & ImpTable & " SET " & ImpField & " = '" & dP1 & "' WHERE Source = '" & pK & "'"
    
    Debug.Print ImpSQL
    
    dB.Execute ImpSQL
    
        


End Sub
Public Sub DimTest()

Call DimDataQueries("tblDimApiQueries", "ReferralRevAndTraff", "2108448", "ya29.AHES6ZQq-AgTdDevfAkyRgmKnI44ldT-y_3Wj1M5kYghHwo")

End Sub

Public Sub DimDataQueries(qT As String, m As String, c As String, aT As String)
    'qT is the dimensionalized queries table
    'm is the metric name
    'c is the client profile ID


    Dim URL As String
    Dim bodySend As String
    Dim MetricName As String
    Dim QueryString As String
    Dim ProfID As String
    Dim MainAPI As String
    Dim RespNum As Integer
    
    MainAPI = qT
    
    RespNum = DLookup("[RespNum]", MainAPI, "[queryname] = " & "'" & m & "'")
    RespNum = RespNum - 2 'subtracting one from the number of data points to account for no trailing comma at end of data set
    
    QueryString = DLookup("[QueryParam]", MainAPI, "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    Debug.Print QueryString
    
    
    Dim TabN As String
    Dim FieldN1 As String
    Dim FieldN2 As String
    Dim FieldN3 As String

    TabN = DLookup("[TableName]", MainAPI, "[QueryName] = " & "'" & m & "'") 'Looks up the metric id for current query
    FieldN1 = DLookup("[FieldName1]", MainAPI, "[QueryName] = " & "'" & m & "'") 'Looks up the fieldname1 for the current query.
    FieldN2 = DLookup("[FieldName2]", MainAPI, "[QueryName] = " & "'" & m & "'") 'Looks up the fieldname2 for the current query.
    FieldN3 = DLookup("[FieldName3]", MainAPI, "[QueryName] = " & "'" & m & "'")
    
    Debug.Print m
    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblDateControls")
    
    Dim Sdate As Date
    Dim Edate As Date

    Dim GAStartDate As String
    Dim GAEndDate As String
            
    Sdate = rs("StartDate").Value
    Edate = rs("EndDate").Value
    
    GAStartDate = Format(Sdate, "yyyy-mm-dd")
    GAEndDate = Format(Edate, "yyyy-mm-dd")
    
    ProfID = c
    
    
    'This adds the Stella & Dot segment if the profile ID matches the Master Profile id
    If ProfID = "24294295" Then
        QueryString = QueryString & "&segment=gaid%3A%3AmyvdFr97RnWuty5z5nNL5Q"
    End If
                       
    URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & ProfID & "&" & QueryString & "&start-date=" & GAStartDate & "&end-date=" & GAEndDate
            
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            Dim QueryResult As String
            QueryResult = objhttp.responseText
            
            
            
    Dim objJson As MSScriptControl.ScriptControl
    Set objJson = New MSScriptControl.ScriptControl
    objJson.Language = "JScript"
    
    Dim objResp As Object
    Set objResp = objJson.Eval("(" & QueryResult & ")")
    
    Dim Response As String
    Response = objResp.rows
    
    Dim NumberOfCommas As Integer
    Dim NumberOfRows As Integer
    NumberOfCommas = Len(Response) - Len(Replace(Response, ",", ""))
    NumberOfRows = (NumberOfCommas + 1) / 2
    NumberOfRows = NumberOfRows - 2
    Debug.Print "Number of ROWSSSS"
    Debug.Print NumberOfRows
    
    If NumberOfRows < RespNum Then
        RespNum = NumberOfRows
    Else
        Debug.Print "no data deficiency errors here."
    End If
    
    Debug.Print Response
                
    Dim rs2 As Recordset
    Set rs2 = dB.OpenRecordset(TabN)
    
    'This is for dimensional queries with 2 fields
    
    
    Dim CharToNextComma As Integer
    Dim DataStr1 As String
    Dim DataStr2 As String
    Dim DataStr3 As String
    Dim Data1 As String
    Dim Data2 As Double
    Dim Data3 As Double
    
    dB.Execute "DELETE * FROM " & TabN
       
    Dim I
    For I = 0 To RespNum
        CharToNextComma = InStr(1, Response, ",")
        Debug.Print CharToNextComma
        DataStr1 = Mid(Response, 1, CharToNextComma - 1)
        Response = Mid(Response, CharToNextComma + 1, Len(Response))
        CharToNextComma = InStr(1, Response, ",")
        DataStr2 = IIf(CharToNextComma > 0, Mid(Response, 1, CharToNextComma - 1), Response)
        Response = Mid(Response, CharToNextComma + 1, Len(Response))
        If DataStr1 = "(none)" Then
            Data1 = "direct"
        Else
            Data1 = DataStr1
        End If
        Data2 = CDbl(DataStr2)
        
        If FieldN3 = "NA" Then
            Debug.Print "No third field for this metric"
        Else
            CharToNextComma = InStr(1, Response, ",")
            DataStr3 = IIf(CharToNextComma > 0, Mid(Response, 1, CharToNextComma - 1), Response)
            Response = Mid(Response, CharToNextComma + 1, Len(Response))
            Data3 = CDbl(DataStr3)
        End If
        
        Debug.Print Data1
        Debug.Print Data2
        Debug.Print Data3
        
        rs2.AddNew
            rs2("IndexNumber").Value = I + 1
            rs2(FieldN1).Value = Data1
            rs2(FieldN2).Value = Data2
            
            If FieldN3 = "NA" Then
                Debug.Print "No third field for this query"
            Else
                rs2(FieldN3).Value = Data3
            End If
        rs2.Update
        Debug.Print Response
    Next I
        
    CharToNextComma = InStr(1, Response, ",")
    Debug.Print CharToNextComma
    DataStr1 = Mid(Response, 1, CharToNextComma - 1)
    Response = Mid(Response, CharToNextComma + 1, Len(Response))
    If FieldN3 = "NA" Then
        DataStr2 = Response
    Else
        CharToNextComma = InStr(1, Response, ",")
        DataStr2 = Mid(Response, 1, CharToNextComma - 1)
        Debug.Print DataStr2
        Response = Mid(Response, CharToNextComma + 1, Len(Response))
        DataStr3 = Response
        Data3 = CDbl(DataStr3)
    End If
        
    If DataStr1 = "(none)" Then
        Data1 = "direct"
    Else
        Data1 = DataStr1
    End If
    Data2 = CDbl(DataStr2)
    
    Debug.Print Data1
    Debug.Print Data2
    
    rs2.AddNew
        rs2("IndexNumber").Value = I + 1
        rs2(FieldN1).Value = Data1
        rs2(FieldN2).Value = Data2
        If FieldN3 = "NA" Then
            Debug.Print "No third field."
        Else
            rs2(FieldN3).Value = Data3
        End If
    rs2.Update
    
    
    
    
                                
    rs.Close
    rs2.Close
    dB.Close
    Set rs = Nothing
    Set rs2 = Nothing
    Set dB = Nothing



End Sub

Public Sub TestOrgLandingPageSub()

Call AllOrgLandingPages(3199261, "ya29.1.AADtN_X5zZiA-4iACFnOy5Ry12qHiKvyoXj2NuSOG_jgtDkC7Te9htg6-CpscA")

End Sub


Public Sub AllOrgLandingPages(c As String, aT As String)
    'qT is the dimensionalized queries table
    'm is the metric name
    'c is the client profile ID


    Dim URL As String
    Dim bodySend As String
    Dim ProfID As String

    
    Dim dB As Database
    Dim rs As Recordset
    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblDateControls")
    
    Dim Sdate As Date
    Dim Edate As Date

    Dim GAStartDate As String
    Dim GAEndDate As String
            
    Sdate = rs("StartDate").Value
    Edate = rs("EndDate").Value
    
    GAStartDate = Format(Sdate, "yyyy-mm-dd")
    GAEndDate = Format(Edate, "yyyy-mm-dd")
    
    ProfID = c
    
    'This adds the Stella & Dot segment if the profile ID matches the Master Profile id
    If ProfID = "24294295" Then
        QueryString = QueryString & "&segment=gaid%3A%3AywNuo6PAQ7OiRK5lOUKN_w"
    End If
                       
    URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & ProfID & "&dimensions=ga:landingPagePath&filters=ga:medium%3D%3Dorganic&sort=-ga:visits" & "&start-date=" & GAStartDate & "&end-date=" & GAEndDate
            
    Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            Dim QueryResult As String
            QueryResult = objhttp.responseText
            
            
            
    Dim objJson As MSScriptControl.ScriptControl
    Set objJson = New MSScriptControl.ScriptControl
    objJson.Language = "JScript"
    
    Dim objResp As Object
    Set objResp = objJson.Eval("(" & QueryResult & ")")
    
    Dim Response As String
    Response = objResp.rows
    
    Debug.Print Response
                
                
    Dim rs2 As Recordset
    Set rs2 = dB.OpenRecordset("tblGAOrgAllLandingPages")
    
    Dim CharToNextComma As Integer
    Dim DataStr1 As String
    Dim DataStr2 As String
    Dim DataStr3 As String
    Dim Data1 As String
    Dim Data2 As Double
    Dim Data3 As Double
    
    dB.Execute "DELETE * FROM tblGAOrgAllLandingPages"
       
    Dim NumberOfResults As Integer
    NumberOfResults = Len(Response) - Len(Replace(Response, ",", ""))
       
    Dim I
    For I = 0 To NumberOfResults
        CharToNextComma = InStr(1, Response, ",")
        Debug.Print CharToNextComma
        DataStr1 = Mid(Response, 1, CharToNextComma - 1)
        Response = Mid(Response, CharToNextComma + 1, Len(Response))
        CharToNextComma = InStr(1, Response, ",")
        DataStr2 = IIf(CharToNextComma > 0, Mid(Response, 1, CharToNextComma - 1), Response)
        Response = Mid(Response, CharToNextComma + 1, Len(Response))
        Data1 = DataStr1
        
        Debug.Print Data1
        
        rs2.AddNew
            rs2("LandingPagePath").Value = Data1
        rs2.Update
        Debug.Print Response
    Next I

    Data1 = Response
    
    Debug.Print Data1
    
    rs2.AddNew
        rs2("LandingPagePath").Value = Data1
    rs2.Update
    
                                
    rs.Close
    rs2.Close
    dB.Close
    Set rs = Nothing
    Set rs2 = Nothing
    Set dB = Nothing



End Sub





Public Sub LandingPageCounts(ByVal c As String, ByVal aT As String)
'Queries GA API for list of all organic landing pages from the reporting month, last month, and last year
'Counts and returns the number of organic landing pages in that list
'Inputs:
'   c --> the client google analytics profile view ID as a string
'   aT --> the most recent Oauth2.0 authentication token which grants access to GA data via the query

Dim dB As Database
Dim rs As Recordset

Set dB = CurrentDb  'Creates the dB variable that references the database that is currently open
Set rs = dB.OpenRecordset("tblDateControls")    'opens the tblDateControls table so that it can be referenced/edited


'The following variables are used to store dates from tblDateControls.
Dim Sdate As Date   'reporting month start date.
Dim Edate As Date   'reporting month start date.
Dim Sdate2 As Date  'last months start date.
Dim Edate2 As Date  'last months end date.
Dim Sdate3 As Date  'last year start date.
Dim Edate3 As Date  'last year end date.

'The following variables are used to store the above date variables as strings, so that they can be added to the GA queries.
Dim GAStartDate As String   'reporting month start date
Dim GAEndDate As String     'reporting month end date
Dim LastMonthStartDate As String 'last months start date
Dim LastMonthEndDate As String  'last months end date
Dim LastYearMonthStartDate As String 'last years start date
Dim LastYearMonthEndDate As String  'last years end date
        
'The following collect and store dates from tblDateControls in a given variable
Sdate = rs("StartDate").Value
Edate = rs("EndDate").Value
Sdate2 = rs("Month1Start").Value
Edate2 = rs("Month1End").Value
Sdate3 = rs("Month12Start").Value
Edate3 = rs("Month12End").Value


'The following convert the above dates to strings and format them to be correctly processed by the GA API
GAStartDate = Format(Sdate, "yyyy-mm-dd")
GAEndDate = Format(Edate, "yyyy-mm-dd")
LastMonthStartDate = Format(Sdate2, "yyyy-mm-dd")
LastMonthEndDate = Format(Edate2, "yyyy-mm-dd")
LastYearMonthStartDate = Format(Sdate3, "yyyy-mm-dd")
LastYearMonthEndDate = Format(Edate3, "yyyy-mm-dd")


'Clears any previous data from the output data table for this subroutine
dB.Execute "DELETE * FROM tblGALandingPageTotals"

'Queries data for the reporting month.
'Returns all organic landing pages.
'Response limit is 1000.
'If response limit is reached, a warning will pop up that this data will need to be manually adjusted.
Dim URL As String
URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & c & _
        "&metrics=ga:visits&dimensions=ga:landingPagePath&filters=ga:medium%3D%3Dorganic&sort=-ga:visits" & _
        "&start-date=" & GAStartDate & _
        "&end-date=" & GAEndDate            'concatenates the GA query request
Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")  'initiates an internet connection as an object
                 
            objhttp.Open "GET", URL, False  'opens a GET request and adds the constructed query URL
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"   'sets the request header
            objhttp.setRequestHeader "Host", "googleapis.com"   'sets another request header
            objhttp.setRequestHeader "Authorization", "Bearer " & aT    'sets another request header that includes the OAuth 2.0 token
            objhttp.send    'sends the constructed request
                
            Dim qR As String
            qR = objhttp.responseText   'sets the qR variable to the text GA API responds with.

Call NumberOfJSONRows(qR, "tblGALandingPageTotals", "MonthIdentifier", "LandingPageResults", "ReportingMonth", "Month1")    'calls the function that handles the response


'The month before the reporting months landing page count
URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & c & "&metrics=ga:visits&dimensions=ga:landingPagePath&filters=ga:medium%3D%3Dorganic&sort=-ga:visits" & "&start-date=" & LastMonthStartDate & "&end-date=" & LastMonthEndDate

Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            qR = objhttp.responseText

Call NumberOfJSONRows(qR, "tblGALandingPageTotals", "MonthIdentifier", "LandingPageResults", "ReportingMonth", "Month2")


'The month a year before the reporting months landing page count
URL = "https://www.googleapis.com/analytics/v3/data/ga/?ids=ga:" & c & "&metrics=ga:visits&dimensions=ga:landingPagePath&filters=ga:medium%3D%3Dorganic&sort=-ga:visits" & "&start-date=" & LastYearMonthStartDate & "&end-date=" & LastYearMonthEndDate

Set objhttp = CreateObject("MSXML2.ServerXMLHTTP")
                 
            objhttp.Open "GET", URL, False
            objhttp.setRequestHeader "GET", "/analytics/v3/data/ga/ HTTP/1.1"
            objhttp.setRequestHeader "Host", "googleapis.com"
            objhttp.setRequestHeader "Authorization", "Bearer " & aT
            objhttp.send
                
            qR = objhttp.responseText

Call NumberOfJSONRows(qR, "tblGALandingPageTotals", "MonthIdentifier", "LandingPageResults", "ReportingMonth", "Month3")

End Sub


Public Sub BrandFilter(ByVal m As String, URL As String)

    'The following code block conditionally adds branded/nonbranded filtering to traffic based on tblClients input brand term
            
    If m = "notprovBranded" Then
        URL = URL & "&filters=ga:keyword%3D@" & bT
    End If
    
    If m = "notprovNonBranded" Then
        URL = URL & "&filters=ga:keyword!@" & bT
    End If
    
    If m = "notprovRevBranded" Then
        URL = URL & "&filters=ga:keyword%3D@" & bT
    End If
    
    If m = "notprovRevNonBranded" Then
        URL = URL & "&filters=ga:keyword!@" & bT
    End If
    
    
    


End Sub


Function Get_Variable_Type(myVar)

' ---------------------------------------------------------------
' Written By Shanmuga Sundara Raman for http://vbadud.blogspot.com
' ---------------------------------------------------------------

If VarType(myVar) = vbNull Then
MsgBox "Null (no valid data) "
ElseIf VarType(myVar) = vbInteger Then
MsgBox "Integer "
ElseIf VarType(myVar) = vbLong Then
MsgBox "Long integer "
ElseIf VarType(myVar) = vbSingle Then
MsgBox "Single-precision floating-point number "
ElseIf VarType(myVar) = vbDouble Then
MsgBox "Double-precision floating-point number "
ElseIf VarType(myVar) = vbCurrency Then
MsgBox "Currency value "
ElseIf VarType(myVar) = vbDate Then
MsgBox "Date value "
ElseIf VarType(myVar) = vbString Then
MsgBox "String "
ElseIf VarType(myVar) = vbObject Then
MsgBox "Object "
ElseIf VarType(myVar) = vbError Then
MsgBox "Error value "
ElseIf VarType(myVar) = vbBoolean Then
MsgBox "Boolean value "
ElseIf VarType(myVar) = vbVariant Then
MsgBox "Variant (used only with arrays of variants) "
ElseIf VarType(myVar) = vbDataObject Then
MsgBox "A data access object "
ElseIf VarType(myVar) = vbDecimal Then
MsgBox "Decimal value "
ElseIf VarType(myVar) = vbByte Then
MsgBox "Byte value "
ElseIf VarType(myVar) = vbUserDefinedType Then
MsgBox "Variants that contain user-defined types "
ElseIf VarType(myVar) = vbArray Then
MsgBox "Array "
Else
MsgBox VarType(myVar)
End If

' Excel VBA, Visual Basic, Get Variable Type, VarType


End Function
