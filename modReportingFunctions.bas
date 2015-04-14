Attribute VB_Name = "modReportingFunctions"
Option Compare Database

Public Sub temporarystupid()

    Dim FN As String
    Dim DTN As String
    
    FN = "TotalTraffic"
    DTN = "tblGAMainKPIs"
    
    'Call m2m(FN, DTN)
    'Call y2y(FN, DTN)
    Call AddMonthCalcs

End Sub


Public Function m2m(FieldName, DataTableName) As Double
'Takes a given month's input and divides it by the last month's data point FieldName, DataTableName

    Dim Month1 As Date
    Dim Month2 As Date
    
    Month1 = DLookup("[EndDate]", "tblDateControls")
    Month2 = DLookup("[Month1End]", "tblDateControls")
    
    Dim Data1 As Long
    Dim Data2 As Long
        
    Data1 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month1 & "#")
    Data2 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month2 & "#")

 
    Dim Result As Double
    Result = Data1 / Data2
    
    Debug.Print Result
    
End Function

Public Function y2y(FieldName, DataTableName) As Double
'Takes a given month's input and divides it by that month last year

    Dim Month1 As Date
    Dim Month2 As Date
    
    Month1 = DLookup("[EndDate]", "tblDateControls")
    Month2 = DLookup("[Month12End]", "tblDateControls")
    
    Dim Data1 As Long
    Dim Data2 As Long
        
    Data1 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month1 & "#")
    Data2 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month2 & "#")

 
    Dim Result As Double
    Result = Data1 / Data2
    
    Debug.Print Result
    
End Function

Public Sub AddMonthCalcs(ByVal API1 As String, ByVal Api2 As String, ByVal Api3 As String)

    Dim dB As Database
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim rs3 As Recordset
    Dim rs4 As Recordset
    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblCalculations")
    Set rs2 = dB.OpenRecordset(API1)
    If Api2 = "" Then
        Debug.Print "no specific api queries for this client."
    Else
        Set rs3 = dB.OpenRecordset(Api2)
    End If
    Set rs4 = dB.OpenRecordset(Api3)
    
    
    dB.Execute "DELETE * FROM tblCalculations"
    
    Dim MetricName As String
    Dim FieldName As String
    Dim DataTableName As String
        
    If Not (rs2.EOF And rs2.BOF) Then
        Do Until rs2.EOF = True
        
            MetricName = rs2("QueryName").Value
            FieldName = rs2("FieldName").Value
            DataTableName = rs2("TableName").Value
            
            Debug.Print MetricName
            Debug.Print FieldName
            Debug.Print DataTableName
                        
            Dim Month1 As Date
            Dim Month2 As Date

            Month1 = DLookup("[EndDate]", "tblDateControls")
            Month2 = DLookup("[Month1End]", "tblDateControls")

            Dim Data1 As Double
            Dim Data2 As Double
        
            Data1 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month1 & "#")
            Data2 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month2 & "#")
    
            Dim Year1 As Date
            Dim Year2 As Date
    
            Year1 = DLookup("[EndDate]", "tblDateControls")
            Year2 = DLookup("[Month12End]", "tblDateControls")
    
            Dim Data3 As Double
            Dim Data4 As Double
    
            Data3 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year1 & "#")
            Data4 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year2 & "#")
    
            
   
            rs.AddNew
            rs("MetricName").Value = MetricName
            rs("Month1").Value = Data1
            rs("Month2").Value = Data2
            rs("Year1").Value = Data3
            rs("Year2").Value = Data4
            rs.Update
    
            rs2.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    If Api2 = "" Then
        Debug.Print "no client specific api queries"
    Else
        If Not (rs3.EOF And rs3.BOF) Then
            Do Until rs3.EOF = True
            
                MetricName = rs3("QueryName").Value
                FieldName = rs3("FieldName").Value
                DataTableName = rs3("TableName").Value
                
                Debug.Print MetricName
                Debug.Print FieldName
                Debug.Print DataTableName
                            
                'removed dim month1 and month2
    
                Month1 = DLookup("[EndDate]", "tblDateControls")
                Month2 = DLookup("[Month1End]", "tblDateControls")
    
                'removed dim data1 and data2
            
                Data1 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month1 & "#")
                Data2 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month2 & "#")
        
                'removed dim year1 and year2
        
                Year1 = DLookup("[EndDate]", "tblDateControls")
                Year2 = DLookup("[Month12End]", "tblDateControls")
        
                'removed dim data3 and data4
        
                Data3 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year1 & "#")
                Data4 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year2 & "#")
        
        
       
                rs.AddNew
                rs("MetricName").Value = MetricName
                rs("Month1").Value = Data1
                rs("Month2").Value = Data2
                rs("Year1").Value = Data3
                rs("Year2").Value = Data4
                rs.Update
        
                rs3.MoveNext
            Loop
        Else
            MsgBox "There are no records in the recordset."
        End If
    End If
    
    If Not (rs4.EOF And rs4.BOF) Then
        Do Until rs4.EOF = True
        
            MetricName = rs4("QueryName").Value
            FieldName = rs4("FieldName").Value
            DataTableName = rs4("TableName").Value
            
            Debug.Print MetricName
            Debug.Print FieldName
            Debug.Print DataTableName
                        

            Month1 = DLookup("[EndDate]", "tblDateControls")
            Month2 = DLookup("[Month1End]", "tblDateControls")

        
            Data1 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month1 & "#")
            Data2 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Month2 & "#")
    
    
            Year1 = DLookup("[EndDate]", "tblDateControls")
            Year2 = DLookup("[Month12End]", "tblDateControls")
    
    
            Data3 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year1 & "#")
            Data4 = DLookup("[" & FieldName & "]", DataTableName, "[" & DataTableName & "]![MonthEndDate] = #" & Year2 & "#")
    
            
   
            rs.AddNew
            rs("MetricName").Value = MetricName
            rs("Month1").Value = Data1
            rs("Month2").Value = Data2
            rs("Year1").Value = Data3
            rs("Year2").Value = Data4
            rs.Update
    
            rs4.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
    
    rs.Close
    rs2.Close
    rs4.Close
    Set rs = Nothing
    Set rs2 = Nothing
    Set rs4 = Nothing
    
    If Api2 = "" Then
        Debug.Print "No client specific api queries"
    Else
        rs3.Close
        Set rs3 = Nothing
    End If
End Sub
