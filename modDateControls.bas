Attribute VB_Name = "modDateControls"
Option Compare Database

Public Sub DateController(ByVal sD As Date, ByVal eD As Date)

    Dim dB As Database
    Dim rs As Recordset
                    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset("tblDateControls")
    
    dB.Execute "DELETE * FROM tblDateControls" 'clears tblDateControls' data from any previous report. This table should only every have one record in it.
    
    rs.AddNew
    eD = LastDayInMonth(eD) 'Converts User Input for End Date as last day of that month
    rs("EndDate").Value = eD  'Adds this to the table
    For MonthNum = 1 To 12 'For Loop that adds the last day of every month to tblDateControls for the 12 months prior to the user input end date
        eD = DateAdd("m", -1, [eD]) 'decriments date by 1 month.
        eD = LastDayInMonth(eD) 'Makes this value the last date of the month
        rs("Month" & MonthNum & "End").Value = eD 'adds this date to the table
    Next MonthNum
    sD = FirstDayInMonth(sD)
    rs("StartDate").Value = sD
    For MonthNum2 = 1 To 12
        sD = DateAdd("m", -1, [sD])
        rs("Month" & MonthNum2 & "Start").Value = sD
    Next MonthNum2
    rs.Update
        
End Sub

Public Sub DateUpdate(ByVal sD As Date, ByVal eD As Date, ByVal Api As String)
    'This sub-routine clears all data from the listed input tables from the API queries table(s)
    'and adds updated date values to each table.

    Dim dB As Database
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim rs3 As Recordset
                    
    Set dB = CurrentDb
    Set rs = dB.OpenRecordset(Api)
    Set rs3 = dB.OpenRecordset("tblDateControls")
    
    If Not (rs.EOF And rs.BOF) Then
        Do Until rs.EOF = True
            
            gaTable = rs("TableName").Value
            dB.Execute "DELETE * FROM " & gaTable
            
            Set rs2 = dB.OpenRecordset(gaTable)
            
            Debug.Print gaTable
            Debug.Print eD
            Debug.Print sD
            
            rs2.AddNew
            rs2("MonthEndDate").Value = eD
            rs2("MonthStartDate").Value = sD
            rs2("MonthIndex").Value = 1
            rs2.Update
            
            Dim mN As String 'stores a string containing the field name for a given months date from the date controller table
            Dim mS As String 'month start date
                    
            For MonthNum = 1 To 12
                mN = "Month" & MonthNum & "End"
                mS = "Month" & MonthNum & "Start"
                mI = MonthNum + 1 'this adds 1 to the current MonthNum value to align the MonthIndex to the sequential number of months. 1 is already assigned as the most recent month.
                rs2.AddNew
                rs2("MonthEndDate").Value = rs3(mN).Value
                rs2("MonthStartDate").Value = rs3(mS).Value
                rs2("MonthIndex").Value = mI
                rs2.Update
            Next MonthNum
                    
            rs.MoveNext
        Loop
    Else
        MsgBox "There are no records in the recordset."
    End If
        

End Sub

Public Function FirstDayInMonth(Optional dtmDate As Variant) As Date
    ' Return the first day in the specified month.
    
    ' In:
    '   dtmDate:
    '       The specified date.
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       The date of the first day in the specified month.
    ' Example:
    '   FirstDayInMonth(#5/7/70#) returns 5/1/70.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    
    ' Note that IsMissing only works for Variant types, so
    ' the parameter must be a Variant for this method to work.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInMonth = DateSerial( _
     Year(dtmDate), Month(dtmDate), 1)
End Function


Public Function LastDayInMonth(Optional dtmDate As Variant) As Date
    ' Return the last day in the specified month.
    
    ' In:
    '   dtmDate:
    '       The specified date
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       The date of the last day in the specified month.
    ' Comments:
    '   This function counts on odd behavior of dateSerial. That is, each of the
    '   numeric values can be an expression containing a relative value. Here, the
    '   Day value becomes 1 - 1 (that is, the day before the first day of the month).
    ' Example:
    '   LastDayInMonth(#5/7/70#) returns 5/1/70.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInMonth = DateSerial( _
     Year(dtmDate), Month(dtmDate) + 1, 0)
End Function

Public Function FirstDayInWeek(Optional dtmDate As Variant) As Date
    ' Returns the first day in the week specified by the
    ' date in dtmDate. Uses localized settings for the first
    ' day of the week.
    
    ' In:
    '   dtmDate:
    '       date specifying the week in which to work.
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       First day of the specified week, taking into account the
    '       user's locale.
    ' Example:
    '   FirstDayInWeek(#5/12/2010#) returns 5/9/2010 in the US.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInWeek = dtmDate - _
     Weekday(dtmDate, vbUseSystemDayOfWeek) + 1
End Function

Public Function LastDayInWeek(Optional dtmDate As Variant) As Date
    ' Returns the last day in the week specified by the date in dtmDate.
    ' Uses localized settings for the first day of the week.
    
    ' In:
    '   dtmDate:
    '       date specifying the week in which to work.
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       Last day of the specified week, taking into account the
    '       user's locale.
    ' Example:
    '   LastDayInWeek(#4/1/97#) returns 4/5/97 in the US.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInWeek = dtmDate - _
     Weekday(dtmDate, vbUseSystemDayOfWeek) + 7
End Function

Public Function FirstDayInYear(Optional dtmDate As Variant) As Date
    ' Return the first day in the specified year.
    
    ' In:
    '   dtmDate:
    '       The specified date
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       The date of the first day in the specified year.
    ' Example:
    '   FirstDayInYear(#5/7/1970#) returns 1/1/1970.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    FirstDayInYear = DateSerial(Year(dtmDate), 1, 1)
End Function

Public Function LastDayInYear(Optional dtmDate As Variant) As Date
    ' Return the last day in the specified year.
    
    ' In:
    '   dtmDate (Optional)
    '       The specified date
    '       Use the current date, if none was specified.
    ' Out:
    '   Return Value:
    '       The date of the last day in the specified year.
    ' Example:
    '   LastDayInYear(#5/7/1970#) returns 12/31/1970.
    
    ' Did the caller pass in a date? If not, use
    ' the current date.
    If IsMissing(dtmDate) Then
        dtmDate = Date
    End If
    
    LastDayInYear = DateSerial(Year(dtmDate), 12, 31)
End Function


