Sub Now()
'Inser current date and time in 2 different cells

Dim data_range As String, time_range As String

' celle dove inserire data e ora
data_range = "C" & ActiveCell.row
time_range = "D" & ActiveCell.row

Range(data_range) = Date
Range(time_range) = Time

End Sub
