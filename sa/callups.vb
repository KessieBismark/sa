Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Security.Cryptography
Imports MySql.Data.MySqlClient

Module callups


#Region "Call ups declaration"
    'Public server, pw, userid, db As String

    ' Public cs As String = "server='localhost'; userid='root'; password='';  database='payroll_rf'"
    Public cs As String = "SERVER=199.127.63.79;PORT=3306;DATABASE=royalfo5_payroll_royal;UID=royalfo5_shop;PASSWORD=2dL^7.j~#.2Y"
    Dim info As String
    Dim day, sats, sundays, working_daya As Integer
    Public academic_year As String
    Public term As String
    Dim mm As New Integer
    Public sid As String
    Public ul, tclass, staffid, staff_name As String
    Public show_conn_status As Boolean = False
    Public enc As System.Text.UTF8Encoding
    Public encryptor As ICryptoTransform
    Public decryptor As ICryptoTransform
    Public up As Integer = 0
    Public AppName As String
    Public uid As String
    Public ip, port As String

    Dim daym As Int16 = DateTime.DaysInMonth(Now.Date.Year, Now.Date.Month)

#End Region

#Region "Declarations"
    Public con As MySqlConnection
    Public cmd, cmd2, cmd3, cmd4, cmd5, cmd6, cmd7, cmd8, cmd9, cmdb As MySqlCommand
    Public adp, adp2, adp4, adp3, adp5, adp6, adpb As MySqlDataAdapter
    Public tb, tb2, tb3, tb4, tb5, tb6, tbb As DataTable
    Public ds, ds2, ds3, ds4 As DataSet
    Public querryString As String
#End Region

#Region "Encription and decription"
    Public Function Encript(ByVal value As String)
        Dim KEY_128 As Byte() = {42, 1, 52, 67, 231, 13, 94, 101, 123, 6, 0, 12, 32, 91, 4, 111, 31, 70, 21, 141, 123, 142, 234, 82, 95, 129, 187, 162, 12, 55, 98, 23}
        Dim IV_128 As Byte() = {234, 12, 52, 44, 214, 222, 200, 109, 2, 98, 45, 76, 88, 53, 23, 78}
        Dim symmetricKey As RijndaelManaged = New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC

        enc = New System.Text.UTF8Encoding
        encryptor = symmetricKey.CreateEncryptor(KEY_128, IV_128)
        decryptor = symmetricKey.CreateDecryptor(KEY_128, IV_128)

        Dim sPlainText As String = value
        If Not String.IsNullOrEmpty(sPlainText) Then
            Dim memoryStream As MemoryStream = New MemoryStream()
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write)
            cryptoStream.Write(enc.GetBytes(sPlainText), 0, sPlainText.Length)
            cryptoStream.FlushFinalBlock()
            Return Convert.ToBase64String(memoryStream.ToArray())
            memoryStream.Close()
            cryptoStream.Close()
        End If
        Return ""
    End Function

    Public Function Decrypt(ByVal value As String)
        Dim KEY_128 As Byte() = {42, 1, 52, 67, 231, 13, 94, 101, 123, 6, 0, 12, 32, 91, 4, 111, 31, 70, 21, 141, 123, 142, 234, 82, 95, 129, 187, 162, 12, 55, 98, 23}
        Dim IV_128 As Byte() = {234, 12, 52, 44, 214, 222, 200, 109, 2, 98, 45, 76, 88, 53, 23, 78}
        Dim symmetricKey As RijndaelManaged = New RijndaelManaged()
        symmetricKey.Mode = CipherMode.CBC

        enc = New System.Text.UTF8Encoding
        encryptor = symmetricKey.CreateEncryptor(KEY_128, IV_128)
        decryptor = symmetricKey.CreateDecryptor(KEY_128, IV_128)
        If Not String.IsNullOrEmpty(value) Then
            Dim cypherTextBytes As Byte() = Convert.FromBase64String(value)
            Dim memoryStream As MemoryStream = New MemoryStream(cypherTextBytes)
            Dim cryptoStream As CryptoStream = New CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read)
            Dim plainTextBytes(cypherTextBytes.Length) As Byte
            Dim decryptedByteCount As Integer = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length)
            memoryStream.Close()
            cryptoStream.Close()
            Return enc.GetString(plainTextBytes, 0, decryptedByteCount)
        End If
        Return ""
    End Function
#End Region

#Region "Search Filter"
    Public Function filterdata(Valuetosearch As String, ByVal q As String)
        Try
            con = New MySqlConnection(cs)
            Dim search As String = q & " LIKE '%" & Valuetosearch & "%'"
            Dim cmd As New MySqlCommand(search, con)
            Dim da As New MySqlDataAdapter(cmd)
            Dim table As New DataTable
            da.Fill(table)
            Return table
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

#End Region



#Region "clear all textbox"
    Public Sub Clear(c As Control)
        Dim currentTextBox As TextBox
        For Each ctrl As Control In c.Controls
            If TypeOf ctrl Is TextBox Then
                currentTextBox = CType(ctrl, TextBox)
                currentTextBox.Clear()
            End If
            If ctrl.HasChildren Then
                Clear(ctrl)
            End If
        Next
    End Sub
#End Region

#Region "Day of the week"
    Function daysofweeks()
        Dim dw As String = ""
        Dim num As Integer
        num = Now.Date.DayOfWeek
        If num = 1 Then
            dw = "Monday"
        ElseIf num = 2 Then
            dw = "Tuesday"
        ElseIf num = 3 Then
            dw = "Wednesday"
        ElseIf num = 4 Then
            dw = "Thurday"
        ElseIf num = 5 Then
            dw = "Friday"
        End If
        Return dw
    End Function
#End Region

#Region "Count records in the table"
    Public Function record_count(ByVal a As String)
        Try
            con = New MySqlConnection(cs)
            cmd = New MySqlCommand(a, con)
            adp = New MySqlDataAdapter(cmd)
            tb = New DataTable
            adp.Fill(tb)
            Return tb.Rows.Count
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#End Region

#Region "get a string from a database table"
    Public Function getValue(ByVal query As String)
        Try
            con = New MySqlConnection(cs)
            con.Open()
            cmd = New MySqlCommand(query, con)
            Return cmd.ExecuteScalar
            con.Close()
        Catch ex As Exception
            Return "Null"
        End Try
    End Function
#End Region

#Region "get a number from a database table"
    Public Function getnumber(ByVal query As String)
        Try
            con = New MySqlConnection(cs)
            con.Open()
            cmd = New MySqlCommand(query, con)
            Return CDbl(cmd.ExecuteScalar)
            con.Close()
        Catch ex As Exception
            con.Close()
            Return 0
        End Try
    End Function
#End Region

#Region "get the beginning of the month date"
    Public Function frmdate()
        day = "1"
        Return Now.Date.Year & "-" & Now.Date.Month & "-" & day
    End Function
#End Region

#Region "get the end of the month date"
    Public Function todate()
        Dim m As New Integer
        m = Now.Date.Month
        If m = "1" Then
            day = "31"
        ElseIf m = "3" Then
            day = "31"
        ElseIf m = "5" Then
            day = "31"
        ElseIf m = "7" Then
            day = "31"
        ElseIf m = "8" Then
            day = "31"
        ElseIf m = "10" Then
            day = "31"
        ElseIf m = "12" Then
            day = "31"
        ElseIf m = "6" Then
            day = "30"
        ElseIf m = "4" Then
            day = "30"
        ElseIf m = "9" Then
            day = "30"
        ElseIf m = "11" Then
            day = "30"
        ElseIf m = "2" Then
            day = "28"
        End If
        Return Now.Date.Year & "-" & Now.Date.Month & "-" & day
    End Function
#End Region

#Region "total days in a month"
    Public Function TotalDays()
        Dim dy As Integer
        Dim m As New Integer
        m = Now.Date.Month - 1
        If m = 0 Then
            dy = 31
        ElseIf m = 12 Then
            dy = 31
        ElseIf m = 3 Then
            dy = 31
        ElseIf m = 5 Then
            dy = 31
        ElseIf m = 7 Then
            dy = 31
        ElseIf m = 8 Then
            dy = 31
        ElseIf m = 10 Then
            dy = 31
        ElseIf m = 6 Then
            dy = 30
        ElseIf m = 4 Then
            dy = 30
        ElseIf m = 9 Then
            dy = 30
        ElseIf m = 11 Then
            dy = 30
        ElseIf m = 2 Then
            dy = 28
        End If
        Return dy
    End Function
#End Region

#Region "get the beginning of last month date"
    Public Function LastMontFrmDate()
        day = "1"
        Dim m, y As New Integer
        m = Now.Date.Month - 1
        y = Now.Date.Year

        If m = 0 Then
            m = 12
            y = y - 1
        End If
        Return y & "-" & m & "-" & day
    End Function
#End Region

#Region "get user start month"
    Public Function UserStartMonth(ByVal m As Integer)
        day = "1"
        Dim y As New Integer
        y = Now.Date.Year
        Return y & "-" & m & "-" & day
    End Function
#End Region

#Region "get user end  month"
    Public Function UserEndMonth(ByVal m As Integer)
        Dim y As New Integer

        If m = "1" Then
            day = "31"
        ElseIf m = "3" Then
            day = "31"
        ElseIf m = "5" Then
            day = "31"
        ElseIf m = "7" Then
            day = "31"
        ElseIf m = "8" Then
            day = "31"
        ElseIf m = "10" Then
            day = "31"
        ElseIf m = "12" Then
            day = "31"
        ElseIf m = "6" Then
            day = "30"
        ElseIf m = "4" Then
            day = "30"
        ElseIf m = "9" Then
            day = "30"
        ElseIf m = "11" Then
            day = "30"
        ElseIf m = "2" Then
            day = "28"
        End If
        y = Now.Date.Year
        Return y & "-" & m & "-" & day
    End Function
#End Region

#Region "get the end of last month date"
    Public Function LastMonthEndDate()
        Dim m, y As New Integer
        m = Now.Date.Month - 1
        y = Now.Date.Year

        If m = 0 Then
            m = 12
            y = y - 1
        End If
        If m = "1" Then
            day = "31"
        ElseIf m = "3" Then
            day = "31"
        ElseIf m = "5" Then
            day = "31"
        ElseIf m = "7" Then
            day = "31"
        ElseIf m = "8" Then
            day = "31"
        ElseIf m = "10" Then
            day = "31"
        ElseIf m = "12" Then
            day = "31"
        ElseIf m = "6" Then
            day = "30"
        ElseIf m = "4" Then
            day = "30"
        ElseIf m = "9" Then
            day = "30"
        ElseIf m = "11" Then
            day = "30"
        ElseIf m = "2" Then
            day = "28"
        End If
        Return y & "-" & m & "-" & day
    End Function
#End Region

#Region "Save data"
    Public Sub execute(ByVal query As String, ByVal a As Integer)
        Try
            ' ping()
            con = New MySqlConnection(cs)
            con.Open()
            cmd = New MySqlCommand(query, con)
            cmd.ExecuteNonQuery()
            con.Close()
            If a = 0 Then
                MessageBox.Show("All Information was saved successfully", "Royal Foam", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf a = 1 Then
                MessageBox.Show("All Information was updated successfully", "Royal Foam", MessageBoxButtons.OK, MessageBoxIcon.Information)
            ElseIf a = 2 Then
                MessageBox.Show(" Information was deleted successfully", "Royal Foam", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            showerror(ex.Message)
        End Try
    End Sub
#End Region

#Region "display data"
    Public Function ShowData(ByVal query As String)
        Try
            con = New MySqlConnection(cs)
            cmd = New MySqlCommand(query, con)
            adp = New MySqlDataAdapter(cmd)
            tb = New DataTable
            adp.Fill(tb)
            Return tb
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#End Region

#Region "show error message"
    Public Sub showerror(ByVal s As String)
        MessageBox.Show(s, "Royal Foam : Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub
#End Region

#Region "show info message"
    Public Sub showinfo(ByVal s As String)
        MessageBox.Show(s, "Royal Foam", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
#End Region

#Region "Calculate days present"
    Public Sub present(ByVal calmonth As Integer, ByVal year As Integer)
        Try
            Dim d As Date = Date.Now
            Dim month As DateTime = d.AddMonths(-1).ToString("yyy-MM-dd")
            Dim sl As Decimal = 0
            tb = New DataTable
            tb = ShowData("SELECT * FROM employee ")
            adp.Fill(tb)
            For i As Integer = 0 To tb.Rows.Count - 1 Step +1
                Dim hr As Integer = getnumber("select sum(Hours) from attendance where Staff_ID= '" + tb.Rows(i)("Staff_ID").ToString + "' and DAYOFMONTH(Date) = '" & mm & "' and sid = '" & sid & "'")
                tb2 = New DataTable
                tb2 = ShowData("select * from attendance where Staff_ID= '" + tb.Rows(i)("Staff_ID").ToString + "' and DAYOFMONTH(Date) = '" & calmonth & "' and sid = '" & year & "'")
                Dim h As Integer = 0
                Dim p As Integer = 0
                Dim cut As Decimal = 0
                Dim df As Integer = 0
                Dim s As Decimal = 0
                Dim wk As Integer = 0
                Dim total_LeaveDays As Integer = 0
                h = tb.Rows(i)("hours").ToString
                If tb.Rows(0)("Start_Date").ToString = "" Or tb.Rows(0)("End_Date").ToString = "" Then
                    wk = h * Weekdays()
                    sl = tb.Rows(0)("salary").ToString
                    df = wk - hr
                    cut = sl / wk
                    s = sl - (cut * df)
                Else
                    Dim leaveStart As Date = tb.Rows(0)("Start_Date")
                    Dim leaveend As Date = tb.Rows(0)("End_Date")
                    If leaveStart.Month <= mm And academic_year = tb.Rows(0)("academic").ToString And leaveend.Month <= Now.Date.Month - 1 Then
                        Dim wkend As Integer = 0
                        If leaveStart.Month = mm And leaveend.Month = mm Then
                            wkend = Weekends(leaveStart, leaveend)
                            total_LeaveDays = (Val((leaveend - leaveStart).ToString("dd")) + 1) - wkend
                            hr = hr + (h * total_LeaveDays)
                        ElseIf leaveStart.Month = mm And Not leaveend.Month = mm Then
                            Dim EndLastDate As Date = LastMonthEndDate()
                            wkend = Weekends(leaveStart, EndLastDate)
                            total_LeaveDays = (Val((leaveend - leaveStart).ToString("dd")) + 1) - wkend
                            hr = hr + (h * total_LeaveDays)
                        Else
                            wkend = Weekends(LastMontFrmDate, LastMonthEndDate)
                            Dim StartLastDate As Date = LastMontFrmDate()
                            Dim EndLastDate As Date = LastMonthEndDate()
                            total_LeaveDays = (Val(((EndLastDate - StartLastDate).ToString("dd") + 1)) - wkend)
                            hr = h * total_LeaveDays
                        End If
                        wk = h * Weekdays()
                        sl = tb.Rows(0)("salary").ToString
                        df = wk - hr
                        cut = sl / wk
                        s = sl - (cut * df)
                    End If
                End If
                execute("update payroll set Days_Present='" & p & "',Total_Leave_Days='" & total_LeaveDays & "',Total_Hours='" & wk & "',Hours_Present='" & hr & "',Working_Days='" & Weekdays() & "',Payment='" & s & "' where Staff_ID = '" + tb.Rows(i)("Staff_ID").ToString + "' and Academic ='" & academic_year & "' and Month='" & mm & "' and sid = '" & sid & "' and NOT manual = 1", 3)
            Next i
            con.Close()
        Catch ex As Exception
            showerror(ex.Message)
        End Try
    End Sub
#End Region

#Region "calculate working days"
    Public Function Weekdays()
        Dim d As Date = Date.Now
        Dim startDate As Date = startmonth()
        Dim endDate As Date = finaldate()
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        Dim mm As Integer = Now.Date.Month - 1
        numWeekdays = 0
        WeekendDays = 0
        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1
        For i As Integer = 1 To totalDays
            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next
        numWeekdays = totalDays - WeekendDays
        'deducting public holidays
        If mm = 0 Then
            numWeekdays = numWeekdays - 3
        ElseIf mm = 1 Then
            numWeekdays = numWeekdays - 2
        ElseIf mm = 3 Then
            numWeekdays = numWeekdays - 1
        ElseIf mm = 4 Then
            numWeekdays = numWeekdays - 2
        ElseIf mm = 5 Then
            numWeekdays = numWeekdays - 3
        ElseIf mm = 7 Then
            numWeekdays = numWeekdays - 2
        ElseIf mm = 8 Then
            numWeekdays = numWeekdays - 1
        ElseIf mm = 9 Then
            numWeekdays = numWeekdays - 1
        End If
        Return numWeekdays
    End Function
#End Region

#Region "calculate number of weeKends"
    Public Function Weekends(ByVal st As Date, ByVal et As Date)
        Dim d As Date = Date.Now
        Dim month As DateTime = st
        Dim startDate As Date = month
        Dim endDate As Date = et
        Dim numWeekdays As Integer
        Dim totalDays As Integer
        Dim WeekendDays As Integer
        Dim mm As Integer = Now.Date.Month - 1
        numWeekdays = 0
        WeekendDays = 0
        totalDays = DateDiff(DateInterval.Day, startDate, endDate) + 1
        For i As Integer = 1 To totalDays
            If DatePart(DateInterval.Weekday, startDate) = 1 Then
                WeekendDays = WeekendDays + 1
            End If
            If DatePart(DateInterval.Weekday, startDate) = 7 Then
                WeekendDays = WeekendDays + 1
            End If
            startDate = DateAdd("d", 1, startDate)
        Next
        numWeekdays = WeekendDays
        'adding public holidays
        If mm = 0 Then
            numWeekdays = numWeekdays + 3
        ElseIf mm = 1 Then
            numWeekdays = numWeekdays + 2
        ElseIf mm = 3 Then
            numWeekdays = numWeekdays + 1
        ElseIf mm = 4 Then
            numWeekdays = numWeekdays + 2
        ElseIf mm = 5 Then
            numWeekdays = numWeekdays + 3
        ElseIf mm = 7 Then
            numWeekdays = numWeekdays + 2
        ElseIf mm = 8 Then
            numWeekdays = numWeekdays + 1
        ElseIf mm = 9 Then
            numWeekdays = numWeekdays + 1
        End If
        Return numWeekdays
    End Function
#End Region

#Region "final todate"
    Public Function finaldate()
        Dim tdd As Date
        mm = Now.Date.Month - 1
        Dim y As Integer = Now.Date.Year
        If mm = 0 Then
            mm = 12
            y = y - 1
        End If
        tdd = y & "-" & mm & "-" & TotalDays()
        Return tdd
    End Function
#End Region

#Region "month check"
    Public Function mcheck(ByVal day As Integer, ByVal mm As Integer)
        Dim tdd As Date
        Dim y As Integer = Now.Date.Year
        tdd = y & "-" & mm & "-" & day
        Return tdd
    End Function
#End Region

#Region "total days in a month"
    Public Function Daays(ByVal m As Integer)
        Dim dy As Integer
        If m = 1 Then
            dy = 31
        ElseIf m = 12 Then
            dy = 31
        ElseIf m = 3 Then
            dy = 31
        ElseIf m = 5 Then
            dy = 31
        ElseIf m = 7 Then
            dy = 31
        ElseIf m = 8 Then
            dy = 31
        ElseIf m = 10 Then
            dy = 31
        ElseIf m = 6 Then
            dy = 30
        ElseIf m = 4 Then
            dy = 30
        ElseIf m = 9 Then
            dy = 30
        ElseIf m = 11 Then
            dy = 30
        ElseIf m = 2 Then
            dy = 28
        End If
        Return dy
    End Function
#End Region

#Region "to get start month date"
    Public Function startmonth()
        Dim mm As New Integer
        Dim tdd As Date
        mm = Now.Date.Month - 1
        Dim y As Integer = Now.Date.Year
        If mm = 0 Then
            mm = 12
            y = y - 1
        End If
        tdd = y & "-" & mm & "-" & 1
        Return tdd
    End Function
#End Region

#Region "monthly attendance"
    Public Sub monthatt()
        Dim ac As Integer = record_count("select * from academic_year where status= '1' and sid = '" & sid & "'")
        If ac > 0 Then
            mm = Now.Date.Month
            'If mm = 0 Then
            '    mm = 12
            'End If
            Try
                'Dim first As Integer
                'first = Now.Date.Day
                'If first = 1 Or first = 5 Or first = 8 Then
                Dim lm As Integer = 0
                lm = Now.Date.Month - 1
                Dim tb6 As New DataTable
                tb6 = ShowData("select * from staff where sid = '" & sid & "'")
                For i As Integer = 0 To tb6.Rows.Count - 1
                    Dim num As Integer = record_count("select * from payroll where Staff_ID = '" & tb6.Rows(i)("staff_id").ToString & "' and Academic='" & academic_year & "' and Month= '" & mm & "' and sid = '" & sid & "'")
                    If num < 1 Then
                        Dim salary As Double = getnumber("select salary from salary where staff_id = '" & tb6.Rows(i)("staff_id").ToString & "' and sid = '" & sid & "'")
                        execute("insert into payroll(Staff_ID,Name,initial_Salary,Academic,Month,sid) values ('" + tb6.Rows(i)("staff_id").ToString + "','" + tb6.Rows(i)("fullname").ToString + "','" & salary & "','" & academic_year & "','" & mm & "','" & sid & "') ", 3)
                    End If
                Next
                '  present()
                '   End If
            Catch ex As Exception
                showerror(ex.Message)
            End Try
        End If
    End Sub

#End Region

#Region "get saturday and sundays"
    Public Sub getwkends()
        Dim month As Integer = Now.Date.Month
        Dim year As Integer = Now.Date.Year
        Dim current As New DateTime(year, month, 1)
        Dim ending As New DateTime(year, month, daym)

        Dim numSat As Integer = 0
        Dim numSun As Integer = 0

        While current < ending
            If current.DayOfWeek = DayOfWeek.Saturday Then
                numSat += 1
            End If
            If current.DayOfWeek = DayOfWeek.Sunday Then
                numSun += 1
            End If
            current = current.AddDays(1)
        End While
        sats = numSat
        sundays = numSun
    End Sub
#End Region

#Region "get saturdays only"
    Public Function saturdays(ByVal current As DateTime, ByVal ending As DateTime)
        Dim numSat As Integer = 0
        ending = ending.AddDays(1)
        While current < ending
            If current.DayOfWeek = DayOfWeek.Saturday Then
                numSat += 1
            End If
            current = current.AddDays(1)
        End While
        Return numSat
    End Function
#End Region

#Region "check for sunday only"
    Public Function checkSunday(ByVal current As DateTime)
        Dim sun As Boolean = False
        If current.DayOfWeek = DayOfWeek.Sunday Then
            sun = True
        End If
        Return sun
    End Function
#End Region

#Region "final todate"
    Public Function mchck(ByVal ddy As Integer, ByVal mm As Integer)
        Dim tdd As Date
        Dim y As Integer = Now.Date.Year
        tdd = y & "-" & mm & "-" & ddy
        Return tdd
    End Function
#End Region

#Region "total days for a month"
    Public Function deys(ByVal m As Integer)
        Dim dy As Integer
        m = Now.Date.Month - 1
        If m = 0 Then
            dy = 31
        ElseIf m = 12 Then
            dy = 31
        ElseIf m = 3 Then
            dy = 31
        ElseIf m = 5 Then
            dy = 31
        ElseIf m = 7 Then
            dy = 31
        ElseIf m = 8 Then
            dy = 31
        ElseIf m = 10 Then
            dy = 31
        ElseIf m = 6 Then
            dy = 30
        ElseIf m = 4 Then
            dy = 30
        ElseIf m = 9 Then
            dy = 30
        ElseIf m = 11 Then
            dy = 30
        ElseIf m = 2 Then
            dy = 28
        End If
        Return dy
    End Function
#End Region

#Region "get hours"
    Public Function TotalSalesHours()
        getwkends()
        Dim thours As Integer
        Dim weekhours, satsstore As Integer
        working_daya = daym - (sats + sundays)
        mm = Now.Date.Month
        If mm = 12 Then
            working_daya = working_daya - 3
        ElseIf mm = 1 Then
            working_daya = working_daya - 2
        ElseIf mm = 3 Then
            working_daya = working_daya - 1
        ElseIf mm = 4 Then
            working_daya = working_daya - 2
        ElseIf mm = 5 Then
            working_daya = working_daya - 3
        ElseIf mm = 7 Then
            working_daya = working_daya - 2
        ElseIf mm = 8 Then
            working_daya = working_daya - 1
        ElseIf mm = 9 Then
            working_daya = working_daya - 1
        End If
        weekhours = working_daya * 9
        working_daya = working_daya + sats
        satsstore = sats * 8
        thours = weekhours + satsstore
        Return thours
    End Function
    Public Function TotalFactoryHours()
        getwkends()
        Dim thours As Integer
        Dim working_daya As Integer
        Dim weekhours, satsfactory As Integer
        working_daya = daym - (sats + sundays)
        mm = Now.Date.Month
        If mm = 12 Then
            working_daya = working_daya - 3
        ElseIf mm = 1 Then
            working_daya = working_daya - 2
        ElseIf mm = 3 Then
            working_daya = working_daya - 1
        ElseIf mm = 4 Then
            working_daya = working_daya - 2
        ElseIf mm = 5 Then
            working_daya = working_daya - 3
        ElseIf mm = 7 Then
            working_daya = working_daya - 2
        ElseIf mm = 8 Then
            working_daya = working_daya - 1
        ElseIf mm = 9 Then
            working_daya = working_daya - 1
        End If
        weekhours = working_daya * 9
        satsfactory = sats * 5
        working_daya = working_daya + sats
        thours = weekhours + satsfactory
        Return thours
    End Function
#End Region

#Region "insert into working days"
    Public Sub input_working_data()
        Try
            Dim wc, factory, sales As Integer
            wc = record_count("select * from working_days where Month = '" & Now.Date.Month & "' and Year = '" & Now.Date.Year & "'")
            If wc < 1 Then
                factory = TotalFactoryHours()
                sales = TotalSalesHours()
                execute("INSERT INTO `working_days`( `Days`, `Total_Factory_Hours`, `Total_Sales_Hours`, `Month`, `Year`) VALUES ('" & working_daya & "','" & factory & "','" & sales & "','" & Now.Date.Month & "','" & Now.Date.Year & "')", 3)
            End If
        Catch ex As Exception
            showerror(ex.Message)
        End Try
    End Sub
#End Region



End Module

