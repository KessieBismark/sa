Public Class Form1
    Dim spTime As DateTime
    Dim currentTime As TimeSpan
    Dim counter As Integer
    Dim h As Integer = 0
    Dim LstLog As New List(Of LogsInfo)
    Dim yest, noww As DateTime
    Dim r_sdate, r_edate As Date
    Dim tb As New DataTable
    Dim hours As Integer = 0
    Dim ptimer, t3, duration As Integer
    Dim tt, ssec, esec As TimeSpan
    Dim it, ot, log_date As DateTime
    Dim difhour, start_seconds, end_seconds, difout As Integer
    Dim maxhour As Integer = 9
    Dim dt, overtime As String
    Dim atime, itt As TimeSpan
    Dim tblist As New DataTable
    Dim c, ca, ov, Tot_rows, PROgres_count As Integer
    Dim adate As Date

    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        counter += 1
        Label1.Text = counter
        If counter = 50 Then
            Try
                spTime = GetDateTime()
                yest = spTime.AddDays(-1)
                noww = spTime
                Timer1.Start()
                Timer2.Stop()
            Catch ex As Exception
                counter = 0
            End Try
        End If
    End Sub
    Private Sub connectDevice(ByVal ip As String, ByVal port As String)
        Try
            Dim Device As New DeviceMain
            Timer3.Stop()
            Device.Connect(ip, port)

            Dim lastDate As Date
            Dim c = record_count("select * from attendance where '" & sid & "'")
            If (c > 0) Then
                lastDate = getValue("select max(`date`) from attendance where '" & sid & "'")
                LstLog = Device.DownloadLogs(lastDate, Now)
                LstLog.GroupBy(Function(s) s.Staff_ID)
                dgrid.DataSource = LstLog
                import_attendance()
            Else
                LstLog = Device.DownloadLogs(yest, Now)
                LstLog.GroupBy(Function(s) s.Staff_ID)
                dgrid.DataSource = LstLog
                import_attendance()
            End If
        Catch ex As Exception
            Timer3.Start()
        End Try
    End Sub
    Private Sub Request(ByVal ip As String, ByVal port As String)
        Try
            Timer4.Stop()
            ptimer = 0
            Dim Device As New DeviceMain
            Device.Connect(ip, port)

            LstLog = Device.DownloadLogs(r_sdate, r_edate)
            LstLog.GroupBy(Function(s) s.Staff_ID)
            dgrid.DataSource = LstLog
            import_attendance()
            execute("delete from request where start_date = '" & r_sdate.ToString("yyyy-MM-dd") & "' and end_date ='" & r_edate.ToString("yyyy-MM-dd") & "' and shop = '" & sid & "' and status = 1 ", 3)
        Catch ex As Exception
            Timer4.Start()
        End Try
    End Sub
    Private Sub Timer3_Tick(sender As Object, e As EventArgs) Handles Timer3.Tick
        t3 += 1
        If (t3 = duration) Then
            If h = 8 Or h = 9 Or h = 10 Then
                Label1.Text = h
                connectDevice("192.168.1.201", "4370")
            End If
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim Device As New DeviceMain
        Device.Connect("192.168.1.201", "4370")

        LstLog = Device.DownloadLogs("2021-12-17", "2021-12-17")
        LstLog.GroupBy(Function(s) s.Staff_ID)
        dgrid.DataSource = LstLog
    End Sub
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        ptimer = ptimer + 1
        If (ptimer = 1000000) Then
            get_requests()
            ptimer = 0
        End If
    End Sub

    Sub load_sid()
        Try
            Dim file As System.IO.StreamReader
            file = My.Computer.FileSystem.OpenTextFileReader("sid.txt")
            sid = file.ReadLine
            file.Close()
        Catch ex As Exception
        End Try
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.sdk = False Then
            Try
                Process.Start(Application.StartupPath & "\TheSDK\Register_SDK.bat")
                My.Settings.sdk = True
            Catch ex As Exception
                showerror(ex.Message)
            End Try
        End If
        Try
            load_sid()
            duration = getnumber("select duration from refresh")
            Timer2.Start()
            Timer3.Start()
            Timer4.Start()
        Catch ex As Exception
        End Try
    End Sub
    Sub get_requests()
        Try
            tb = New DataTable
            tb = ShowData("select * from request where shop = '" & sid & "' and  status = 1")
            If tb.Rows.Count > 0 Then
                r_sdate = tb.Rows(0)("start_date").ToString
                r_edate = tb.Rows(0)("end_date").ToString
                Request("192.168.1.201", "4370")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
#Region "get internet date and time"
    Public Shared Function GetDateTime() As DateTime
        Dim cd As Date
        Dim request As System.Net.HttpWebRequest = DirectCast(System.Net.WebRequest.Create("http://www.microsoft.com"), System.Net.HttpWebRequest)
        request.Method = "GET"
        request.Accept = "text/html, application/xhtml+xml, */*"
        request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)"
        request.ContentType = "application/x-www-form-urlencoded"
        request.CachePolicy = New System.Net.Cache.RequestCachePolicy(System.Net.Cache.RequestCacheLevel.NoCacheNoStore)
        Dim response As System.Net.HttpWebResponse = DirectCast(request.GetResponse(), System.Net.HttpWebResponse)
        If response.StatusCode = System.Net.HttpStatusCode.OK Then
            cd = DateTime.Parse(response.Headers("date"))
        End If
        Return cd
    End Function
#End Region
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        spTime = spTime.AddSeconds(1)
        Label2.Text = spTime
        h = spTime.Hour
    End Sub
    Sub import_attendance()
        'save to database
        Dim itd As New DataTable
        Dim weekclosetime, weekendclosetime As TimeSpan
        Tot_rows = dgrid.Rows.Count
        For i As Integer = 0 To dgrid.Rows.Count - 1
            log_date = dgrid.Rows(i).Cells("Log").Value.ToString
            dt = log_date.ToString("yyyy-MM-dd HH:mm:ss")
            Dim ad As String() = dt.Split(" ")
            adate = ad(0).Trim
            tb3 = New DataTable
            tb3 = ShowData("select e.staff_id,e.name,e.shop as shop_id,s.weekdays,s.weekends from employee e inner join shops s on e.shop = s.id where e.staff_id =  '" & dgrid.Rows(i).Cells("Staff_ID").Value.ToString.Trim & "' and e.shop = '" & sid & "'")
            If tb3.Rows.Count > 0 Then
                itd = ShowData("select * from attendance where staff_id= '" & dgrid.Rows(i).Cells("staff_id").Value.ToString.Trim & "' and Date= '" & adate.ToString("yyyy-MM-dd") & "' and shop= '" & sid & "'")
                If (itd.Rows.Count < 1) Then
                    execute("insert into attendance(Staff_ID,In_Time,Date,shop) values ('" & tb3.Rows(0)("staff_id").ToString & "','" & ad(1) & "','" & adate.ToString("yyyy-MM-dd") & "','" & sid & "')", 3)
                Else
                    Dim h As String() = ad(1).Split(":")
                    difhour = Val(h(0))
                    If difhour > maxhour Then
                        maxhour = difhour
                        If itd.Rows.Count > 0 Then
                            itt = TimeSpan.Parse(itd.Rows(0)("in_time").ToString)
                        End If
                        Dim intime As String() = itd.Rows(0)("in_time").ToString.Split(":")
                        difhour = Val(intime(0))
                        ot = ad(1)
                        Dim ioout As String() = ad(1).Split(":")
                        difout = Val(ioout(0))
                        If difhour < difout Then
                            If difhour < maxhour Then
                                If difhour >= 8 Then
                                    it = itt.ToString
                                Else
                                    it = "8:00:00"
                                End If
                                tt = ot - it
                                If sat(itd.Rows(0)("date").ToString) = True Then
                                    ov = tt.TotalSeconds - weekendclosetime.TotalSeconds
                                Else
                                    ov = tt.TotalSeconds - weekclosetime.TotalSeconds
                                End If
                                Dim intSeconds As Int64 = ov
                                Dim timSpan As New TimeSpan(timSpan.TicksPerSecond * intSeconds)
                                Dim strTime As String = timSpan.Hours.ToString("00") & ":" & timSpan.Minutes.ToString("00") & ":" & timSpan.Seconds.ToString("00")
                                overtime = intSeconds & " ~ " & strTime
                                hours = tt.TotalHours
                                If hours > 9 Then
                                    hours = 9
                                End If
                                execute("update attendance set out_time= '" & ad(1) & "', hours='" & hours & "', overtime='" & overtime & "'  where  staff_id ='" & dgrid.Rows(i).Cells("staff_id").Value.ToString.Trim & "' and date = '" & adate.ToString("yyyy-MM-dd") & "'", 3)
                            Else
                                overtime = 0 & " ~ 00:00:00"
                                If difhour >= 8 Then
                                    it = itt.ToString
                                Else
                                    it = "8:00:00"
                                End If
                                tt = ot - it
                                hours = tt.TotalHours
                                execute("update attendance set out_time= '" & ad(1) & "', hours='" & hours & "', overtime='" & overtime & "'  where  staff_id ='" & dgrid.Rows(i).Cells("staff_id").Value.ToString.Trim & "' and date = '" & adate.ToString("yyyy-MM-dd") & "'", 3)
                            End If
                        End If
                    End If
                End If
            End If
            maxhour = 0
        Next
    End Sub
    Private Function sat(ByVal current As DateTime)
        Dim numSat As Boolean = False
        If current.DayOfWeek = DayOfWeek.Saturday Then
            numSat = True
        End If
        Return numSat
    End Function
End Class
