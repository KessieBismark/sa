Imports zkemkeeper.CZKEMClass
Public Class DeviceMain
    Public ZKTime As New zkemkeeper.CZKEM

    Dim LstLogsData As New List(Of LogsInfo)

    Dim iMachineNumber As Integer
    Dim idwErrorCode As Integer
    Dim idwEnrollNumber As Integer = 0
    Dim idwVerifyMode As Integer = 0
    Dim idwInOutMode As Integer = 0
    Dim idwYear As Integer = 0
    Dim idwMonth As Integer = 0
    Dim idwDay As Integer = 0
    Dim idwHour As Integer = 0
    Dim idwMinute As Integer = 0
    Dim idwSecond As Integer = 0
    Dim idwWorkCode As Integer = 0
    Dim idwReserved As Integer = 0

    Dim IsConnected = False

    Sub Connect(IP As String, Port As String)

        If IP.Trim() = "" Or Port.Trim() = "" Then
            showerror("IP and Port cannot be null")
            Return
        End If

        Dim idwErrorCode As Integer

        ZKTime.Disconnect()
        IsConnected = False

        IsConnected = ZKTime.Connect_Net(IP, Convert.ToInt32(Port))

        If IsConnected = True Then
            iMachineNumber = 1
            ZKTime.RegEvent(iMachineNumber, 65535) 'Here you can register the realtime events that you want to be triggered(the parameters 65535 means registering all)
        Else
            ZKTime.GetLastError(idwErrorCode)
            showerror("Unable to connect the device,ErrorCode=" & idwErrorCode)
        End If

    End Sub

    Function DownloadLogs(Fromdate As Date, Todate As Date) As List(Of LogsInfo)

        If IsConnected = False Then
            showerror("Please connect the device first")
            Return LstLogsData
        End If

        ZKTime.EnableDevice(iMachineNumber, False) 'disable the device

        If ZKTime.ReadAllGLogData(iMachineNumber) Then 'read the records to the memory

            LstLogsData.Clear()

            While ZKTime.SSR_GetGeneralLogData(iMachineNumber, idwEnrollNumber, idwVerifyMode, idwInOutMode, idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond, idwWorkCode)

                '------------------
                Dim DateNow As Date = (Me.idwYear) & "/" & (Me.idwMonth) & "/" & (Me.idwDay) & " " & (Me.idwHour) & ":" & (Me.idwMinute) & ":" & (Me.idwSecond)

                Dim oDate As System.DateTime = Convert.ToDateTime(DateNow)
                Dim [date] As System.DateTime = oDate.[Date]
                Dim value As System.DateTime = Fromdate
                Dim dateTime As System.DateTime = oDate.[Date]
                Dim value1 As System.DateTime = Todate
                Dim date1 As System.DateTime = oDate.[Date]
                Dim dateTime1 As System.DateTime = Fromdate
                Dim date2 As System.DateTime = oDate.[Date]
                Dim value2 As System.DateTime = Todate

                If (Not (System.DateTime.Compare([date], value.[Date]) > 0 And System.DateTime.Compare(dateTime, value1.[Date]) < 0 Or System.DateTime.Compare(date1, dateTime1.[Date]) = 0 Or System.DateTime.Compare(date2, value2.[Date]) = 0)) Then
                    Continue While
                End If

                '------------------

                Dim FingerPrintDate As String = New DateTime(idwYear, idwMonth, idwDay, idwHour, idwMinute, idwSecond).ToString()

                Dim OInfo As LogsInfo = New LogsInfo()

                ' OInfo.MachineNumber = iMachineNumber
                OInfo.Staff_ID = Integer.Parse(idwEnrollNumber)
                OInfo.Log = FingerPrintDate
                LstLogsData.Add(OInfo)

            End While

        Else

            ZKTime.GetLastError(idwErrorCode)

            If idwErrorCode <> 0 Then
                showerror("Reading data from terminal failed,ErrorCode: " & idwErrorCode)
            Else
                showerror("No data from terminal returns!")
            End If

        End If

        ZKTime.Disconnect()

        ZKTime.EnableDevice(iMachineNumber, True) 'enable the device

        Return LstLogsData

    End Function
End Class
