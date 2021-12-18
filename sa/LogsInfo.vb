Public Class LogsInfo

    ' Public Property MachineNumber As Integer
    ' Dim datte As Date
    ' Dim timme As String
    Public Property Staff_ID As Integer
    Private _Name As DateTime

    'Public Property Name As String
    '    Get
    '        Name = getValue("select name from employee where Staff_ID= '" & Staff_ID & "' ")
    '    End Get
    '    Set(value As String)
    '        _Name = Name
    '    End Set
    'End Property

    Private _DateTimeRecord As DateTime
    Public Property Log As DateTime
        Set(value As DateTime)
            _DateTimeRecord = value

            'TimeOnly = value
            '  DateOnly = value
        End Set
        Get
            Return _DateTimeRecord.ToString("yyyy-MM-dd HH:mm:ss")
        End Get
    End Property

    'Private _DateOnly As DateTime
    'Public Property DateOnly As DateTime
    '    Get
    '        '   datte = _DateOnly.ToShortDateString
    '        Return _DateOnly.ToLongDateString
    '    End Get
    '    Set(value As DateTime)
    '        _DateOnly = value
    '    End Set
    'End Property

    'Private _TimeOnly As DateTime
    'Public Property TimeOnly As DateTime
    '    Get
    '        ' timme = _TimeOnly.ToString("hh:mm:ss tt")
    '        Return _TimeOnly.ToShortTimeString
    '    End Get
    '    Set(value As DateTime)
    '        _TimeOnly = value
    '    End Set
    'End Property



End Class
