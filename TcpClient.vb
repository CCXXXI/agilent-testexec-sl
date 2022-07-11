Imports System.Linq
Imports System.Net.Sockets

Public Class TcpClient
    Const Host As String = "localhost"
    Const Port As Integer = 8000
    ReadOnly _readCmd As Byte() = New Byte() {&HFE, &H2, &H0, &H0, &H0, &H8, &H6D, &HC3}
    Const YesCode As Byte = &H1
    ReadOnly _report As Byte() = New Byte() {&HFE, &H5, &H0, &H0, &HFF, &H0, &H98, &H35}
    ReadOnly _expectResponse As Byte() = New Byte() {&HFE, &H5, &H0, &H0, &HFF, &H0, &H98, &H35}

    ReadOnly _tcpClient As New Net.Sockets.TcpClient()
    Dim _networkStream As NetworkStream

    Public Function Connect() As Boolean
        ' Try to connect to the server
        ' Return true if connected, false if not
        Try
            _tcpClient.Connect(Host, Port)
        Catch ex As SocketException
            Return False
        End Try
        _networkStream = _tcpClient.GetStream()
        Return True
    End Function

    Public Function Check() As Boolean
        ' Check if the button is pressed

        ' send command
        _networkStream.Write(_readCmd, 0, _readCmd.Length)

        ' read result
        Dim bytes(_tcpClient.ReceiveBufferSize) As Byte
        _networkStream.Read(bytes, 0, CInt(_tcpClient.ReceiveBufferSize))

        ' check
        Return bytes(3) = YesCode
    End Function

    Public Function Report() As Boolean
        _networkStream.Write(_report, 0, _report.Length)

        Dim bytes(_expectResponse.Length) As Byte
        _networkStream.Read(bytes, 0, CInt(_expectResponse.Length))

        Return bytes.SequenceEqual(_expectResponse)
    End Function

    Public Function Connected() As Boolean
        Return _tcpClient.Connected
    End Function
End Class
