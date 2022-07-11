Imports System.Net.Sockets

Public Class TcpClient
    Const Host As String = "localhost"
    Const Port As Integer = 8000
    ReadOnly _readCmd As Byte() = New Byte() {&HFE, &H2, &H0, &H0, &H0, &H8, &H6D, &HC3}
    Const YesCode As Byte = &H1

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
End Class
