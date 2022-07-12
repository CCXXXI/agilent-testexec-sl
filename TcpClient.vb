Imports System.Linq
Imports System.Net.Sockets

Public Class TcpClient
    Const Host As String = "localhost"
    Const Port As Integer = 8000
    ReadOnly _readCmd As Byte() = New Byte() {&HFE, &H2, &H0, &H0, &H0, &H8, &H6D, &HC3}
    Const RunCode As Byte = &H1
    ReadOnly _reportCmd As Byte() = New Byte() {&HFE, &H5, &H0, &H0, &HFF, &H0, &H98, &H35}

    ReadOnly _tcpClient As New Net.Sockets.TcpClient()
    Dim _networkStream As NetworkStream

    Public Function Connect() As Boolean
        ' Try to connect to the server
        ' Return true if connected successfully, false if not
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
        Return Command(_readCmd)(3) = RunCode
    End Function


    Public Function Report() As Boolean
        Return Not _tcpClient.Connected Or Command(_reportCmd).Take(_reportCmd.Length).SequenceEqual(_reportCmd)
    End Function

    Private Function Command(cmd As Byte()) As Byte()
        ' Send a command to the server and return the response.

        ' send the command
        _networkStream.Write(cmd, 0, cmd.Length)

        ' read the response
        Dim response(_tcpClient.ReceiveBufferSize) As Byte
        _networkStream.Read(response, 0, CInt(_tcpClient.ReceiveBufferSize))

        Return response
    End Function
End Class
