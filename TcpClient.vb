Imports System.Net.Sockets
Imports System.Text

Public Class TcpClient
    ReadOnly _tcpClient As New System.Net.Sockets.TcpClient()
    ReadOnly _networkStream As NetworkStream

    Public Sub New()
        _tcpClient.Connect("127.0.0.1", 8000)
        _networkStream = _tcpClient.GetStream()
    End Sub

    Public Function Check() As Boolean
        ' write
        Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes("test")
        _networkStream.Write(sendBytes, 0, sendBytes.Length)

        ' read
        Dim bytes(_tcpClient.ReceiveBufferSize) As Byte
        _networkStream.Read(bytes, 0, CInt(_tcpClient.ReceiveBufferSize))

        ' check
        Dim response As String = Encoding.ASCII.GetString(bytes)
        Return String.Compare(response, "test") = 0
    End Function
End Class
