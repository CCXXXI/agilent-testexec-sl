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
        If _networkStream.CanWrite And _networkStream.CanRead Then
            ' Do a simple write.
            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes("test")
            _networkStream.Write(sendBytes, 0, sendBytes.Length)
            ' Read the NetworkStream into a byte buffer.
            Dim bytes(_tcpClient.ReceiveBufferSize) As Byte
            _networkStream.Read(bytes, 0, CInt(_tcpClient.ReceiveBufferSize))
            ' todo: check
            Dim response As String = Encoding.ASCII.GetString(bytes)
            Return String.Compare(response, "test") = 0
        Else
            If Not _networkStream.CanRead Then
                Console.WriteLine("cannot not write data to this stream")
                _tcpClient.Close()
            Else
                If Not _networkStream.CanWrite Then
                    Console.WriteLine("cannot read data from this stream")
                    _tcpClient.Close()
                End If
            End If
        End If
    End Function

    Public Sub Close()
        _tcpClient.Close()
    End Sub
End Class