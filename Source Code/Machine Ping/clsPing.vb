'modified by Thomas Fischer 04/2004 (http://www.thomas-fischer.org) & ohters (see below)
Option Explicit On 

Imports System
Imports System.Net
Imports System.Net.Sockets
Imports System.Threading

'*
'* Parts of the code based on information from Visual Studio Magazine
'* more info : http://www.fawcette.com/vsm/2002_03/magazine/columns/qa/default.asp
'*

Public Class StateObj
    Public sck As Sockets.Socket
    Public Buffer(255) As Byte
    Public BufferSize As Integer
    Public from As EndPoint
    Public BytesReceived As Integer
    Public TimeOut As Boolean = False
    Public rectime As Double

End Class

Public Class clsPing

    Public Structure stcError
        Dim Number As Integer
        Dim Description As String
    End Structure

    Private Const PING_ERROR_BASE As Long = &H8000

    Public Const PING_SUCCESS As Long = 0
    Public Const PING_ERROR As Long = (-1)
    Public Const PING_ERROR_HOST_NOT_FOUND As Long = PING_ERROR_BASE + 1
    Public Const PING_ERROR_SOCKET_DIDNT_SEND As Long = PING_ERROR_BASE + 2
    Public Const PING_ERROR_HOST_NOT_RESPONDING As Long = PING_ERROR_BASE + 3
    Public Const PING_ERROR_TIME_OUT As Long = PING_ERROR_BASE + 4

    Private Const ICMP_ECHO As Integer = 8
    Private Const SOCKET_ERROR As Integer = -1

    Private udtError As stcError

    Private Const intPortICMP As Integer = 7
    Private Const intBufferHeaderSize As Integer = 8
    Private Const intPackageHeaderSize As Integer = 28

    Private byteDataSize As Byte
    Private lngTimeOut As Integer
    Private ipheLocalHost As System.Net.IPHostEntry

    Protected alldone As New ManualResetEvent(False)

    'Public IP As String
    Private IPad As IPAddress
    Private ID As Short
    Private epServer As System.Net.EndPoint
    Private sequence As Byte = 0


    '*
    '* Class Constructor
    '*
    Public Sub New(Optional ByVal Identifier As Short = 0)

        ID = Identifier
        udtError = New stcError()

        '*
        '* Get local IP and transform in EndPoint
        '*
        ipheLocalHost = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName())

    End Sub

    ' Ping using asynchronous read.
    'Author: Lionel BERTON
    'Source: clsPing class from Paulo dos Santos Silva Jr
    'Date: 18th may 2003

    Public Function Ping(ByVal IP As String, Optional ByVal DataSize As Byte = 32, Optional ByVal Timeout As Integer = 1000) As Double

        Dim aReplyBuffer(255) As Byte
        Dim intStart As Double

        Dim epFrom As System.Net.EndPoint

        Dim ipepServer As System.Net.IPEndPoint


        byteDataSize = DataSize
        lngTimeOut = Timeout
        IPad = IPAddress.Parse(IP)
        '*
        '* Transforms the IP address in EndPoint
        '*
        ipepServer = New System.Net.IPEndPoint(IPad, 0)
        epServer = CType(ipepServer, System.Net.EndPoint)

        epFrom = New System.Net.IPEndPoint(ipheLocalHost.AddressList(0), 0)

        '*
        '* Builds the packet to send
        '*
        DataSize = Convert.ToByte(DataSize + intBufferHeaderSize)

        '*
        '* The packet must by an even number, so if the DataSize is and odd number adds one 
        '* 
        If (DataSize Mod 2 = 1) Then
            DataSize += Convert.ToByte(1)
        End If
        Dim aRequestBuffer(DataSize - 1) As Byte

        '*
        '* --- ICMP Echo Header Format ---
        '* (first 8 bytes of the data buffer)
        '*
        '* Buffer (0) ICMP Type Field
        '* Buffer (1) ICMP Code Field
        '*     (must be 0 for Echo and Echo Reply)
        '* Buffer (2) checksum hi
        '*     (must be 0 before checksum calc)
        '* Buffer (3) checksum lo
        '*     (must be 0 before checksum calc)
        '* Buffer (4) ID hi
        '* Buffer (5) ID lo
        '* Buffer (6) sequence hi
        '* Buffer (7) sequence lo
        '* Buffer (8)..(n)  Ping Data
        '*

        '*
        '* Set Type Field
        '*
        aRequestBuffer(0) = Convert.ToByte(8) ' ECHO Request

        '*
        '* Set ID field
        '*
        BitConverter.GetBytes(ID).CopyTo(aRequestBuffer, 4)

        '*
        '* Set Sequence
        '*
        BitConverter.GetBytes(sequence).CopyTo(aRequestBuffer, 6)

        '*
        '* Load some data into buffer
        '*
        Dim i As Integer
        For i = 8 To DataSize - 1
            aRequestBuffer(i) = Convert.ToByte(i Mod 8)
        Next i

        '*
        '* Calculate Checksum
        '*
        Call CreateChecksum(aRequestBuffer, DataSize, aRequestBuffer(2), aRequestBuffer(3))


        '*
        '* Create the socket
        '*
        Dim sckSocket As New System.Net.Sockets.Socket( _
                                        Net.Sockets.AddressFamily.InterNetwork, _
                                        Net.Sockets.SocketType.Raw, _
                                        Net.Sockets.ProtocolType.Icmp)
        sckSocket.Blocking = False

        '*
        '* Sends Package
        '*
        alldone.Reset()
        Dim so As New StateObj()
        so.sck = sckSocket
        'socket returns a packet with the IP header (20 bytes) so with 8 bytes 
        'header for ICMP and 32 bytes of data we should get 60 bytes but sometimes
        ' it is more or less... so let's take it large
        '
        so.BufferSize = DataSize + 30

        so.from = epServer
        Dim ar As IAsyncResult


        sckSocket.SendTo(aRequestBuffer, 0, DataSize, SocketFlags.None, ipepServer)

        '*
        '* Records the Start Time, after sending the Echo Request
        '*
        intStart = DateTime.Now.Ticks  'System.Environment.TickCount


        ar = sckSocket.BeginReceiveFrom(so.Buffer, 0, so.BufferSize, SocketFlags.None, so.from, New AsyncCallback(AddressOf callbackfunc), so)

        Dim ret As Boolean = False

        ret = alldone.WaitOne(lngTimeOut, False) 'return false if timeout
        alldone.Close()

        If ret = False Then
            so.TimeOut = True
            sckSocket.Close()
            sckSocket = Nothing
            udtError.Number = PING_ERROR_TIME_OUT
            udtError.Description = "Time Out"
            'Console.WriteLine("timeout for")
            Return (PING_ERROR)
        End If



        '*
        '* Informs on GetLastError the state of the server
        '*
        '
        udtError.Number = BitConverter.ToInt16(so.Buffer, 19)
        Select Case so.Buffer(20)
            Case 0 : udtError.Description = "Success"
            Case 1 : udtError.Description = "Buffer too Small"
            Case 2 : udtError.Description = "Destination Unreahable"
            Case 3 : udtError.Description = "Dest Host Not Reachable"
            Case 4 : udtError.Description = "Dest Protocol Not Reachable"
            Case 5 : udtError.Description = "Dest Port Not Reachable"
            Case 6 : udtError.Description = "No Resources Available"
            Case 7 : udtError.Description = "Bad Option"
            Case 8 : udtError.Description = "Hardware Error"
            Case 9 : udtError.Description = "Packet too Big"
            Case 10 : udtError.Description = "Reqested Timed Out"
            Case 11 : udtError.Description = "Bad Request"
            Case 12 : udtError.Description = "Bad Route"
            Case 13 : udtError.Description = "TTL Exprd In Transit"
            Case 14 : udtError.Description = "TTL Exprd Reassemb"
            Case 15 : udtError.Description = "Parameter Problem"
            Case 16 : udtError.Description = "Source Quench"
            Case 17 : udtError.Description = "Option too Big"
            Case 18 : udtError.Description = "Bad Destination"
            Case 19 : udtError.Description = "Address Deleted"
            Case 20 : udtError.Description = "Spec MTU Change"
            Case 21 : udtError.Description = "MTU Change"
            Case 22 : udtError.Description = "Unload"
            Case Else : udtError.Description = "General Failure"
        End Select


        Return (so.rectime - intStart)

    End Function

    Private Sub callbackfunc(ByVal ar As IAsyncResult)
        Try
            Dim rectime As Double = DateTime.Now.Ticks '= System.Environment.TickCount
            Dim so As StateObj = CType(ar.AsyncState, StateObj)

            If so.TimeOut = False Then
                so.BytesReceived = so.sck.EndReceiveFrom(ar, so.from)
                If so.BytesReceived > 0 Then
                    If ID = BitConverter.ToInt16(so.Buffer, 24) Then
                        so.rectime = rectime
                        alldone.Set()
                    Else
                        ar = so.sck.BeginReceiveFrom(so.Buffer, 0, so.BufferSize, SocketFlags.None, epServer, New AsyncCallback(AddressOf callbackfunc), so)

                    End If

                End If
            End If
        Catch

        End Try

    End Sub

    Public Function GetLastError() As stcError
        Return udtError
    End Function

    ' ICMP requires a checksum that is the one's
    ' complement of the one's complement sum of
    ' all the 16-bit values in the data in the
    ' buffer.
    ' Use this procedure to load the Checksum
    ' field of the buffer.
    ' The Checksum Field (hi and lo bytes) must be
    ' zero before calling this procedure.
    Private Sub CreateChecksum(ByRef data() As Byte, ByVal Size As Integer, ByRef HiByte As Byte, ByRef LoByte As Byte)
        Dim i As Integer
        Dim chk As Integer = 0

        For i = 0 To Size - 1 Step 2
            chk += Convert.ToInt32(data(i) * &H100 + data(i + 1))
        Next

        chk = Convert.ToInt32((chk And &HFFFF&) + Fix(chk / &H10000&))
        chk += Convert.ToInt32(Fix(chk / &H10000&))
        chk = Not (chk)

        HiByte = Convert.ToByte((chk And &HFF00&) / &H100)
        LoByte = Convert.ToByte(chk And &HFF)
    End Sub

End Class