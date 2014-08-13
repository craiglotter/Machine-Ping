'(c) Thomas Fischer 04/2004 (http://www.thomas-fischer.org)
Imports System.Net
Imports System.Net.Sockets
Public Enum ConnectionStates
    Down
    Up
    Unkown
End Enum

Public Structure ConnectionType
    Dim Hostname As String
    Dim Description As String
    Dim newState As ConnectionStates
    Dim oldState As ConnectionStates
    Dim Ping As Double
    Dim PingString As String
    Dim Sleeptime As Integer
    Dim Timeout As Integer
End Structure

Public Class clsConnectionTest
    Dim myData As ConnectionType
    Dim PingThread As New Threading.Thread(AddressOf pingit)
    Dim DNSResolveThread As New Threading.Thread(AddressOf ResolveThread)
    Dim myFile As System.IO.File

    Dim out As System.IO.StreamWriter
    Dim IP As String = "127.0.0.1"

    Public Sub New(ByVal Hostname As String, Optional ByVal Description As String = "", Optional ByVal Sleeptime As Integer = 1500, Optional ByVal TimeOut As Integer = 1200)
        myData = New ConnectionType()
        myData.Hostname = Hostname
        myData.Description = Description
        myData.Sleeptime = Sleeptime
        myData.Timeout = TimeOut
        generalnew()
    End Sub
    Public Sub New(ByVal ConnectionData As ConnectionType)
        myData = ConnectionData
        generalnew()
    End Sub

    Private Sub generalnew()
        PingThread.IsBackground = True
        DNSResolveThread.IsBackground = True

        out = myFile.CreateText(myData.Hostname & ".txt")
        out.AutoFlush = True

        DNSResolveThread.Start()
        PingThread.Start()
    End Sub

    Private Enum testphases
        NormalChecking
        RetryTest1
        RetryTest2
        RetryTest3
        DownTest
    End Enum

    Private Sub ResolveThread()
        Dim newIP As String
        Dim resolvesleeptime As Integer = 30000
        Do
            newIP = IPAddress.Parse(New clsIP().HostNameToIP(myData.Hostname)).ToString
            If newIP <> IP Then
                If IP = "" Then
                    out.WriteLine(DateTime.Now & " | Initially Resolved Hostname " & myData.Hostname & " to " & newIP & ", resolvecheck every " & (resolvesleeptime / 1000) & " s")
                Else
                    out.WriteLine(DateTime.Now & " | Resolved IP has changed from " & IP & " to " & newIP)
                End If
                IP = newIP
            End If
            DNSResolveThread.Sleep(resolvesleeptime)
        Loop
    End Sub

    Private Sub pingit()
        Dim myPing As clsPing
        Dim ID As String 'for display only!
        Dim testphase As testphases


        ID = myData.Hostname
        out.WriteLine(DateTime.Now & " | Thread started, testing " & myData.Hostname & " every " & myData.Sleeptime & " ms")

        myData.newState = ConnectionStates.Unkown
        myData.oldState = ConnectionStates.Unkown
        testphase = testphases.NormalChecking
        Do
            With myData
                'If testphase <> testphases.NormalChecking And testphase <> testphases.DownTest Then
                '    out.WriteLine(ID & " | " & DateTime.Now & " | CheckingState: " & testphase.ToString)
                'End If
                myPing = New clsPing()
                Select Case testphase
                    Case testphases.NormalChecking : myData.Timeout = 1200
                    Case testphases.RetryTest1 : myData.Timeout = 2000 : .Sleeptime = 3000
                    Case testphases.RetryTest2 : myData.Timeout = 5000 : .Sleeptime = 6000
                    Case testphases.RetryTest3 : myData.Timeout = 10000 : .Sleeptime = 11000
                    Case testphases.DownTest : myData.Timeout = 5000 : .Sleeptime = 6000
                End Select
                .Ping = myPing.Ping(IP, 32, myData.Timeout)

                If .Ping < 0 Then
                    .PingString = "down"
                    '.newState = ConnectionStates.Down  'nicht sofort setzen, erst 3 retries !!!

                    Select Case testphase
                        Case testphases.NormalChecking : testphase = testphases.RetryTest1
                        Case testphases.RetryTest1 : testphase = testphases.RetryTest2
                        Case testphases.RetryTest2 : testphase = testphases.RetryTest3
                        Case testphases.RetryTest3 : .newState = ConnectionStates.Down : testphase = testphases.DownTest
                    End Select
                ElseIf .Ping = 0 Then
                    .PingString = "local (<1 ms)"
                    .newState = ConnectionStates.Up
                Else

                    myData.Timeout = 1200
                    .Sleeptime = 1500
                    testphase = testphases.NormalChecking

                    .PingString = Math.Round((.Ping / 10000), 3).ToString & " ms"
                    .newState = ConnectionStates.Up
                End If
                'out.WriteLine(ID & " | " & DateTime.Now & " | " & .newState.ToString & " | Delay: " & .PingString)

                If .oldState <> .newState Then
                    out.WriteLine(DateTime.Now & " | " & .newState.ToString & IIf(.newState = ConnectionStates.Down, "", " | Delay: " & .PingString))
                End If

                Threading.Thread.Sleep(.Sleeptime)
                .oldState = .newState
            End With
        Loop
    End Sub

    Public Sub stopit()
        PingThread.Abort()
        DNSResolveThread.Abort()
        out.WriteLine(DateTime.Now & " | Thread stopped")
        out.Close()
    End Sub

End Class
