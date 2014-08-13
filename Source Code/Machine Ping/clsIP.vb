'modified by Thomas Fischer 04/2004 (http://www.thomas-fischer.org)
Imports System.Net
Imports System.Net.Dns
Public Class clsIP
    Public Function GetLocalHostIP() As String

        Dim objAddress As IPAddress
        Dim sAns As String

        Try

            objAddress = New IPAddress( _
              GetHostByName(GetLocalHostName).AddressList(0).Address)
            sAns = objAddress.ToString


        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            sAns = ""
        End Try

        Return sAns

    End Function

    Public Function GetLocalHostName() As String
        Return GetHostName()
    End Function

    Public Function IPToHostName(ByVal IPAddress As String) _
          As String

        Dim objEntry As IPHostEntry
        Dim sAns As String
        Try
            objEntry = Resolve(IPAddress)
            sAns = objEntry.HostName
        Catch ex As Exception

            sAns = ""
        End Try

        Return sAns

    End Function

    Public Function HostNameToIP(ByVal Host As String) As String
        Dim objAddress As IPAddress
        Dim sAns As String

        Try
            objAddress = New IPAddress(GetHostByName(Host).AddressList(0).Address)
            sAns = objAddress.ToString
        Catch
            sAns = ""
        End Try
        Return sAns


    End Function
End Class
