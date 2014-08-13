Imports System.IO
Imports System.Web.Mail
Public Class Worker

    Inherits System.ComponentModel.Component

    ' Declares the variables you will use to hold your thread objects.

    Public WorkerThread As System.Threading.Thread

    Public targetnames As String = ""
    Public interval As String = ""
    Public email As String = ""

    Public result As String = ""

    Public Event WorkerComplete(ByVal Result As String)
    Public Event WorkerProgress(ByVal value As Integer)


#Region " Component Designer generated code "

    Public Sub New(ByVal Container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        Container.Add(Me)
    End Sub

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Component overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container
    End Sub

#End Region

    Private Sub Error_Handler(ByVal message As String)
        Try
            Dim Display_Message1 As New Display_Message(message)
            Display_Message1.ShowDialog()
        Catch ex As Exception
            MsgBox("An error occurred in Machine Ping's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Public Sub ChooseThreads(ByVal threadNumber As Integer)
        Try
            ' Determines which thread to start based on the value it receives.
            Select Case threadNumber
                Case 1
                    ' Sets the thread using the AddressOf the subroutine where
                    ' the thread will start.
                    WorkerThread = New System.Threading.Thread(AddressOf WorkerExecute)
                    ' Starts the thread.
                    WorkerThread.Start()

            End Select
        Catch ex As Exception
            Error_Handler(ex.Message)
        End Try
    End Sub

    Public Sub WorkerExecute()
        Try

            Dim targets As String()
            targets = targetnames.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To targets.Length - 1



                Dim error_encounter As Boolean = False
                Dim error_message As String = ""
                Try
                    Dim ProcID As Integer
                    Dim apppath As String = Application.StartupPath


                    If apppath.EndsWith("\") Then
                        apppath = apppath.Remove(apppath.Length - 1, 1)
                    End If

                    If System.IO.File.Exists(apppath & "\result.txt") = True Then
                        System.IO.File.Delete(apppath & "\result.txt")
                    End If
                    Dim runprog As String = """" & apppath & "\ping.exe"" """ & targets(counter).Trim & """ > """ & apppath & "\result.txt"""
                    DosShellCommand(runprog)

                    If System.IO.File.Exists(apppath & "\result.txt") = True Then
                        Dim reader As System.IO.StreamReader = New System.IO.StreamReader(apppath & "\result.txt", True)
                        Dim resultstring As String
                        resultstring = reader.ReadToEnd()
                        reader.Close()
                        If resultstring.LastIndexOf("could not find host") > 0 Then
                            error_encounter = True
                            error_message = resultstring
                        End If
                        If resultstring.LastIndexOf("Destination host unreachable.") > 0 Then
                            error_encounter = True
                            error_message = resultstring
                        End If
                        If resultstring.LastIndexOf("Request timed out.") > 0 Then
                            Dim teststr As String() = resultstring.Split("Request timed out.")
                            If teststr.Length > 5 Then
                                error_encounter = True
                                error_message = resultstring
                            End If
                        End If

                    Else
                        error_encounter = True
                    End If

                Catch ex As Exception
                    error_encounter = True
                End Try
                If error_encounter = True Then
                    RaiseEvent WorkerProgress(1)
                    
                    Try
                        Dim body As String
                        body = "An ICMP ping test has failed and yielded the following results below. This could imply that the machine is either down or that network traffic is being seriously hampered. Please dispatch or inform the necessary maintenance staff."
                        body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf
                        error_message = error_message.Replace(Chr(13), " ")
                        body = body & error_message.Trim
                        body = body & vbCrLf & vbCrLf & "******************************" & vbCrLf & vbCrLf & "This is an auto-generated email submitted from Machine Ping 1.0 at " & Format(Now(), "dd/MM/yyyy hh:mm:ss tt") & ", running on:"
                        body = body & vbCrLf & vbCrLf & "Machine Name: " + Environment.MachineName
                        body = body & vbCrLf & "OS Version: " & Environment.OSVersion.ToString()
                        body = body & vbCrLf & "User Name: " + Environment.UserName

                        TextMail("mail.uct.ac.za", "webserver@commerce.uct.ac.za", email, "Possible Machine Down on Network Detected", body)
                    Catch ex As Exception
                        Error_Handler("An """ & ex.Message & """ error occurred while sending a notification email. The program will attempt to recover shortly.")
                    End Try
                End If

            Next
            result = "Success"
            RaiseEvent WorkerComplete(result)
        Catch ex As Exception
            result = "Failure"
            RaiseEvent WorkerComplete(result)
        End Try

        WorkerThread.Abort()
    End Sub



    Private Function DosShellCommand(ByVal AppToRun As String) As String
        Dim s As String = ""
        Try
            Dim myProcess As Process = New Process

            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.CreateNoWindow = True
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.StartInfo.RedirectStandardError = True
            myProcess.Start()
            Dim sIn As StreamWriter = myProcess.StandardInput
            sIn.AutoFlush = True

            Dim sOut As StreamReader = myProcess.StandardOutput
            Dim sErr As StreamReader = myProcess.StandardError
            sIn.Write(AppToRun & _
               System.Environment.NewLine)
            sIn.Write("exit" & System.Environment.NewLine)
            s = sOut.ReadToEnd()
            If Not myProcess.HasExited Then
                myProcess.Kill()
            End If

            'MessageBox.Show("The 'dir' command window was closed at: " & myProcess.ExitTime & "." & System.Environment.NewLine & "Exit Code: " & myProcess.ExitCode)

            sIn.Close()
            sOut.Close()
            sErr.Close()
            myProcess.Close()
            'MessageBox.Show(s)
        Catch ex As Exception
            Error_Handler("An """ & ex.Message & """ error occurred while launching DOS shell. The program will attempt to recover shortly.")
        End Try
        Return s
    End Function

    Public Function TextMail(ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1
                objMail = New MailMessage
                objMail.BodyFormat = MailFormat.Text
                objMail.From = "webserver@commerce.uct.ac.za"
                objMail.To = emailaddys(counter).Trim
                objMail.Subject = strSubj
                objMail.Body = strBody

            SmtpMail.SmtpServer = "mail.uct.ac.za"
            SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

    Public Function TextMail(ByVal SmtpServer As String, ByVal strFrom As String, ByVal strTo As String, ByVal strSubj As String, ByVal strBody As String, Optional ByRef strErrMsg As String = "") As Boolean
        Dim objMail As MailMessage

        Try
            Dim emailaddys As String()
            emailaddys = strTo.Split(";")

            Dim counter As Integer = 0
            For counter = 0 To emailaddys.Length - 1


            objMail = New MailMessage
            objMail.BodyFormat = MailFormat.Text
            objMail.From = strFrom
                objMail.To = emailaddys(counter).Trim
            objMail.Subject = strSubj
            objMail.Body = strBody

            SmtpMail.SmtpServer = SmtpServer
                SmtpMail.Send(objMail)
            Next
            TextMail = True

        Catch ex As Exception
            TextMail = False
            Error_Handler("An """ & ex.Message & """ error occurred while sending the Error Alert email. The program will attempt to recover shortly.")
        End Try
    End Function

End Class
