Imports System
Imports System.IO.Ports
Imports System.Threading
Imports System.Diagnostics
Imports MySql.Data.MySqlClient

Public Class frmMain
    Private serialPort As New IO.Ports.SerialPort

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click

        Dim DataArray(100) As Integer
        Dim Counter As Integer
        Dim DCP, Parameter, Value As Integer
        Dim conn As New MySqlConnection
        Dim SQLCmd As New MySqlCommand
        Dim I, r As Integer

        Dim PrevRXbyte, RXByte As Byte
        Dim CBArray(255) As Byte 'Command Block
        Dim FullData As String = ""
        Dim DPC As Integer 'Data Part Counter
        Dim Restart As Boolean

        Dim DatabaseName As String = "Nibe"
        Dim server As String = "192.168.192.18"
        Dim userName As String = "nibe"
        Dim password As String = "nibe"
        If Not conn Is Nothing Then conn.Close()
        conn.ConnectionString = String.Format("server={0}; user id={1}; password={2}; database={3}; pooling=false", server, userName, password, DatabaseName)

        cmdStart.Enabled = False

        Me.Refresh()

Start:
        Restart = False

        System.Windows.Forms.Application.DoEvents()

        Try

            If serialPort.IsOpen Then
                serialPort.Close()
            End If

            With serialPort
                .PortName = "COM5"
                .BaudRate = 19200
                .Parity = IO.Ports.Parity.Mark
                .DataBits = 8
                .StopBits = IO.Ports.StopBits.One
                .ParityReplace = Nothing
                .RtsEnable = False
                .ReadTimeout = 2000
            End With

            serialPort.Open()
            serialPort.DiscardInBuffer()
            System.Windows.Forms.Application.DoEvents()


            lblMessage.Text = "Port connected."

            Do

                System.Windows.Forms.Application.DoEvents()

                RXByte = serialPort.ReadByte 'Read on byte from serial

                If PrevRXbyte.ToString("X2") = "00" And RXByte.ToString("X2") = "14" Then ' Check for leading 0014
                    With serialPort 'sending "ACK"
                        .Parity = Parity.Space
                        .RtsEnable = True
                        .Write(Chr(6))
                        .RtsEnable = False
                        .Parity = Parity.Mark
                    End With
                    lblMessage.Text = "Find Message"

                    RXByte = serialPort.ReadByte
                    If RXByte.ToString("X2") = "C0" Then
                        CBArray(0) = RXByte 'Command Block Start
                    Else
                        lblMessage.Text = "Missade Command Block C0"
                        Restart = True
                        Exit Do
                    End If

                    RXByte = serialPort.ReadByte
                    If RXByte.ToString("X2") = "00" Then
                        CBArray(1) = RXByte

                    Else
                        lblMessage.Text = "Missade Command Block 00"
                        Restart = True
                        Exit Do
                    End If

                    RXByte = serialPort.ReadByte
                    If RXByte.ToString("X2") = "7C" Then
                        CBArray(2) = RXByte

                    Else
                        lblMessage.Text = "Missade Sender"
                        Restart = True
                        Exit Do
                    End If

                    RXByte = serialPort.ReadByte
                    CBArray(3) = RXByte

                    For DPC = 1 To Val("&h" & CBArray(3).ToString("X2"))
                        System.Windows.Forms.Application.DoEvents()
                        RXByte = serialPort.ReadByte
                        CBArray(DPC + 3) = RXByte
                    Next

                    RXByte = serialPort.ReadByte
                    CBArray(DPC + 3) = RXByte

                    With serialPort
                        .Parity = Parity.Space
                        .RtsEnable = True
                        .Write(Chr(6))
                        .RtsEnable = False
                        .Parity = Parity.Mark
                    End With

                    RXByte = serialPort.ReadByte
                    If RXByte.ToString("X2") <> "03" Then
                        lblMessage.Text = "Missade ETX"
                        Restart = True
                        Exit Do
                    End If

                    For I = 0 To DPC + 3
                        FullData = FullData & CBArray(I).ToString("X2") & " "
                        System.Windows.Forms.Application.DoEvents()
                    Next

                    DCP = Val("&h" & CBArray(3).ToString("X2"))

                    For I = 5 To DCP + 4
                        Parameter = Val("&h" & CBArray(I).ToString("X2"))
                        Select Case Parameter
                            Case 0, 2, 3, 4, 5, 6, 7, 11, 12, 13, 14, 15, 17, 21, 22, 23, 24, 25, 32, 40, 41, 45, 47, 48, 49, 50, 51, 52, 53, 54, 56, 58, 59, 61, 62, 63, 64, 65, 66, 74, 75, 76, 77, 78, 79, 80, 85, 86, 87, 88, 89, 90, 91, 92, 96
                                Value = Val("&h" & CBArray(I + 1).ToString("X2"))
                                I = I + 2
                            Case 1, 9, 10, 16, 18, 19, 20, 26, 27, 28, 29, 30, 34, 35, 36, 37, 38, 39, 42, 43, 44, 55, 57, 60, 67, 68, 69, 70, 71, 72, 73, 81, 82, 83, 84, 93, 94, 95
                                Value = Val("&h" & CBArray(I + 1).ToString("X2") & CBArray(I + 2).ToString("X2"))
                                I = I + 3
                            Case 8, 31, 33, 46
                                Value = Val("&h" & CBArray(I + 1).ToString("X2") & CBArray(I + 2).ToString("X2") & CBArray(I + 3).ToString("X2") & CBArray(I + 4).ToString("X2"))
                                I = I + 5
                        End Select

                        DataArray(Parameter) = Value

                        'txtDataReceived.AppendText(Parameter & " " & Value & " " & Chr(13) & Chr(10))
                        Counter = Counter + 1
                        System.Windows.Forms.Application.DoEvents()
                    Next
                    lblMessage.Text = "Parsed message"


                    ReDim CBArray(255)
                    DPC = 0
                    PrevRXbyte = Nothing

                Else

                    PrevRXbyte = RXByte

                End If

                If Counter > 96 Then

                    lblMessage.Text = "Writing Data"

                    SQLCmd.Connection = conn
                    SQLCmd.CommandText = "Insert into Data Values ( "


                    For I = 0 To 95
                        System.Windows.Forms.Application.DoEvents()
                        SQLCmd.CommandText = SQLCmd.CommandText & DataArray(I) & ", "

                    Next
                    SQLCmd.CommandText = SQLCmd.CommandText & DataArray(96) & ")"

                    conn.Open()
                    r = SQLCmd.ExecuteNonQuery
                    conn.Close()

                    Counter = 0
                    ReDim DataArray(100)
                    serialPort.Close()

                    lblMessage.Text = "Sleeping"
                    For T As Integer = 1 To 60
                        System.Windows.Forms.Application.DoEvents()
                        System.Threading.Thread.Sleep(1000)
                    Next T

                    serialPort.Open()
                    serialPort.DiscardInBuffer()

                End If

            Loop

        Catch ex As Exception
            lblMessage.Text = ex.ToString
            Restart = True
        End Try


        ReDim CBArray(255)
            ReDim DataArray(100)
            DPC = 0
            PrevRXbyte = Nothing
            Counter = 0

            If Restart Then GoTo Start

Klar:
            lblMessage.AppendText(" Klar")

    End Sub
    Private Sub frmMain_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.FormClosing
        If serialPort.IsOpen Then
            serialPort.Close()
        End If
    End Sub
End Class

