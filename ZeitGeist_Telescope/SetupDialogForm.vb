Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports ASCOM.Utilities
Imports ASCOM.ZeitGeist.eq5

<ComVisible(False)> _
Public Class SetupDialogForm

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click ' OK button event handler
        ' Persist new values of user settings to the ASCOM profile
        Telescope.comPort = ComboBoxComPort.SelectedItem ' Update the state variables with results from the dialogue
        Telescope.serialSpeed = ComboBoxSerialSpeed.SelectedItem ' Update the state variables with results from the dialogue
        Telescope.traceState = chkTrace.Checked
        Telescope.followState = chkFollow.Checked
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click 'Cancel button event handler
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ShowAscomWebPage(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.DoubleClick, PictureBox1.Click
        ' Click on ASCOM logo event handler
        Try
            System.Diagnostics.Process.Start("http://ascom-standards.org/")
        Catch noBrowser As System.ComponentModel.Win32Exception
            If noBrowser.ErrorCode = -2147467259 Then
                MessageBox.Show(noBrowser.Message)
            End If
        Catch other As System.Exception
            MessageBox.Show(other.Message)
        End Try
    End Sub

    Private Sub SetupDialogForm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load ' Form load event handler
        ' Retrieve current values of user settings from the ASCOM Profile
        InitUI()
    End Sub

    Private Sub InitUI()
        chkTrace.Checked = Telescope.traceState
        chkFollow.Checked = Telescope.followState
        ' set the list of com ports to those that are currently available
        ComboBoxComPort.Items.Clear()
        ComboBoxComPort.Items.AddRange(System.IO.Ports.SerialPort.GetPortNames())       ' use System.IO because it's static
        ' select the current port if possible
        If ComboBoxComPort.Items.Contains(Telescope.comPort) Then
            ComboBoxComPort.SelectedItem = Telescope.comPort
        End If

        ComboBoxSerialSpeed.Items.Clear()
        Dim serialSpeeds = New String() {
            ASCOM.Utilities.SerialSpeed.ps300,
            ASCOM.Utilities.SerialSpeed.ps1200,
            ASCOM.Utilities.SerialSpeed.ps2400,
            ASCOM.Utilities.SerialSpeed.ps4800,
            ASCOM.Utilities.SerialSpeed.ps9600,
            ASCOM.Utilities.SerialSpeed.ps14400,
            ASCOM.Utilities.SerialSpeed.ps19200,
            ASCOM.Utilities.SerialSpeed.ps28800,
            ASCOM.Utilities.SerialSpeed.ps38400,
            ASCOM.Utilities.SerialSpeed.ps57600,
            ASCOM.Utilities.SerialSpeed.ps115200,
            ASCOM.Utilities.SerialSpeed.ps230400
        }
        ComboBoxSerialSpeed.Items.AddRange(serialSpeeds)       ' use System.IO because it's static
        ' select the current port if possible
        If ComboBoxSerialSpeed.Items.Contains(Telescope.serialSpeed) Then
            ComboBoxSerialSpeed.SelectedItem = Telescope.serialSpeed
        End If
    End Sub

End Class
