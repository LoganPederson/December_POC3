<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.cboMicDevice = New System.Windows.Forms.ComboBox()
        Me.cboLoopbackDevice = New System.Windows.Forms.ComboBox()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.btnStop = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.txtTranscription = New System.Windows.Forms.RichTextBox()
        Me.txtSummary = New System.Windows.Forms.RichTextBox()
        Me.cboCombinedDevice = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cboMicDevice
        '
        Me.cboMicDevice.FormattingEnabled = True
        Me.cboMicDevice.Location = New System.Drawing.Point(41, 12)
        Me.cboMicDevice.Name = "cboMicDevice"
        Me.cboMicDevice.Size = New System.Drawing.Size(414, 21)
        Me.cboMicDevice.TabIndex = 0
        '
        'cboLoopbackDevice
        '
        Me.cboLoopbackDevice.FormattingEnabled = True
        Me.cboLoopbackDevice.Location = New System.Drawing.Point(41, 45)
        Me.cboLoopbackDevice.Name = "cboLoopbackDevice"
        Me.cboLoopbackDevice.Size = New System.Drawing.Size(414, 21)
        Me.cboLoopbackDevice.TabIndex = 1
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(125, 99)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(75, 23)
        Me.btnStart.TabIndex = 2
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.Location = New System.Drawing.Point(238, 99)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.btnStop.Size = New System.Drawing.Size(75, 23)
        Me.btnStop.TabIndex = 3
        Me.btnStop.Text = "Stop"
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnClear
        '
        Me.btnClear.Location = New System.Drawing.Point(351, 99)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(75, 23)
        Me.btnClear.TabIndex = 4
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = True
        '
        'txtTranscription
        '
        Me.txtTranscription.Location = New System.Drawing.Point(41, 139)
        Me.txtTranscription.Name = "txtTranscription"
        Me.txtTranscription.Size = New System.Drawing.Size(630, 157)
        Me.txtTranscription.TabIndex = 5
        Me.txtTranscription.Text = ""
        '
        'txtSummary
        '
        Me.txtSummary.Location = New System.Drawing.Point(41, 313)
        Me.txtSummary.Name = "txtSummary"
        Me.txtSummary.Size = New System.Drawing.Size(630, 125)
        Me.txtSummary.TabIndex = 6
        Me.txtSummary.Text = ""
        '
        'cboCombinedDevice
        '
        Me.cboCombinedDevice.FormattingEnabled = True
        Me.cboCombinedDevice.Location = New System.Drawing.Point(41, 72)
        Me.cboCombinedDevice.Name = "cboCombinedDevice"
        Me.cboCombinedDevice.Size = New System.Drawing.Size(414, 21)
        Me.cboCombinedDevice.TabIndex = 7
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.cboCombinedDevice)
        Me.Controls.Add(Me.txtSummary)
        Me.Controls.Add(Me.txtTranscription)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnStop)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.cboLoopbackDevice)
        Me.Controls.Add(Me.cboMicDevice)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cboMicDevice As ComboBox
    Friend WithEvents cboLoopbackDevice As ComboBox
    Friend WithEvents btnStart As Button
    Friend WithEvents btnStop As Button
    Friend WithEvents btnClear As Button
    Friend WithEvents txtTranscription As RichTextBox
    Friend WithEvents txtSummary As RichTextBox
    Friend WithEvents cboCombinedDevice As ComboBox
End Class
