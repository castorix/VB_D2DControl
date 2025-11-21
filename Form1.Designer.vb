<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    Friend WithEvents D2DControl1 As D2DControl
    Friend WithEvents D2DControl2 As D2DControl
    Friend WithEvents D2DControl3 As D2DControl

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.cbOctopus = New System.Windows.Forms.CheckBox()
        Me.cbGirl = New System.Windows.Forms.CheckBox()
        Me.cbButterfly = New System.Windows.Forms.CheckBox()
        Me.cbDragon = New System.Windows.Forms.CheckBox()
        Me.D2DControl3 = New VB_D2DControl.D2DControl()
        Me.D2DControl2 = New VB_D2DControl.D2DControl()
        Me.D2DControl1 = New VB_D2DControl.D2DControl()
        Me.D2DControl4 = New VB_D2DControl.D2DControl()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'PictureBox1
        '
        Me.PictureBox1.Location = New System.Drawing.Point(228, 29)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(533, 325)
        Me.PictureBox1.TabIndex = 7
        Me.PictureBox1.TabStop = False
        '
        'cbOctopus
        '
        Me.cbOctopus.AutoSize = True
        Me.cbOctopus.Location = New System.Drawing.Point(35, 29)
        Me.cbOctopus.Name = "cbOctopus"
        Me.cbOctopus.Size = New System.Drawing.Size(116, 17)
        Me.cbOctopus.TabIndex = 8
        Me.cbOctopus.Text = "Octopus draggable"
        Me.cbOctopus.UseVisualStyleBackColor = True
        '
        'cbGirl
        '
        Me.cbGirl.AutoSize = True
        Me.cbGirl.Location = New System.Drawing.Point(35, 74)
        Me.cbGirl.Name = "cbGirl"
        Me.cbGirl.Size = New System.Drawing.Size(91, 17)
        Me.cbGirl.TabIndex = 9
        Me.cbGirl.Text = "Girl draggable"
        Me.cbGirl.UseVisualStyleBackColor = True
        '
        'cbButterfly
        '
        Me.cbButterfly.AutoSize = True
        Me.cbButterfly.Location = New System.Drawing.Point(35, 164)
        Me.cbButterfly.Name = "cbButterfly"
        Me.cbButterfly.Size = New System.Drawing.Size(114, 17)
        Me.cbButterfly.TabIndex = 10
        Me.cbButterfly.Text = "Butterfly draggable"
        Me.cbButterfly.UseVisualStyleBackColor = True
        '
        'cbDragon
        '
        Me.cbDragon.AutoSize = True
        Me.cbDragon.Location = New System.Drawing.Point(35, 119)
        Me.cbDragon.Name = "cbDragon"
        Me.cbDragon.Size = New System.Drawing.Size(111, 17)
        Me.cbDragon.TabIndex = 12
        Me.cbDragon.Text = "Dragon draggable"
        Me.cbDragon.UseVisualStyleBackColor = True
        '
        'D2DControl3
        '
        Me.D2DControl3.BackColor = System.Drawing.Color.Magenta
        Me.D2DControl3.Location = New System.Drawing.Point(12, 97)
        Me.D2DControl3.Name = "D2DControl3"
        Me.D2DControl3.Size = New System.Drawing.Size(127, 104)
        Me.D2DControl3.TabIndex = 6
        '
        'D2DControl2
        '
        Me.D2DControl2.BackColor = System.Drawing.Color.Black
        Me.D2DControl2.Location = New System.Drawing.Point(251, 397)
        Me.D2DControl2.Name = "D2DControl2"
        Me.D2DControl2.Size = New System.Drawing.Size(119, 100)
        Me.D2DControl2.TabIndex = 5
        '
        'D2DControl1
        '
        Me.D2DControl1.BackColor = System.Drawing.Color.Red
        Me.D2DControl1.Location = New System.Drawing.Point(341, 97)
        Me.D2DControl1.Name = "D2DControl1"
        Me.D2DControl1.Size = New System.Drawing.Size(157, 60)
        Me.D2DControl1.TabIndex = 4
        '
        'D2DControl4
        '
        Me.D2DControl4.BackColor = System.Drawing.Color.Blue
        Me.D2DControl4.Location = New System.Drawing.Point(529, 383)
        Me.D2DControl4.Name = "D2DControl4"
        Me.D2DControl4.Size = New System.Drawing.Size(109, 79)
        Me.D2DControl4.TabIndex = 11
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(796, 652)
        Me.Controls.Add(Me.D2DControl4)
        Me.Controls.Add(Me.D2DControl1)
        Me.Controls.Add(Me.D2DControl2)
        Me.Controls.Add(Me.D2DControl3)
        Me.Controls.Add(Me.cbDragon)
        Me.Controls.Add(Me.cbButterfly)
        Me.Controls.Add(Me.cbGirl)
        Me.Controls.Add(Me.cbOctopus)
        Me.Controls.Add(Me.PictureBox1)
        Me.Name = "Form1"
        Me.Text = "Test Direct2D layered control (for dragging and transparency with animated sprite" &
    ")"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PictureBox1 As PictureBox
    Friend WithEvents cbOctopus As CheckBox
    Friend WithEvents cbGirl As CheckBox
    Friend WithEvents cbButterfly As CheckBox
    Friend WithEvents D2DControl4 As D2DControl
    Friend WithEvents cbDragon As CheckBox
End Class
