<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Actions
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

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.ForeColor = System.Drawing.Color.CornflowerBlue
        Me.Button1.Location = New System.Drawing.Point(25, 12)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(150, 39)
        Me.Button1.TabIndex = 0
        Me.Button1.Text = "Make Bounding Box"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.ForeColor = System.Drawing.Color.DarkCyan
        Me.Button2.Location = New System.Drawing.Point(25, 69)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(150, 39)
        Me.Button2.TabIndex = 0
        Me.Button2.Text = "Hide All Bounding Box"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.ForeColor = System.Drawing.Color.DarkOrchid
        Me.Button3.Location = New System.Drawing.Point(25, 181)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(150, 39)
        Me.Button3.TabIndex = 0
        Me.Button3.Text = "Export to Excel"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.ForeColor = System.Drawing.Color.Crimson
        Me.Button4.Location = New System.Drawing.Point(25, 125)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(150, 39)
        Me.Button4.TabIndex = 0
        Me.Button4.Text = "Show All Bounding Box"
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Actions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(200, 232)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Name = "Actions"
        Me.Text = "Actions"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button4 As Button
End Class
