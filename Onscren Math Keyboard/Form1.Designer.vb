﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.cboWindows = New System.Windows.Forms.ComboBox()
        Me.SuspendLayout()
        '
        'cboWindows
        '
        Me.cboWindows.DropDownHeight = 500
        Me.cboWindows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboWindows.DropDownWidth = 220
        Me.cboWindows.Font = New System.Drawing.Font("Microsoft Sans Serif", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboWindows.FormattingEnabled = True
        Me.cboWindows.IntegralHeight = False
        Me.cboWindows.ItemHeight = 18
        Me.cboWindows.Location = New System.Drawing.Point(0, 0)
        Me.cboWindows.Margin = New System.Windows.Forms.Padding(1, 9, 1, 9)
        Me.cboWindows.Name = "cboWindows"
        Me.cboWindows.Size = New System.Drawing.Size(350, 26)
        Me.cboWindows.TabIndex = 0
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 84.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1884, 146)
        Me.Controls.Add(Me.cboWindows)
        Me.Font = New System.Drawing.Font("Cambria Math", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 33, 4, 33)
        Me.Name = "frmMain"
        Me.Text = "Math Keyboard"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents cboWindows As ComboBox
End Class
