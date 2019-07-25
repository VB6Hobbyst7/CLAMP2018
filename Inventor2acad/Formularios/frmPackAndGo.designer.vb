Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPackAndGo
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPackAndGo))
        Me.lblDatos = New System.Windows.Forms.Label()
        Me.pb1 = New System.Windows.Forms.ProgressBar()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.cbOpen = New System.Windows.Forms.CheckBox()
        Me.lblASM = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblDatos
        '
        Me.lblDatos.Location = New System.Drawing.Point(0, 0)
        Me.lblDatos.Name = "lblDatos"
        Me.lblDatos.Size = New System.Drawing.Size(100, 23)
        Me.lblDatos.TabIndex = 6
        '
        'pb1
        '
        Me.pb1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pb1.Location = New System.Drawing.Point(12, 93)
        Me.pb1.Name = "pb1"
        Me.pb1.Size = New System.Drawing.Size(648, 23)
        Me.pb1.TabIndex = 5
        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.Enabled = False
        Me.btnClose.Location = New System.Drawing.Point(585, 129)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 28)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(389, 129)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(162, 28)
        Me.btnStart.TabIndex = 3
        Me.btnStart.Text = "Start PachAndGo"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'cbOpen
        '
        Me.cbOpen.AutoSize = True
        Me.cbOpen.Checked = True
        Me.cbOpen.CheckState = System.Windows.Forms.CheckState.Checked
        Me.cbOpen.Location = New System.Drawing.Point(15, 12)
        Me.cbOpen.Name = "cbOpen"
        Me.cbOpen.Size = New System.Drawing.Size(195, 21)
        Me.cbOpen.TabIndex = 4
        Me.cbOpen.Text = "Open folder when finished"
        Me.cbOpen.UseVisualStyleBackColor = True
        '
        'lblASM
        '
        Me.lblASM.Location = New System.Drawing.Point(12, 51)
        Me.lblASM.Name = "lblASM"
        Me.lblASM.Size = New System.Drawing.Size(648, 22)
        Me.lblASM.TabIndex = 7
        Me.lblASM.Text = "PackAndGo ---> 0 of XX Files"
        '
        'frmPackAndGo
        '
        Me.AcceptButton = Me.btnStart
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(672, 168)
        Me.Controls.Add(Me.lblASM)
        Me.Controls.Add(Me.cbOpen)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.pb1)
        Me.Controls.Add(Me.lblDatos)
        Me.Cursor = System.Windows.Forms.Cursors.AppStarting
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPackAndGo"
        Me.Text = "PackAndGo"
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblDatos As Label
    Friend WithEvents pb1 As ProgressBar
    Friend WithEvents btnClose As Button
    Friend WithEvents btnStart As Button
    Friend WithEvents cbOpen As CheckBox
    Friend WithEvents lblASM As Label
End Class
