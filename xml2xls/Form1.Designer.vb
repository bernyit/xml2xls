<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
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

    'Richiesto da Progettazione Windows Form
    Private components As System.ComponentModel.IContainer

    'NOTA: la procedura che segue è richiesta da Progettazione Windows Form
    'Può essere modificata in Progettazione Windows Form.  
    'Non modificarla mediante l'editor del codice.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form1))
        Me.btnConvert = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.txtSource = New System.Windows.Forms.TextBox()
        Me.btnSelectSource = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.btnSelectDestination = New System.Windows.Forms.Button()
        Me.txtDestination = New System.Windows.Forms.TextBox()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'btnConvert
        '
        Me.btnConvert.Location = New System.Drawing.Point(553, 130)
        Me.btnConvert.Name = "btnConvert"
        Me.btnConvert.Size = New System.Drawing.Size(147, 41)
        Me.btnConvert.TabIndex = 0
        Me.btnConvert.Text = "Convert XML 2 XLSX"
        Me.btnConvert.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "XML File|*.xml|All Files|*.*"
        '
        'txtSource
        '
        Me.txtSource.Location = New System.Drawing.Point(101, 32)
        Me.txtSource.Name = "txtSource"
        Me.txtSource.Size = New System.Drawing.Size(554, 20)
        Me.txtSource.TabIndex = 1
        '
        'btnSelectSource
        '
        Me.btnSelectSource.Location = New System.Drawing.Point(662, 32)
        Me.btnSelectSource.Name = "btnSelectSource"
        Me.btnSelectSource.Size = New System.Drawing.Size(38, 20)
        Me.btnSelectSource.TabIndex = 2
        Me.btnSelectSource.Text = "..."
        Me.btnSelectSource.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 35)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "XML Source"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 78)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "XLS Destination"
        '
        'btnSelectDestination
        '
        Me.btnSelectDestination.Location = New System.Drawing.Point(662, 75)
        Me.btnSelectDestination.Name = "btnSelectDestination"
        Me.btnSelectDestination.Size = New System.Drawing.Size(38, 20)
        Me.btnSelectDestination.TabIndex = 5
        Me.btnSelectDestination.Text = "..."
        Me.btnSelectDestination.UseVisualStyleBackColor = True
        '
        'txtDestination
        '
        Me.txtDestination.Location = New System.Drawing.Point(101, 75)
        Me.txtDestination.Name = "txtDestination"
        Me.txtDestination.Size = New System.Drawing.Size(554, 20)
        Me.txtDestination.TabIndex = 4
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = "Excel File|*.xlsx|All Files|*.*"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(717, 199)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnSelectDestination)
        Me.Controls.Add(Me.txtDestination)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSelectSource)
        Me.Controls.Add(Me.txtSource)
        Me.Controls.Add(Me.btnConvert)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form1"
        Me.Text = "     XML 2 XLSX"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnConvert As Button
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents txtSource As TextBox
    Friend WithEvents btnSelectSource As Button
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents btnSelectDestination As Button
    Friend WithEvents txtDestination As TextBox
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
End Class
