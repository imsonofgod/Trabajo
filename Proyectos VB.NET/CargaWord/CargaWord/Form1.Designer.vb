<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
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

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.listbox2 = New System.Windows.Forms.ListBox()
        Me.lstBox = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(653, 295)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(119, 43)
        Me.Button1.TabIndex = 22
        Me.Button1.Text = "Cerrar"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(116, 312)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(13, 13)
        Me.Label4.TabIndex = 21
        Me.Label4.Text = "0"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(54, 295)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(150, 13)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "Numero de Rutas Examinadas"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(470, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(228, 13)
        Me.Label2.TabIndex = 19
        Me.Label2.Text = "Documentos no cargados durante la ejecucion"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(116, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(109, 13)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "Documentos a cargar"
        '
        'listbox2
        '
        Me.listbox2.FormattingEnabled = True
        Me.listbox2.Location = New System.Drawing.Point(408, 87)
        Me.listbox2.Name = "listbox2"
        Me.listbox2.Size = New System.Drawing.Size(364, 186)
        Me.listbox2.TabIndex = 17
        '
        'lstBox
        '
        Me.lstBox.FormattingEnabled = True
        Me.lstBox.Location = New System.Drawing.Point(20, 87)
        Me.lstBox.Name = "lstBox"
        Me.lstBox.Size = New System.Drawing.Size(364, 186)
        Me.lstBox.TabIndex = 16
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(795, 359)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.listbox2)
        Me.Controls.Add(Me.lstBox)
        Me.Name = "Form1"
        Me.Text = "Carga de Texto locución"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents listbox2 As System.Windows.Forms.ListBox
    Friend WithEvents lstBox As System.Windows.Forms.ListBox

End Class
