Public Class FormFormatter

    ' ******************************************************
    ' Método para dar formato a un formulario de selección
    ' ******************************************************
    Public Sub GiveSelectionFormat(form As Form)

        ' form
        With form
            .AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            .AllowDrop = True
            .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            .BackColor = System.Drawing.SystemColors.Control
            .ClientSize = New System.Drawing.Size(357, 225)
            .Name = "Form1"
            .Text = "Form1"
            .Text = "BOM - Selection "
            .ShowIcon = False
            .BackColor = Color.FromArgb(255, 241, 213)
            .Size = New System.Drawing.Size(373, 264)
            .MinimizeBox = False
            .MaximizeBox = False
            .FormBorderStyle = FormBorderStyle.FixedDialog
        End With


        ' Label1 
        Dim Label1 As New Label With {
            .BackColor = System.Drawing.SystemColors.MenuHighlight,
            .BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle,
            .FlatStyle = System.Windows.Forms.FlatStyle.Popup,
            .ForeColor = System.Drawing.SystemColors.ButtonHighlight,
            .Location = New System.Drawing.Point(70, 38),
            .Name = "Label1",
            .Size = New System.Drawing.Size(198, 19),
            .TabIndex = 2,
            .Text = "No Selection",
            .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
             }
        form.Controls.Add(Label1)


        ' Label2
        Dim Label2 As New Label With {
            .AutoSize = True,
            .ForeColor = System.Drawing.SystemColors.ActiveCaptionText,
            .Location = New System.Drawing.Point(15, 38),
            .Name = "Label2",
            .Size = New System.Drawing.Size(54, 13),
            .TabIndex = 3,
            .Text = "Selection:"
            }
        form.Controls.Add(Label2)


        ' Label3
        Dim Label3 As New Label With {
            .AutoSize = True,
            .ForeColor = System.Drawing.SystemColors.ActiveCaptionText,
            .Location = New System.Drawing.Point(15, 138),
            .Name = "Label3",
            .Size = New System.Drawing.Size(52, 13),
            .TabIndex = 5,
            .Text = "Directory:"
            }
        form.Controls.Add(Label3)



        ' FolderBrowserDialog1 (esta asociado al combobox)
        Dim FolderBrowserDialog1 As New FolderBrowserDialog With {
            .RootFolder = System.Environment.SpecialFolder.MyComputer,
            .SelectedPath = "C:\"
            }



        ' Button1
        Dim Button1 As New Button With {
            .Location = New System.Drawing.Point(203, 184),
            .Name = "Button1",
            .Size = New System.Drawing.Size(68, 24),
            .TabIndex = 1,
            .Text = "OK",
            .UseVisualStyleBackColor = True
            }
        form.Controls.Add(Button1)


        'Button2
        Dim Button2 As New Button With {
            .Location = New System.Drawing.Point(277, 135),
            .Name = "Button2",
            .Size = New System.Drawing.Size(67, 21),
            .TabIndex = 4,
            .Text = "Browse...",
            .UseVisualStyleBackColor = True
            }
        form.Controls.Add(Button2)


        'Button3
        Dim Button3 As New Button With {
            .Location = New System.Drawing.Point(277, 184),
            .Name = "Button3",
            .Size = New System.Drawing.Size(68, 24),
            .TabIndex = 8,
            .Text = "Cancel",
            .UseVisualStyleBackColor = True
            }
        form.Controls.Add(Button3)


        'ComboBox1
        Dim ComboBox1 As New ComboBox With {
            .FormattingEnabled = True,
            .Location = New System.Drawing.Point(73, 135),
            .Name = "ComboBox1",
            .Size = New System.Drawing.Size(198, 21),
            .TabIndex = 7
            }
        form.Controls.Add(ComboBox1)


        ' Checkbox1
        Dim CheckBox1 As New CheckBox With {
            .AutoSize = True,
            .Location = New System.Drawing.Point(71, 66),
            .Name = "CheckBox1",
            .Size = New System.Drawing.Size(102, 17),
            .TabIndex = 4,
            .Text = "Add NCU Sheet",
            .UseVisualStyleBackColor = True
            }
        form.Controls.Add(CheckBox1)


        'CheckBox2
        Dim CheckBox2 As New CheckBox With {
            .AutoSize = True,
            .Location = New System.Drawing.Point(71, 90),
            .Name = "CheckBox2",
            .Size = New System.Drawing.Size(98, 17),
            .TabIndex = 5,
            .Text = "Include Images",
            .UseVisualStyleBackColor = True
            }
        form.Controls.Add(CheckBox2)


        ' GroupBox1
        Dim GroupBox1 As New GroupBox With {
            .Location = New System.Drawing.Point(8, 13),
            .Name = "GroupBox1",
            .Size = New System.Drawing.Size(336, 107),
            .TabIndex = 9,
            .TabStop = False,
            .Text = "Options"
        }
        form.Controls.Add(GroupBox1)


        AddHandler Button3.Click, AddressOf Button3_Click


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub



    ' *******************************************************
    ' Metodo para dar formato a un formulario de progressBar
    ' *******************************************************
    Public Sub GiveProgressBarFormat(form As Form)

        ' Formato dgeneral
        With form
            .SuspendLayout()
            .BackColor = Color.FromArgb(255, 241, 213)
            .AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
            .AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
            .ClientSize = New System.Drawing.Size(346, 83)
            .Name = "Form2"
            .Text = "Please Wait..."
            .ResumeLayout(False)
            .PerformLayout()
            .ShowIcon = False
            .ShowInTaskbar = False
        End With


        ' Crear una ProgressBar1
        Dim ProgressBar1 As New ProgressBar With {
            .Location = New System.Drawing.Point(12, 44),
            .Size = New System.Drawing.Size(322, 14),
            .Name = "ProgressBar1",
            .TabIndex = 4
            }
        form.Controls.Add(ProgressBar1)


        ' Crear un Label1 para el texto "Processing..."
        Dim Label1 As New Label With {
            .AutoSize = True,
            .Text = "Processing...",
            .Location = New System.Drawing.Point(12, 61),
            .Size = New System.Drawing.Size(68, 13),
            .Name = "Label1",
            .TabIndex = 1
        }
        form.Controls.Add(Label1)


        ' Crear un Label2 para el porcentaje "100%"
        Dim Label2 As New Label With {
            .Anchor = System.Windows.Forms.AnchorStyles.Right,
            .Name = "Label2",
            .Text = "100%",
            .TabIndex = 2,
            .Location = New System.Drawing.Point(233, 24),
            .Size = New System.Drawing.Size(101, 17),
            .TextAlign = System.Drawing.ContentAlignment.MiddleRight
        }
        form.Controls.Add(Label2)


        ' Crear un Label3 para el texto "Completed"
        Dim Label3 As New Label With {
            .AutoSize = True,
            .Location = New System.Drawing.Point(277, 61),
            .Name = "Label3",
            .Size = New System.Drawing.Size(57, 13),
            .TabIndex = 5,
            .Text = "Completed",
            .TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        }
        form.Controls.Add(Label3)

    End Sub



End Class