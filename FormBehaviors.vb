Public Class FormBehaviors
    ' *******************************************************************************
    ' FormBehavior → se encarga de inyectar lógica, handlers, timers y APIs externas.
    ' *******************************************************************************



    Public Sub GiveFormBehaviors(form As Form)
        ' Atributos del formulario
        With form
            .StartPosition = FormStartPosition.CenterScreen
            .TopMost = True
            .ShowInTaskbar = False
        End With
    End Sub

End Class
