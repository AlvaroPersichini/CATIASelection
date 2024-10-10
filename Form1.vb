Option Explicit On

Imports System.Runtime.InteropServices

Public Class Form1
    ' Declarar las funciones de la API de Windows para el comportamiento del formulario
    <DllImport("user32.dll")>
    Private Shared Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    End Function

    <DllImport("user32.dll")>
    Private Shared Function IsIconic(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Shared Function IsZoomed(ByVal hWnd As IntPtr) As Boolean
    End Function



    Dim oCATIA As CATIA
    Dim oSel As INFITF.Selection
    Dim oSelectedElement As INFITF.SelectedElement
    Dim oSelectedProduct As ProductStructureTypeLib.Product
    Dim seleccion As Boolean = False
    Dim cancelado As Boolean
    Dim Superficie As Graphics = Me.CreateGraphics
    Dim rect As Rectangle
    Dim folderBrowser As New FolderBrowserDialog
    Dim folderOpened As Boolean
    Dim isComboBoxOpen As Boolean = False

    Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        ' Se ejecuta cuando el ComboBox se despliega
        isComboBoxOpen = True
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' Cargar los últimos directorios almacenados en My.Settings
        If My.Settings.UltimosDirectorios IsNot Nothing Then

            ComboBox1.Items.AddRange(My.Settings.UltimosDirectorios.Cast(Of String).ToArray())

        End If

        Me.Text = "BOM - Selection"
        Me.ShowIcon = False
        Me.BackColor = Color.FromArgb(255, 241, 213)
        Me.ShowInTaskbar = False
        Label1.Text = "No Selection"
        Label1.TextAlign = ContentAlignment.MiddleLeft


        rect = Label1.DisplayRectangle

        oCATIA = New CATIA

        BackgroundWorker1.RunWorkerAsync()

    End Sub




    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork

        seleccion = False ' por si se hace click en el campo de seleccion, para que no vuelve a llamar a la funcion

        oSel = Nothing

        ' antes de graficar el rectangulo para simular el foco en el control, espero 50ms para que no se borre
        System.Threading.Thread.Sleep(50)
        Label1.CreateGraphics().DrawRectangle(New Pen(Color.Black, 2) With {.DashStyle = Drawing2D.DashStyle.Dot}, rect)


        ' invoca a la funcion y espera por la seleccion del usuario
        oSel = Selection.SingleSelection2(oCATIA)  ' aca va a esperar a la seleccion


        ' Hay que evitar el "Undo" en algunas circunstancias:
        ' Si el usuario aprieta "ESC" antes de seleccionar algo
        ' entonces la funcion de seleccion va a devolver "nothing"

        If oSel Is Nothing Then


            cancelado = True ' es una bandera para otro subprocesos

            End  ' CUIDADO: aca estoy utilizando "End" / hay que ver como finalizar la applicacion

        End If


        ' si ya paso por la llamada a la funcion de seleccion entonces, si puede volver a llamar haciendo click en el campo de seleccion
        seleccion = True


        ' Re dibuja el rectangulo con el color de fondo del formulario para ocultarlo
        Label1.CreateGraphics().DrawRectangle(New Pen(Label1.BackColor, 2) With {.DashStyle = Drawing2D.DashStyle.Dot}, rect)


        ' referencia el elemento seleccionado
        oSelectedElement = oSel.Item2(1)

        ' referencia el product contenido en oSelectedelement
        oSelectedProduct = oSelectedElement.Value

    End Sub


    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted

        If cancelado = True Then
            Exit Sub
        End If

        Label1.Text = oSelectedProduct.PartNumber
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If Label1.Text = "No Selection" Then
            Exit Sub
        Else

            Form2.Show()


        End If

    End Sub


    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

        If seleccion = False Then
            Exit Sub
        End If
        oSel.Clear()
        BackgroundWorker1.RunWorkerAsync()

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        'Si el usuario no llegó a seleccionar algo y aprieta "ESC" entonces no hay que hacer "Undo"

        If oSel Is Nothing And Label1.Text = "No Selection" Then

            oCATIA.AppCATIA.StartCommand("Undo")
            End  ' CUIDADO: aca estoy utilizando "End" / hay que ver como finalizar la applicacion

        End If

        If oSel Is Nothing And Label1.Text <> "No Selection" Then
            oCATIA.AppCATIA.StartCommand("Undo")
            Exit Sub
        End If


    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click



        ' esto es una bandera para que el formulario "folderBrowser" no pierda el foco
        folderOpened = True



        ' Mostrar el diálogo y verificar si se seleccionó una carpeta
        If folderBrowser.ShowDialog() = DialogResult.OK Then
            Dim selectedDirectory As String = folderBrowser.SelectedPath
            ' Carga la ruta seleccionada en el combobox

            ComboBox1.Text = folderBrowser.SelectedPath

            ' Agregar el directorio a My.Settings si no está ya en la lista
            If Not My.Settings.UltimosDirectorios.Contains(selectedDirectory) Then

                My.Settings.UltimosDirectorios.Add(selectedDirectory)

                ' Mantener solo los últimos 5 directorios
                If My.Settings.UltimosDirectorios.Count > 5 Then
                    My.Settings.UltimosDirectorios.RemoveAt(5) ' Eliminar el más antiguo
                End If
            End If

            ' Guardar la configuración
            My.Settings.Save()

        End If
        ' bandera para que el comportamiento de visibilidad del formulario
        folderOpened = False


        If ComboBox1.SelectedItem Is Nothing Then
            Exit Sub

        End If


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        ' Si el usuario no llegó a seleccionar algo y aprieta "ESC" entonces no hay que hacer "Undo"

        If oSel Is Nothing Then
            oCATIA.AppCATIA.StartCommand("Undo")
        Else
            End  ' CUIDADO: aca estoy utilizando "End"
        End If

    End Sub


    '**************************************************************************
    ' Esto que sigue a continuacion es para el comportamiento del formulario.
    ' Si CATIA tiene el foco, entonces el fomulario es Topmost
    ' Si CATIA se minimiza entonces el formulario se minimiza
    ' Si CATIA se maximiza entonces el fomulario aparece
    '**************************************************************************

    ' Temporizador para verificar periódicamente la ventana activa
    Private WithEvents CheckForegroundAppTimer As Timer
    Public Sub New()
        ' Inicializar el formulario
        InitializeComponent()


        ' Configurar el temporizador / Verificar cada 1 segundo
        CheckForegroundAppTimer = New Timer With {
            .Interval = 100
            }
        CheckForegroundAppTimer.Start()

    End Sub

    ' Evento que se ejecuta cada vez que el temporizador hace tick
    Private Sub CheckForegroundApp(sender As Object, e As EventArgs) Handles CheckForegroundAppTimer.Tick
        ' Obtener la ventana en primer plano
        Dim hwnd As IntPtr = GetForegroundWindow()
        Dim processID As UInteger
        GetWindowThreadProcessId(hwnd, processID)

        ' Obtener el proceso en primer plano
        Dim foregroundProcess As Process = Process.GetProcessById(CInt(processID))


        ' Si se está mostrando el formulario de progreso, entonces que no siga con lo demas
        If Form2.Visible = True Then
            Exit Sub
        End If


        ' Si si está mostrando el formulario de folderBrowser entonces que no siga con lo demas
        If folderOpened Then
            Exit Sub
        End If

        If isComboBoxOpen = True Then
            Exit Sub
        End If

        ' Comparar con el nombre del proceso de la aplicación deseada (por ejemplo "notepad")
        If foregroundProcess.ProcessName.ToLower() = "cnext" Or foregroundProcess.ProcessName.ToLower() = "catiaselection" Then
            ' La aplicación tiene el foco, mostrar el formulario y hacerlo TopMost
            Me.TopMost = True
            Me.Show()

            ' Verificar si la aplicación está minimizada
            If IsIconic(hwnd) Then
                Me.WindowState = FormWindowState.Minimized ' Minimizar el formulario
            Else
                Me.WindowState = FormWindowState.Normal ' Restaurar el formulario si no está minimizada
            End If
        Else
            ' Si la aplicación no tiene el foco
            Me.TopMost = False ' No es TopMost
            Me.Hide() ' Ocultar el formulario
        End If
    End Sub


End Class