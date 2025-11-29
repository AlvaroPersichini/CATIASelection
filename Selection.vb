Module Selection

    ' NOTAS GENERALES:

    ' El "InputObject" es un array Indicar el "tipo de objeto" que se busca seleccionar.
    ' Esto establece un filtro en la selección.
    ' Si se le indica "Product" entiendo que son las ramas que cuelgan de una Product raíz.
    ' Esto incluye un Part también.
    ' Pag.98: The order of elements in the array is important. If you entered Line and Point into an array in that order,
    ' then teh line must be selected first and then the point.
    ' To help guide the user as to what needs to be selected you can display the text in the bottom corner of the screen.
    ' De esta forma se usa en los comando que necesitan multiple seleccion:
    ' "Primero seleccione un plano, luego seleccione una recta"
    ' En este caso solo se busca seleccionar un solo product.

    ' Siempre para acceder al objeto "Selection" se debe hacer que el metodo para generarlo
    ' sea contra el documento del product en el que se quiera usar la seleccion.
    ' Entonces, cuidado con usar "ActiveDocument" porque puede llegar a apuntar a un product
    ' que no es el product donde se quiere usar la seleccion. Luego de acceder al objeto hacer CleanSelection.

    ' oOutputState:
    ' The state of the selection command once SelectElement2 returns.
    ' It can be either:
    ' "Normal":the selection has succeeded)
    ' "Cancel" (the user wants to cancel the VB command, which must exit immediately)
    ' "Undo" Or "Redo".
    ' Note: The "Cancel" value Is returned if one of the following cases occured:
    ' an external command has been selected 
    ' the ESCAPE key has been selected 
    ' another window has been selected, the window document beeing another document than the current document

    ' *********************************************************************************
    ' Name        : SingleSelection2
    ' Version     : 2
    ' Purpose     : Devolver el objeto selection
    ' Inputs      : oCATIA As CATIA
    ' Output      : oSelection As INFITF.Selection 
    ' Assumptions : strObjectType As String = "Product" (se va a seleccionar un product)
    ' Locales     : 
    ' OBS         : 
    ' *********************************************************************************
    Public Function SingleSelection2(oCATIA As CatiaSession) As INFITF.Selection

        Dim oDocumentSelected As INFITF.Document
        Dim oAppCATIA As INFITF.Application = oCATIA.Application
        Dim oActiveDocu As INFITF.Document = oAppCATIA.ActiveDocument
        Dim strObjectType As String = "Product"  ' Se indica el tipo que se quiere seleccionar

        ' Cuidado, lo hace contra el "ActiveDocument". Ver comentario:
        ' "Comprobar si el objeto "selection" se hizo sobre el documento activo"
        ' Si se llama a esta funcion sin documento abierto en CATIA da error
        Dim oSelection As INFITF.Selection = oAppCATIA.ActiveDocument.Selection
        Dim strStatus As String
        Dim InputObject(0) As Object
        Dim message As String = "Select a " & strObjectType & " from the tree to proceed"

        ' Objeto Selection e invocacion a la funcion 
        oSelection.Clear()
        InputObject(0) = strObjectType  'El "InputObject" es un array Indicar el "tipo de objeto" que se busca seleccionar.
        oSelection.Clear()
        strStatus = oSelection.SelectElement2(InputObject, message, True) 'línea que invoca al método de seleccion, el programa espera por la seleccion

        ' Comprobar si el objeto "selection" se hizo sobre el documento activo (ver sección: "Interfaces VisPropertySet (Object)").
        oDocumentSelected = oSelection.Application.ActiveDocument
        If oDocumentSelected IsNot oActiveDocu Then
            MsgBox("Ha seleccionando un objeto de una ventana que no era la activa" & vbCrLf _
                   & "Click aceptar y volver a seleccionar")
            '  oSelection.Clear()
            ' el clear se hace sobre el nuevo documento activo. La línea "oSelection.Clear()" no hace nada en este caso.
            oDocumentSelected.Selection.Clear()
            Return Nothing
            Exit Function
        End If

        ' oOutputState:
        Select Case strStatus
            Case "Normal"
                Exit Select
            Case "Cancel"
                oSelection.Clear()
                Return Nothing
                Exit Function
            Case "Undo"
                oSelection.Clear()
                Return Nothing
                Exit Function
        End Select
        Return oSelection
    End Function

End Module
