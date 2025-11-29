

Public Class Form2


    Sub Principal()

        '  Me.Label3.Visible = False

        '************************************************************************************

        '        Form2.Show()
        '        For i = 1 To localCurrentProduct.ReferenceProduct.Products.Count
        '            Form2.Label1.Text = "Listing... " & localCurrentProduct.Products.Item(i).PartNumber
        '            If localDictionary.Exists(localCurrentProduct.Products.Item(i).PartNumber) Then
        '                localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber) = localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber) + 1
        '                GoTo Finish
        '            ElseIf localCurrentProduct.Products.Item(i).PartNumber = localCurrentProduct.Products.Item(i).ReferenceProduct.Parent.Product.PartNumber Then
        '                localDictionary.Add(localCurrentProduct.Products.Item(i).PartNumber, 1)
        '                FileCount(localCurrentProduct.Products.Item(i), localDictionary)
        '            End If
        'Finish:
        '        Next


        '***********************************************************************************************

        ' Form2.ProgressBar1.Value = (localDictionary.Count / intTotalArchivos) * 100
        ' For i = 1 To objCurrentProduct.Products.Count
        '    Form2.Label1.Text = "Renamming... " & objCurrentProduct.Products.Item(i).PartNumber
        '    Form2.Label2.Visible = True
        '    Form2.Label2.Text = Form2.ProgressBar1.Value & " %"
        '    ReDim Preserve arrRename(i)
        '    arrRename(i) = objCurrentProduct.Products.Item(i).PartNumber
        '    k = 0
        '    For j = 1 To i
        '        If arrRename(j) = objCurrentProduct.Products.Item(i).PartNumber Then
        '            k += 1
        '        End If
        '    Next
        '    objCurrentProduct.Products.Item(i).Name = objCurrentProduct.Products.Item(i).PartNumber & "TEMP." & k
        'Next



        '**********************************************************************************************************

        'If Form2.ProgressBar1.Value = 100 Then
        '    Form2.Label3.Visible = True
        '    Form2.Label3.Text = "Completed"
        'Else
        '    Form2.Label3.Visible = False
        'End If
        'Next



        '*********************************************************************************************************




        '        If InStr(objCurrentProduct.Products.Item(i).PartNumber, strToSearch) <> 0 Then
        '            objCurrentProduct.Products.Item(i).PartNumber =
        'Replace(Expression:=objCurrentProduct.Products.Item(i).PartNumber, Find:=strToSearch, Replacement:=strReplacement, 1, Count:=1, Compare:=1)
        '            If strOldPartNumber = objCurrentProduct.Products.Item(i).PartNumber Then
        '                If objDictionary2.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
        '                    objDictionary2.Item(objCurrentProduct.Products.Item(i).PartNumber) = objDictionary2.Item(objCurrentProduct.Products.Item(i).PartNumber) + 1
        '                    GoTo Continuar
        '                Else
        '                    objDictionary2.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
        '                    intCantidadNoRenombrada += 1
        '                End If
        '            Else
        '                intCantidadRenombrada += 1
        '            End If
        'Continuar:
        '        End If
        '        Next





        '******************************************************************************************

        'Form2.Label2.Visible = True
        'Form2.Label2.Text = Form2.ProgressBar1.Value & " %"
        'If Form2.ProgressBar1.Value = 100 Then
        '    Form2.Label3.Visible = True
        '    Form2.Label3.Text = "Completed"
        'Else
        '    Form2.Label3.Visible = False
        'End If
        'Next


    End Sub


End Class