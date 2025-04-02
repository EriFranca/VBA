Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsDados As Worksheet
    Dim wsBase As Worksheet
    Dim rng As Range
    Dim material As String
    Dim descricao As String
    Dim Und As String
    Dim item As String
    Dim contrato As String
    Dim baseRng As Range
    Dim abaOrigem As String
    Dim valorF8 As Variant

    Set wsDados = ThisWorkbook.Worksheets("DADOS")

    If Not Intersect(Target, wsDados.Columns(1)) Is Nothing Then
        valorG8 = wsDados.Range("G8").Value
        
        If IsEmpty(valorG8) Then
            'MsgBox "Centro está vazia."
            Exit Sub
        End If

        For Each rng In Target
            material = rng.Value

            ' Define a planilha base e a aba de origem com base no valor de G8
            Select Case valorG8
                Case 5, 8, 9, 12, 13, 16, 31, 34, 38, 39
                    Set wsBase = ThisWorkbook.Worksheets("4600010381")
                    abaOrigem = "4600010381"
                Case 33, 35, 37, 41
                    Set wsBase = ThisWorkbook.Worksheets("4600010385")
                    abaOrigem = "4600010385"
                Case 40
                    Set wsBase = ThisWorkbook.Worksheets("4600010386")
                    abaOrigem = "4600010386"
                Case Else
                    MsgBox "Valor não configurado para busca de informações. Por favor, verifique o valor na célula F8."
                    Exit For
            End Select

            ' Busca o material na coluna D da planilha base
            Set baseRng = wsBase.Columns("D").Find(What:=material, LookAt:=xlWhole)

            If Not baseRng Is Nothing Then
                ' Usa as funções para buscar descrição, unidade e item
                descricao = BuscaDescricao(material, wsBase)
                Und = BuscaUnd(material, wsBase)
                item = BuscaItem(material, wsBase)
                contrato = BuscaContrato(material, wsBase)
            Else
                descricao = "Verificar contrato"
                Und = "Verificar contrato"
                item = "Verificar contrato"
                contrato = "Verificar contrato"
            End If

            ' Atualiza os valores na planilha DADOS
            wsDados.Cells(rng.Row, 2).Value = descricao   ' Coluna B (2)
            wsDados.Cells(rng.Row, 3).Value = Und         ' Coluna C (3)
            wsDados.Cells(rng.Row, 5).Value = item        ' Coluna E (5)
            wsDados.Cells(rng.Row, 6).Value = contrato    ' Coluna F (6)
            wsDados.Cells(10, 7).Value = abaOrigem        ' Informa a aba de origem na célula F10 (10,6)
        Next rng
    End If
End Sub

' Função para buscar a descrição do material na planilha base
Function BuscaDescricao(material As String, wsBase As Worksheet) As String
    Dim rng As Range
    
    Set rng = wsBase.Columns("D").Find(What:=material, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        BuscaDescricao = rng.Offset(0, 1).Value ' Descrição está na coluna E (1 à direita da coluna D)
    Else
        BuscaDescricao = "Verificar contrato" ' Caso não encontre o material
    End If
End Function

' Função para buscar a unidade do material na planilha base
Function BuscaUnd(material As String, wsBase As Worksheet) As String
    Dim rng As Range
    
    Set rng = wsBase.Columns("D").Find(What:=material, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        BuscaUnd = rng.Offset(0, 3).Value ' Unidade está na coluna G (3 à direita da coluna D)
    Else
        BuscaUnd = "Verificar contrato" ' Caso não encontre o material
    End If
End Function

' Função para buscar o "Item" correspondente ao material na planilha base e retornar para a coluna E de "DADOS"
Function BuscaItem(material As String, wsBase As Worksheet) As String
    Dim rng As Range
    
    Set rng = wsBase.Columns("D").Find(What:=material, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        BuscaItem = rng.Offset(0, -3).Value ' Item está na coluna A (-3 à esquerda da coluna D)
    Else
        BuscaItem = "Verificar contrato" ' Caso o material não seja encontrado
    End If
End Function

' Função para buscar o contrato correspondente ao material na planilha base
Function BuscaContrato(material As String, wsBase As Worksheet) As String
    Dim rng As Range
    
    Set rng = wsBase.Columns("D").Find(What:=material, LookAt:=xlWhole)
    
    If Not rng Is Nothing Then
        BuscaContrato = rng.Offset(0, 8).Value ' Contrato está na coluna L (8 à direita da coluna D)
    Else
        BuscaContrato = "Verificar contrato" ' Caso o material não seja encontrado
    End If
End Function

