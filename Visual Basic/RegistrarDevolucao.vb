Sub RegistrarDevolucao()
    ' Declaração das variáveis de referência às planilhas
    Dim wsRegistro As Worksheet
    Dim wsDev As Worksheet
    
    ' Variável que armazena a linha onde o processo foi encontrado na aba Registros
    Dim linhaProcesso As Long
    
    ' Flag que indica se o processo foi localizado com data de entrega vazia
    Dim processoEncontrado As Boolean
    
    ' Variáveis que armazenam os índices das colunas para evitar chamadas repetidas ao Range
    Dim colProcesso As Long
    Dim colDataEntrega As Long
    
    ' Última linha preenchida na coluna do processo, usada como limite do loop de busca
    Dim ultimaLinha As Long
    
    ' Atribui as referências às planilhas pelos seus nomes
    Set wsRegistro = Sheets("Registros")
    Set wsDev = Sheets("Devolução")
    
    ' Valida se todos os campos obrigatórios da aba Devolução estão preenchidos
    ' Trim() garante que células com apenas espaços sejam tratadas como vazias
    If Trim(wsDev.Range("Data_da_Entrega").Value) = "" Or _
       Trim(wsDev.Range("Arquivo_Calc_dev").Value) = "" Or _
       Trim(wsDev.Range("Valor_PGE_dev").Value) = "" Or _
       Trim(wsDev.Range("Valor_dev").Value) = "" Or _
       Trim(wsDev.Range("Qtde_e_dev").Value) = "" Or _
       Trim(wsDev.Range("Qtde_c_dev").Value) = "" Or _
       Trim(wsDev.Range("Processo_dev").Value) = "" Then
        
        MsgBox "Não foi possível registrar. Existem campos obrigatórios vazios.", vbExclamation, "Registro não realizado"
        Exit Sub
    End If
    
    ' Obtém os índices das colunas de processo e data de entrega na aba Registros
    ' usando nomes definidos para desacoplar o código de posições fixas de coluna
    colProcesso = wsRegistro.Range("Processo_Registro").Column
    colDataEntrega = wsRegistro.Range("Data_Entrega").Column
    
    ' Identifica a última linha preenchida na coluna do processo
    ' End(xlUp) equivale ao atalho Ctrl+Seta para cima a partir da última linha da planilha
    ultimaLinha = wsRegistro.Cells(wsRegistro.Rows.Count, colProcesso).End(xlUp).Row
    processoEncontrado = False
    
    ' Percorre as linhas a partir da linha 8 (cabeçalho reservado acima)
    ' buscando o processo que coincide com o digitado e ainda não possui data de entrega
    ' Trim() em ambos os lados evita falhas por espaços acidentais no início ou fim
    For linhaProcesso = 8 To ultimaLinha
        If Trim(wsRegistro.Cells(linhaProcesso, colProcesso).Value) = Trim(wsDev.Range("Processo_dev").Value) And _
           Trim(wsRegistro.Cells(linhaProcesso, colDataEntrega).Value) = "" Then
            processoEncontrado = True
            ' Interrompe o loop assim que encontrar a primeira ocorrência válida
            Exit For
        End If
    Next linhaProcesso
    
    ' Encerra a execução caso o processo não exista ou já esteja com data de entrega preenchida
    If Not processoEncontrado Then
        MsgBox "Não foi possível registrar. O número do processo não foi encontrado ou já possui data de entrega preenchida.", vbExclamation, "Processo não encontrado"
        Exit Sub
    End If
    
    ' Grava os dados da aba Devolução na linha exata onde o processo foi encontrado
    ' Trim() aplicado nos valores gravados para evitar espaços desnecessários nos registros
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Data_Entrega").Column).Value = Trim(wsDev.Range("Data_da_Entrega").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("N_Arquivo").Column).Value = Trim(wsDev.Range("Arquivo_Calc_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Valor_PGE").Column).Value = Trim(wsDev.Range("Valor_PGE_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Valor_Registro").Column).Value = Trim(wsDev.Range("Valor_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Qtde_exeq").Column).Value = Trim(wsDev.Range("Qtde_e_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Qtde_calc").Column).Value = Trim(wsDev.Range("Qtde_c_dev").Value)
    
    MsgBox "Devolução registrada com sucesso!", vbInformation, "Registro concluído"
End Sub