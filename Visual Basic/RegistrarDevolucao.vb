Sub RegistrarDevolucao()

    ' Declaração das variáveis de referência para as planilhas
    Dim wsRegistro As Worksheet
    Dim wsDev As Worksheet
    
    ' Variáveis de controle para busca do processo
    Dim linhaProcesso As Long
    Dim processoEncontrado As Boolean
    
    ' Variáveis para armazenar os índices das colunas relevantes
    Dim colProcesso As Long
    Dim colDataEntrega As Long
    
    ' Variável que armazena a última linha preenchida na aba Registros
    Dim ultimaLinha As Long
    
    ' Define as planilhas utilizadas
    Set wsRegistro = Sheets("Registros")   ' aba onde os dados estão registrados
    Set wsDev = Sheets("Devolução")        ' aba de entrada dos dados de devolução
    
    ' Validação dos campos obrigatórios na aba Devolução
    ' Trim remove espaços extras antes e depois do conteúdo
    ' Caso algum campo esteja vazio, o registro não será realizado
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
    
    ' Obtém as colunas correspondentes aos campos Processo e Data_Entrega
    ' Utiliza nomes definidos para evitar uso de índices fixos
    colProcesso = wsRegistro.Range("Processo_Registro").Column
    colDataEntrega = wsRegistro.Range("Data_Entrega").Column
    
    ' Localiza a última linha preenchida na coluna de Processo
    ultimaLinha = wsRegistro.Cells(wsRegistro.Rows.Count, colProcesso).End(xlUp).Row
    
    ' Inicializa a variável de controle
    processoEncontrado = False
    
    ' Percorre as linhas da aba Registros a partir da linha 7 (início dos dados)
    ' Busca um processo que:
    ' 1) Tenha o mesmo número informado na aba Devolução
    ' 2) Ainda não possua data de entrega preenchida
    For linhaProcesso = 7 To ultimaLinha
        If Trim(wsRegistro.Cells(linhaProcesso, colProcesso).Value) = Trim(wsDev.Range("Processo_dev").Value) And _
           Trim(wsRegistro.Cells(linhaProcesso, colDataEntrega).Value) = "" Then
           
            ' Caso encontre, marca como verdadeiro e encerra o loop
            processoEncontrado = True
            Exit For
        End If
    Next linhaProcesso
    
    ' Caso o processo não seja encontrado ou já tenha data de entrega,
    ' exibe mensagem e interrompe a execução
    If Not processoEncontrado Then
        MsgBox "Não foi possível registrar. O número do processo não foi encontrado ou já possui data de entrega preenchida.", vbExclamation, "Processo não encontrado"
        Exit Sub
    End If
    
    ' Realiza o registro da devolução na mesma linha onde o processo foi encontrado
    ' Os dados são copiados da aba Devolução para a aba Registros
    
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Data_Entrega").Column).Value = Trim(wsDev.Range("Data_da_Entrega").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("N_Arquivo").Column).Value = Trim(wsDev.Range("Arquivo_Calc_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Valor_PGE").Column).Value = Trim(wsDev.Range("Valor_PGE_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Valor_Registro").Column).Value = Trim(wsDev.Range("Valor_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Qtde_exeq").Column).Value = Trim(wsDev.Range("Qtde_e_dev").Value)
    wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Qtde_calc").Column).Value = Trim(wsDev.Range("Qtde_c_dev").Value)
    
    ' Exibe mensagem de confirmação com informações adicionais do processo
    ' Inclui o Índice de Entrada e o nome do Exequente como retorno ao usuário
    MsgBox "Devolução registrada com sucesso!" & Chr(10) & Chr(10) & _
           "Índice: " & wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Índice_Entrada").Column).Value & Chr(10) & _
           "Exequente: " & wsRegistro.Cells(linhaProcesso, wsRegistro.Range("Exequente_Registro").Column).Value, _
           vbInformation, "Registro concluído"

End Sub