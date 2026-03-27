Sub RegistrarDados()
    ' Declaração da variável de referência à planilha de destino
    Dim wsRegistro As Worksheet
    
    ' Variável que armazena a linha disponível para o novo registro
    Dim linha As Long
    
    ' Atribui a referência à planilha de destino pelo seu nome
    Set wsRegistro = Sheets("Registros")
    
    ' Valida se todos os campos obrigatórios da aba Distribuição estão preenchidos
    ' Range() sem qualificação de planilha busca na aba ativa no momento da execução
    If Range("Processo") = "" Or _
       Range("Exequente") = "" Or _
       Range("Procurador") = "" Or _
       Range("Unidade") = "" Or _
       Range("Condenação") = "" Or _
       Range("Prazo") = "" Or _
       Range("Complexidade") = "" Or _
       Range("Modo") = "" Or _
       Range("Qtde_e") = "" Or _
       Range("Qtde_c") = "" Or _
       Range("Calculista") = "" Then
       
       MsgBox "Não foi possível registrar. Existem campos obrigatórios vazios.", vbExclamation, "Registro não realizado"
       Exit Sub
       
    End If
    
    ' Localiza a próxima linha disponível a partir da linha 8 (cabeçalho reservado acima)
    ' O loop avança enquanto a coluna de Data_Registro estiver preenchida
    linha = 8
    Do While wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value <> ""
        linha = linha + 1
    Loop
    
    ' Registra a data atual na coluna correspondente e aplica a formatação padrão dd/mm/yyyy
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value = Date
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).NumberFormat = "dd/mm/yyyy"
    
    ' Copia os dados da aba Distribuição para a linha encontrada na aba Registros
    ' usando nomes definidos para referenciar as colunas de destino, evitando índices fixos
    wsRegistro.Cells(linha, wsRegistro.Range("Processo_Registro").Column).Value = Range("Processo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Exequente_Registro").Column).Value = Range("Exequente").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Procurador_Registro").Column).Value = Range("Procurador").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Unidade_Registro").Column).Value = Range("Unidade").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Condenacao_Registro").Column).Value = Range("Condenação").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Prazo_Registro").Column).Value = Range("Prazo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Complexidade_Registro").Column).Value = Range("Complexidade").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Qtde_exeq").Column).Value = Range("Qtde_e").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Qtde_calc").Column).Value = Range("Qtde_c").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Modo_Registro").Column).Value = Range("Modo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Calculista_Registro").Column).Value = Range("Calculista").Value
    
    ' Limpeza dos campos da aba Distribuição desativada para fins de teste
    ' Remover os apóstrofos abaixo para habilitar a limpeza automática após o registro
    'Range("Processo").Value = ""
    'Range("Exequente").Value = ""
    'Range("Procurador").Value = ""
    'Range("Unidade").Value = ""
    'Range("Condenação").Value = ""
    'Range("Prazo").Value = ""
    'Range("Complexidade").Value = ""
    'Range("Qtde_e").Value = ""
    'Range("Qtde_c").Value = ""
    
    MsgBox "Registro realizado com sucesso!", vbInformation, "Registro concluído"
End Sub