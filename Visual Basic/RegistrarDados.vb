Sub RegistrarDados()

    ' Declaração da variável que representa a planilha de destino (Registros)
    Dim wsRegistro As Worksheet
    
    ' Variável que armazenará a próxima linha disponível para inserção dos dados
    Dim linha As Long
    
    ' Define a planilha de destino onde os dados serão registrados
    Set wsRegistro = Sheets("Registros") ' aba de destino
    
    ' Validação dos campos obrigatórios na aba Distribuição
    ' Range() sem qualificação utiliza a planilha ativa no momento da execução
    ' Caso algum campo esteja vazio, o registro não será realizado
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
    
    ' Define a linha inicial para inserção dos dados (linha 7, considerando cabeçalho acima)
    linha = 7
    
    ' Loop para encontrar a próxima linha vazia com base na coluna Data_Registro
    ' Enquanto houver valor na célula, avança para a próxima linha
    Do While wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value <> ""
        linha = linha + 1
    Loop
    
    ' Registra a data atual na coluna correspondente
    ' Aplica formatação padrão de data (dd/mm/yyyy)
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value = Date
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).NumberFormat = "dd/mm/yyyy"
    
    ' Copia os dados da aba Distribuição para a aba Registros
    ' Utiliza nomes definidos para localizar as colunas de destino dinamicamente
    
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
    
    ' Para o campo Calculista, extrai apenas o primeiro nome
    ' Trim remove espaços extras e Split separa o texto por espaço
    ' (0) indica o primeiro elemento da divisão (primeira palavra)
    wsRegistro.Cells(linha, wsRegistro.Range("Calculista_Registro").Column).Value = _
        Split(Trim(Range("Calculista").Value), " ")(0)
    
    ' Limpeza dos campos da aba Distribuição (desativada para testes)
    ' Os campos Modo e Calculista foram mantidos intencionalmente
    ' Remover os apóstrofos para ativar a limpeza automática
    
    'Range("Processo").Value = ""
    'Range("Exequente").Value = ""
    'Range("Procurador").Value = ""
    'Range("Unidade").Value = ""
    'Range("Condenação").Value = ""
    'Range("Prazo").Value = ""
    'Range("Complexidade").Value = ""
    'Range("Qtde_e").Value = ""
    'Range("Qtde_c").Value = ""
    
    ' Exibe mensagem de confirmação ao usuário após o registro bem-sucedido
    MsgBox "Registro realizado com sucesso!", vbInformation, "Registro concluído"

End Sub