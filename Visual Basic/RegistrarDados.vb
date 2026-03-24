Sub RegistrarDados()
'Sub indica que é uma sub-rotina (macro) que executa uma série de comandos.
'RegistrarDados é o nome da macro, que você pode chamar ao clicar em um botão ou executando no VBA.

    Dim wsRegistro As Worksheet 'wsRegistro → Variável que vai guardar a aba de destino (Registros).
    Dim linha As Long 'linha → variável numérica para controlar qual linha da aba Registros vai receber os dados
     
    Set wsRegistro = Sheets("Registros") 'aba de destino. Associa a variável wsRegistro à aba Registros. Todas as escritas de dados vão usar wsRegistro para garantir que não vai escrever na aba errada.

    'Verificar campos obrigatórios na aba Distribuição
    'Antes de registrar, verifica se todos os campos importantes da aba Distribuição estão preenchidos.
    'Se algum estiver vazio → exibe uma mensagem de aviso e encerra a macro (Exit Sub).
    If Range("Processo") = "" Or _
       Range("Exequente") = "" Or _
       Range("Tipo") = "" Or _
       Range("Procurador") = "" Or _
       Range("Unidade") = "" Or _
       Range("Condenação") = "" Or _
       Range("Valor") = "" Or _
       Range("Prazo") = "" Or _
       Range("Complexidade") = "" Or _
       Range("Modo") = "" Or _
       Range("Calculista") = "" Then
       
       MsgBox "Não foi possível registrar. Existem campos obrigatórios vazios.", vbExclamation, "Registro não realizado"
       Exit Sub
       
    End If

    'Começar da linha 8 na aba Registros
    'Procura a primeira linha vazia na coluna "Data_Registro" da aba Registros.
    'Isso garante que não sobrescreva registros anteriores.
    linha = 8
    Do While wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value <> ""
        linha = linha + 1
    Loop

    'Registrar data atual
    'Coloca a data atual na coluna "Data_Registro" da linha encontrada. Formata como dd/mm/aaaa.
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).Value = Date
    wsRegistro.Cells(linha, wsRegistro.Range("Data_Registro").Column).NumberFormat = "dd/mm/yyyy"

    'Copiar dados de Distribuição para Registros usando nomes
    'Cada linha pega o valor de um campo na aba Distribuição (usando o nome do intervalo, ex: Processo)
    'Escreve no campo correspondente na aba Registros (usando também nome de intervalo, ex: Processo_Registro).
    'Assim, mesmo que você mova ou troque colunas, a macro ainda escreve no lugar certo.
    wsRegistro.Cells(linha, wsRegistro.Range("Processo_Registro").Column).Value = Range("Processo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Exequente_Registro").Column).Value = Range("Exequente").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Tipo_Registro").Column).Value = Range("Tipo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Procurador_Registro").Column).Value = Range("Procurador").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Unidade_Registro").Column).Value = Range("Unidade").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Condenacao_Registro").Column).Value = Range("Condenação").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Valor_Registro").Column).Value = Range("Valor").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Prazo_Registro").Column).Value = Range("Prazo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Complexidade_Registro").Column).Value = Range("Complexidade").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Obs_Registro").Column).Value = Range("Obs").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Modo_Registro").Column).Value = Range("Modo").Value
    wsRegistro.Cells(linha, wsRegistro.Range("Calculista_Registro").Column).Value = Range("Calculista").Value

    'Limpar campos da aba Distribuição, exceto Modo e Calculista (comentado para teste)
    '===> Como está comentada, atualmente não limpa nada, útil para testes. <===
    'Range("Processo").Value = ""
    'Range("Exequente").Value = ""
    'Range("Tipo").Value = ""
    'Range("Procurador").Value = ""
    'Range("Unidade").Value = ""
    'Range("Condenação").Value = ""
    'Range("Valor").Value = ""
    'Range("Prazo").Value = ""
    'Range("Complexidade").Value = ""
    'Range("Obs").Value = ""

    'Exibe uma caixa de mensagem avisando que o registro foi realizado corretamente.
    MsgBox "Registro realizado com sucesso!", vbInformation, "Registro concluído"

End Sub
