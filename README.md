# 📊 Planilha de Controle - UCALC

Planilha automatizada para controle e distribuição de cálculos da Unidade de Cálculos da PGE.

---

## ⚙️ Macro `RegistrarDados.vb`

A planilha possui uma macro vinculada a um botão, responsável por registrar os dados da aba **Distribuição** na aba **Registros**.

### 🔎 Funcionalidades

A macro executa as seguintes etapas:

1. **Validação de dados**  
   Verifica se todos os campos obrigatórios da aba *Distribuição* estão preenchidos.

2. **Localização da próxima linha disponível**  
   Identifica a primeira linha vazia na aba *Registros*, a partir da linha 8.

3. **Registro da data**  
   Insere automaticamente a data do registro na coluna correspondente.

4. **Transferência de dados**  
   Copia os dados da aba *Distribuição* para a aba *Registros*, utilizando **nomes de intervalo**, garantindo maior robustez.

5. **Limpeza dos campos (opcional)**  
   Limpa os campos da aba *Distribuição*, mantendo apenas os campos **Modo** e **Calculista**.

6. **Confirmação ao usuário**  
   Exibe uma mensagem indicando que o registro foi realizado com sucesso.

---

## 📌 Observações

- A macro **não depende dos cabeçalhos** da linha 6, pois utiliza nomes de intervalo.
- Continua funcionando mesmo com **alterações na ordem ou inclusão de colunas**.
- Implementa uma abordagem **robusta e profissional** para automação de registros no Excel.

---

## 🚀 Objetivo

Facilitar o controle, padronizar registros e otimizar a distribuição de cálculos dentro da unidade.

