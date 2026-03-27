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

## ⚙️ Macro `RegistrarDevolucao.vb`

Macro vinculada a um botão na aba **Devolução**, responsável por registrar os dados de devolução na linha correspondente ao processo na aba **Registros**.

### 🔎 Funcionalidades

A macro executa as seguintes etapas:

1. **Validação de dados**  
   Verifica se todos os campos obrigatórios da aba *Devolução* estão preenchidos.  
   Utiliza `Trim()` para garantir que campos com apenas espaços sejam tratados como vazios.

2. **Busca do processo**  
   Percorre a aba *Registros* a partir da linha 8, localizando o processo que:
   - Coincide com o número informado na aba *Devolução*
   - Ainda **não possui data de entrega** preenchida

3. **Validação do processo**  
   Caso o processo não seja encontrado ou já possua data de entrega preenchida, exibe uma mensagem de erro e encerra a execução.

4. **Transferência de dados**  
   Copia os dados da aba *Devolução* para a **linha exata** onde o processo foi encontrado na aba *Registros*, utilizando **nomes de intervalo**.

5. **Confirmação ao usuário**  
   Exibe uma mensagem indicando que a devolução foi registrada com sucesso.

---

## 📌 Observações

- As macros **não dependem dos cabeçalhos** da linha 6, pois utilizam nomes de intervalo.
- Continuam funcionando mesmo com **alterações na ordem ou inclusão de colunas**.
- Implementam uma abordagem **robusta e profissional** para automação de registros no Excel.
- O `Trim()` aplicado em `RegistrarDevolucao` evita falhas por **espaços acidentais** no início ou fim dos valores comparados e gravados.

---

## 🚀 Objetivo

Facilitar o controle, padronizar registros e otimizar a distribuição de cálculos dentro da unidade.