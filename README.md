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
   Identifica a primeira linha vazia na aba *Registros*, a partir da linha **7**.

3. **Registro da data**  
   Insere automaticamente a data do registro na coluna correspondente, com formatação padrão `dd/mm/yyyy`.

4. **Transferência de dados**  
   Copia os dados da aba *Distribuição* para a aba *Registros*, utilizando **nomes de intervalo**, garantindo maior robustez.

5. **Tratamento do campo Calculista**  
   Armazena apenas o **primeiro nome** do calculista, utilizando `Trim()` e `Split()` para padronização.

6. **Limpeza dos campos (opcional)**  
   Possui estrutura para limpeza dos campos da aba *Distribuição*, mantendo **Modo** e **Calculista** (atualmente desativada para testes).

7. **Confirmação ao usuário**  
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
   Percorre a aba *Registros* a partir da linha **7**, localizando o processo que:
   - Coincide com o número informado na aba *Devolução*
   - Ainda **não possui data de entrega** preenchida

3. **Validação do processo**  
   Caso o processo não seja encontrado ou já possua data de entrega preenchida, exibe uma mensagem de erro e encerra a execução.

4. **Transferência de dados**  
   Copia os dados da aba *Devolução* para a **mesma linha** onde o processo foi encontrado na aba *Registros*, utilizando **nomes de intervalo**.

5. **Registro direcionado**  
   A atualização ocorre diretamente na linha do processo, evitando duplicidade de registros e garantindo consistência dos dados.

6. **Confirmação ao usuário**  
   Exibe uma mensagem indicando que a devolução foi registrada com sucesso, incluindo:
   - Índice do processo  
   - Nome do exequente  

---

## 📌 Observações

- As macros **não dependem dos cabeçalhos fixos**, pois utilizam nomes de intervalo.
- Continuam funcionando mesmo com **alterações na ordem ou inclusão de colunas**.
- A busca por processos considera apenas registros **sem data de entrega**, evitando sobrescritas indevidas.
- O uso de `Trim()` evita falhas por **espaços acidentais** nos dados.
- O tratamento do campo **Calculista** padroniza a informação armazenada (primeiro nome).
- Implementam uma abordagem **robusta e profissional** para automação de registros no Excel.

---

## 🚀 Objetivo

Facilitar o controle, padronizar registros e otimizar a distribuição de cálculos dentro da unidade.

---

## 🎨 Créditos

Os ícones utilizados foram obtidos em:  
[Icons8](https://icons8.com.br/icons/set/relat%C3%B3rio--size-small--white)