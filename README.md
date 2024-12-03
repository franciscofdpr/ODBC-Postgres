### **Conexão ODBC com PostgreSQL no Excel**

Este documento fornece um passo a passo detalhado para configurar e utilizar uma conexão ODBC com um banco de dados PostgreSQL no Excel.

---

### **1. Baixando o Driver ODBC para PostgreSQL**

1. Acesse o site de download do driver ODBC para PostgreSQL:
   - [https://www.postgresql.org/ftp/odbc/versions.old/msi/](https://www.postgresql.org/ftp/odbc/versions.old/msi/)
   
2. Baixe o arquivo `psqlodbc_16_00_0000.zip`.

3. Solicite à equipe de TI que instale o driver em sua máquina.

---

### **2. Configurando a Conexão ODBC**

Após a instalação do driver, siga os passos abaixo para configurar a conexão:

1. No **Windows**, digite na barra de pesquisa "Fontes de Dados ODBC" e abra o aplicativo **Fontes de Dados ODBC**.

2. Na janela do **Administrador ODBC**, vá até a guia **DSN de Usuário**.

3. Clique no botão **Adicionar...**.

4. Selecione a opção **PostgreSQL ODBC Driver (ANSI)** e clique em **Concluir**.

5. Preencha os campos necessários conforme as instruções abaixo:

   - **Data Source**: `PostgreSQL30` (ou qualquer nome de sua preferência para identificar a conexão)
   - **Database**: O nome do banco de dados que deseja acessar
   - **Server**: O endereço do servidor onde o PostgreSQL está hospedado (ex: `localhost` ou IP do servidor)
   - **User Name**: Seu nome de usuário para acessar o banco
   - **Description**: (Opcional) Uma descrição para identificar a conexão
   - **SSL Mode**: `require` (para uma conexão segura)
   - **Port**: A porta do banco de dados PostgreSQL (por padrão, é a 5432)
   - **Password**: Sua senha de acesso ao banco

6. Clique em **Testar** para verificar se a conexão está funcionando corretamente.

7. Se a mensagem **Connection successful** aparecer, clique em **OK** para salvar a configuração da fonte de dados.

---

### **3. Utilizando a Conexão ODBC no Excel**

Agora que você configurou a fonte de dados ODBC, siga os passos abaixo para acessar o banco de dados PostgreSQL no Excel:

1. Abra o **Excel**.

2. Na guia **Dados**, clique em **Obter Dados**.

3. Selecione a opção **De Outras Fontes** e, em seguida, clique em **Do ODBC**.

4. Na janela que aparecer, em **DSN (Nome da Fonte de Dados)**, selecione a fonte de dados configurada (ex: `PostgreSQL30`).

5. Preencha as credenciais de acesso:
   - **Nome do usuário**: Seu nome de usuário
   - **Senha**: Sua senha de acesso ao banco

6. Clique em **Conectar**.

7. Na próxima janela, selecione o **schema** e a **tabela** que deseja consultar ou analisar.

---

### **Dicas Adicionais**

- Certifique-se de que o **driver ODBC** foi instalado corretamente e que o PostgreSQL está acessível no servidor.
- Caso ocorra algum erro de conexão, verifique as configurações do firewall e a possibilidade de bloqueio de porta.
- Para realizar consultas mais avançadas no Excel, você pode usar o **Power Query** ou outras ferramentas de análise de dados.

---
