# Emissor de Pesquisas NPS

![Logo da Aplicação](Designer.png)

Aplicação de desktop para Windows desenvolvida para automatizar o envio de pesquisas de Net Promoter Score (NPS) através da API da plataforma [Amplifique.me](https://amplifique.me/).

O programa lê os dados dos clientes a partir de uma planilha Excel, valida as informações, verifica se o e-mail já foi pesquisado anteriormente ou se pertence a um domínio interno, e envia os dados para a API criar e disparar a pesquisa.

**Autor:** Tadashi Suto
**Versão:** 1.0

---

## ✨ Funcionalidades

- **Interface Gráfica Amigável:** Construído com `CustomTkinter` para uma experiência de uso moderna e intuitiva.
- **Tela de Carregamento:** Exibe uma tela de splash com o logo e a versão da aplicação antes de iniciar.
- **Seleção de Arquivos:** Permite que o usuário selecione facilmente as planilhas de entrada.
- **Processamento Assíncrono:** O envio dos dados é feito em uma thread separada para que a interface não trave durante a execução.
- **Validação de Dados:**
  - Verifica se todos os campos obrigatórios estão preenchidos.
  - Valida o formato dos e-mails.
  - Impede o envio para e-mails de domínios internos (ex: `@avipam.com.br`).
  - Evita o envio duplicado, consultando uma planilha de e-mails já enviados.
- **Feedback em Tempo Real:**
  - Uma barra de progresso indica o andamento do processamento da planilha.
  - Uma área de log na tela exibe o status de cada envio (sucesso, erro, ignorado).
- **Geração de Logs:**
  - Salva um log diário em formato de texto (`.txt`) na pasta `logs`.
  - Exporta um relatório detalhado em Excel (`.xlsx`) ao final de cada execução, com o status de cada linha processada.

---

## ⚙️ Pré-requisitos

- Python 3.8 ou superior.

---

## 🚀 Instalação e Configuração

Siga os passos abaixo para configurar o ambiente e executar o projeto.

1.  **Clone o Repositório**
    Se estiver usando git, clone o repositório. Caso contrário, apenas baixe e descompacte os arquivos em uma pasta.

2.  **Crie e Ative um Ambiente Virtual (venv)**
    É uma boa prática isolar as dependências do projeto. Abra o terminal na pasta do projeto e execute:

    ```bash
    # Cria o ambiente virtual
    python -m venv venv

    # Ativa o ambiente no Windows
    .\venv\Scripts\activate
    ```

3.  **Instale as Dependências**
    Com o ambiente virtual ativo, instale todas as bibliotecas necessárias usando o arquivo `requirements.txt`:

    ```bash
    pip install -r requirements.txt
    ```

---

## ▶️ Como Usar

1.  **Execute a Aplicação**
    Com o ambiente virtual ativo, inicie o programa:
    ```bash
    python NPS.py
    ```

2.  **Preencha os Campos na Interface:**
    - **Planilha de E-mails Já Enviados:** Selecione o arquivo `.xlsx` que contém a lista de e-mails que já receberam a pesquisa. O programa usará a primeira coluna para verificação.
    - **Planilha da Pesquisa:** Selecione o arquivo `.xlsx` com os dados dos clientes a serem enviados. A planilha deve conter **pelo menos 10 colunas** na seguinte ordem: `Nome`, `Email`, `Empresa`, `ID do Cliente`, `ID da Transação`, `Unidade de Negócio`, `Empresa`, `Filial`, `Célula de Atendimento`, `VIP`.
    - **Token da Pesquisa:** Insira o token de autenticação (Bearer Token) fornecido pela API da Amplifique.me.
    - **Tempo de Expiração (dias):** Defina em quantos dias a pesquisa irá expirar após o envio. O valor padrão é 5.

3.  **Inicie o Processamento**
    - Clique no botão **"Executar"**.
    - Acompanhe o progresso na barra e os detalhes na área de log.

4.  **Verifique os Resultados**
    - Ao final, uma mensagem de conclusão será exibida.
    - Um arquivo `log_de_envio_[data_hora].xlsx` será criado na pasta do projeto com o resultado detalhado de cada linha.

---

*Copyright (c) 2025 Tadashi Suto*