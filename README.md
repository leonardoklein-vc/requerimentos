# 🏥 Smart Requests Manager - HCPA (Gestão de Requerimentos)

![Architecture](https://img.shields.io/badge/Architecture-Serverless_SPA-blue?style=for-the-badge)
![Google Apps Script](https://img.shields.io/badge/Backend-Google_Apps_Script-4285F4?style=for-the-badge&logo=google)
![JavaScript](https://img.shields.io/badge/Frontend-Vanilla_JS-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![Security](https://img.shields.io/badge/Compliance-LGPD_Ready-10a37f?style=for-the-badge)

Uma Single Page Application (SPA) segura, resiliente e de alta performance, construída nativamente no ecossistema Google Workspace. Este sistema orquestra todo o ciclo de vida de solicitações de Prontuários Médicos, automatizando a comunicação entre Banco de Dados (Sheets), Armazenamento (Drive) e Disparos (Gmail).

---

## 🚨 O Contexto e o Desafio

A liberação de documentos médicos é uma operação crítica. O fluxo anterior dependia de planilhas estáticas compartilhadas, o que gerava gargalos operacionais e riscos de segurança:
* **Concorrência Destrutiva:** Múltiplos operadores editando as mesmas requisições simultaneamente.
* **Gestão Manual de Arquivos:** Criação de pastas, conversão de imagens e unificação de PDFs feitos de forma manual, consumindo horas da equipe.
* **Risco de Compliance (LGPD):** Dificuldade em garantir e auditar o bloqueio de dados sensíveis quando as solicitações eram feitas por terceiros (procuradores/familiares).

*(Insira a Imagem da Planilha Antiga Aqui)*
`![Planilha Antiga](./assets/planilha_antiga.png)`

---

## 💡 A Solução: Arquitetura e Engenharia

Para resolver esses problemas sem custos de infraestrutura em nuvem, o sistema foi desenhado para operar como um **Backend serverless (GAS)** alimentando um **Frontend reativo (HTML/CSS/JS)**.

*(Insira a Imagem do Novo Painel Aqui)*
`![Painel Novo](./assets/painel_novo.png)`

### 🛠️ Core Features & Soluções Técnicas

#### 1. Motor de Processamento de PDFs Resiliente (Client-Side)
O processamento de arquivos pesados foi movido para a memória do navegador do usuário (Client-Side), evitando os *timeouts* rígidos do servidor do Google.
* **Mesclagem Inteligente:** O sistema recebe múltiplos arquivos (PDFs, JPGs, PNGs), converte imagens em páginas A4 dimensionadas e gera um único volume em tempo real.
* **Sistema de Resgate (Fallback de Corrupção):** Em ambientes de saúde, sistemas geram PDFs fora do padrão que quebram bibliotecas comuns (erro `flate stream`). Desenvolvi um teste de quarentena: se a biblioteca primária (`pdf-lib`) falhar, o sistema aciona silenciosamente uma biblioteca secundária (`pdf.js`), renderiza o PDF corrompido em um `<canvas>` HTML invisível, extrai as imagens (JPEG) e reconstrói um documento limpo e seguro para o Drive.

#### 2. Workflow de Compliance e LGPD (Smart Status)
O sistema aplica regras de negócio rigorosas baseadas no requerente:
* **Próprio Paciente:** A pasta no Google Drive é provisionada com permissões de acesso via link de segurança (`DOMAIN_WITH_LINK`).
* **Terceiros (Restrito):** A pasta nasce **TRANCADA** (`DriveApp.Access.PRIVATE`). A interface do operador é alterada, bloqueando o envio e exigindo o upload de documentos legais (Procuração/RG) na subpasta "Docs". Somente após o operador marcar o *checklist* de validação documental, o backend libera as permissões na nuvem e permite a continuidade do fluxo.

#### 3. Controle de Concorrência (Soft-Locking)
Para evitar conflitos de edição (*race conditions*), o painel implementa um sistema de *Soft Lock* com Time-To-Live (TTL) de 15 minutos. 
* Quando um operador abre uma linha, o sistema sinaliza globalmente: *"Em uso por [Usuário]"*.
* Botões de salvamento são bloqueados para outros usuários na mesma linha, mantendo a integridade da fila de trabalho.

#### 4. Interpretação de Texto via Regex e Fila Assíncrona
* **Smart Parsing:** O módulo de E-mail lê requisições em texto bruto e utiliza Expressões Regulares (Regex) para identificar padrões, preenchendo automaticamente "Datas Iniciais", "Datas Finais" e "Tipos de Documento".
* **Fila de Mensageria:** Para não travar a interface do usuário durante o envio de e-mails pesados, o frontend envia um *payload* para uma Planilha-Robô separada, que processa a fila do Gmail de forma assíncrona no background (Cron Triggers).

---

## 🎨 UI/UX e Produtividade
A interface foi projetada para reduzir a fadiga cognitiva de operadores que passam 8 horas por dia no sistema:
* **Atalhos de Memória:** Botões de cópia com um clique para Prontuários, Nomes e E-mails (integração via `navigator.clipboard`).
* **Destaque Cognitivo:** Campos que exigem atenção (como pedidos feitos por terceiros) ganham destaque imediato com cores de alerta e *tooltips* explicativos.
* **Theming Dinâmico:** Implementação nativa via variáveis CSS permitindo a troca em tempo real de temas (Dark Mode, Light, Alto Contraste para acessibilidade e até temas nostálgicos como Windows XP/Aero).

*(Insira a Imagem do Modal de Validação Documental Aqui)*
`![Validação Documental](./assets/validacao_docs.png)`

---

## 🚀 Tecnologias Utilizadas
* **Backend:** Google Apps Script (V8 Engine), Google Drive API, Gmail API.
* **Frontend:** Vanilla JavaScript, HTML5, CSS3 (Variáveis Dinâmicas, CSS Grid/Flexbox).
* **Bibliotecas Client-Side:** `pdf-lib` (Manipulação binária de PDF), `pdf.js` (Renderização de fallback).
