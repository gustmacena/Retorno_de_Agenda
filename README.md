# 🚀 Automação de Retorno Atacadão  

## 📌 Descrição  

Este projeto automatiza o processo de login e extração de dados na plataforma **HodieBooking do Atacadão**. Utilizando **Selenium** para automação do navegador e **Tkinter** para uma interface gráfica intuitiva, o programa busca o **status e protocolo das cargas** e os atualiza automaticamente em uma **planilha Excel**.  

O objetivo é **economizar tempo**, reduzir **erros manuais** e otimizar o fluxo de trabalho da equipe de **Customer Service na Bauducco**.  

---

## ⚙️ Funcionalidades  

✅ **Automação de Login** na plataforma HodieBooking  
✅ **Consulta e Extração** de status e protocolos das cargas  
✅ **Atualização Dinâmica** da planilha Excel com os dados coletados  
✅ **Interface Amigável** com Tkinter para fácil interação  
✅ **Execução Automatizada** sem necessidade de intervenção manual  

---

## 🛠️ Como Funciona  

1️⃣ O programa inicia o navegador e faz login na plataforma HodieBooking.  
2️⃣ Lê os **números de carga** a partir de uma planilha Excel.  
3️⃣ Para cada carga, navega na plataforma e coleta **status e protocolo**.  
4️⃣ Atualiza automaticamente os **dados extraídos** na planilha.  
5️⃣ Exibe **feedback em tempo real** sobre o progresso da automação.  

---

## 📌 Requisitos  

🔹 **Python 3.x** (apenas para gerar o executável)
🔹 **Bibliotecas Necessárias para compilar:**
   ```sh
   pip install -r requirements.txt
   ```
🔹 **Navegador compatível (Chrome, Edge, etc.)**

---

## 🚧 Como gerar o executável

1. Instale as dependências indicadas em `requirements.txt`.
2. Execute o script `build.py`:
   ```sh
   python build.py
   ```
3. O executável e os arquivos necessários serão criados na pasta `dist/`.
   Basta copiar esse diretório para outra máquina e executar `RetornoAtacadao.exe`.

---

## 📂 Estrutura do Projeto  

📁 `retornos_atacadao.py` → Código principal da automação
📁 `imgs/` → Diretório com imagens usadas na interface gráfica
📁 `Base Retorno Atacadão.xlsx` → Planilha Excel utilizada para leitura e atualização dos dados
📁 `build.py` → Script para gerar o executável

---

## 🎯 Benefícios  

🚀 **Automação Total**: Elimina tarefas repetitivas e manuais.  
⚡ **Eficiência**: Acelera a coleta e organização de informações.  
📊 **Precisão**: Reduz erros humanos no preenchimento da planilha.  
🖥️ **Interface Intuitiva**: Permite iniciar a automação com um clique.  

---

## 👤 Desenvolvedor  

👨‍💻 **Gustavo Macena**  
🔹 Equipe: **Projetos de Melhoria Contínua - Bauducco**  

📩 [LinkedIn](https://www.linkedin.com/in/gustavo-macena/) | 📧 gustavomacena@email.com  
