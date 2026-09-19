import streamlit as st
import streamlit.components.v1 as components

# Configuração da página Streamlit
st.set_page_config(
    page_title="Projeto Sementes - IEQ Guaicurus",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Injeção do código HTML/CSS/JS completo
codigo_html = """
<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Projeto Sementes - IEQ Guaicurus</title>
  <style>
    :root {
      --bg-principal: #0a0a0a;
      --bg-card: #141414;
      --bg-input: #1f1f1f;
      --ouro: #d4af37;
      --ouro-brilho: rgba(212, 175, 55, 0.4);
      --texto-claro: #f5f5f5;
      --texto-escuro: #888888;
      --verde-sucesso: #2e7d32;
      --vermelho-alerta: #c62828;
    }

    * {
      margin: 0;
      padding: 0;
      box-sizing: border-box;
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      -webkit-tap-highlight-color: transparent;
    }

    body {
      background-color: var(--bg-principal);
      color: var(--texto-claro);
      padding-bottom: 80px;
      overflow-x: hidden;
    }

    .header-app {
      background-color: var(--bg-card);
      padding: 15px;
      text-align: center;
      border-bottom: 1px solid var(--ouro-brilho);
      position: sticky;
      top: 0;
      z-index: 100;
    }

    .header-app img {
      max-width: 140px;
      height: auto;
      display: block;
      margin: 0 auto;
    }

    .container {
      padding: 15px;
      max-width: 600px;
      margin: 0 auto;
    }

    .bento-card {
      background-color: var(--bg-card);
      border-radius: 16px;
      padding: 20px;
      margin-bottom: 15px;
      border: 1px solid rgba(255, 255, 255, 0.05);
      transition: transform 0.2s, border-color 0.2s;
    }

    .bento-card.destaque-ouro {
      border: 1px solid var(--ouro);
      box-shadow: 0 4px 15px var(--ouro-brilho);
    }

    .bento-card h2 {
      color: var(--ouro);
      font-size: 1.2rem;
      margin-bottom: 12px;
      display: flex;
      align-items: center;
      gap: 8px;
    }

    .texto-secundario {
      color: var(--texto-escuro);
      font-size: 0.9rem;
      line-height: 1.4;
    }

    .campo-grupo {
      margin-bottom: 15px;
    }

    .campo-grupo label {
      display: block;
      font-size: 0.85rem;
      color: var(--ouro);
      margin-bottom: 5px;
      font-weight: 600;
    }

    .input-app {
      width: 100%;
      background-color: var(--bg-input);
      border: 1px solid rgba(255, 255, 255, 0.1);
      border-radius: 8px;
      padding: 12px;
      color: var(--texto-claro);
      font-size: 1rem;
      outline: none;
    }

    .input-app:focus {
      border-color: var(--ouro);
    }

    .btn-app {
      width: 100%;
      background-color: var(--ouro);
      color: #000;
      border: none;
      border-radius: 8px;
      padding: 14px;
      font-size: 1rem;
      font-weight: bold;
      cursor: pointer;
      transition: background-color 0.2s;
    }

    .btn-app:active {
      background-color: #bfa030;
    }

    .nav-bottom {
      position: fixed;
      bottom: 0;
      left: 0;
      right: 0;
      background-color: var(--bg-card);
      border-top: 1px solid var(--ouro-brilho);
      display: flex;
      justify-content: space-around;
      padding: 10px 0;
      z-index: 100;
    }

    .nav-item {
      background: none;
      border: none;
      color: var(--texto-escuro);
      font-size: 0.75rem;
      display: flex;
      flex-direction: column;
      align-items: center;
      gap: 4px;
      cursor: pointer;
      width: 20%;
    }

    .nav-item.active {
      color: var(--ouro);
    }

    .nav-icon {
      font-size: 1.3rem;
    }

    .tab-conteudo {
      display: none;
    }

    .tab-conteudo.active {
      display: block;
    }

    #painel-assinatura {
      background-color: #ffffff;
      border-radius: 8px;
      width: 100%;
      height: 150px;
      margin-bottom: 10px;
      cursor: crosshair;
      touch-action: none;
    }

    .status-aluno {
      display: inline-block;
      width: 10px;
      height: 10px;
      border-radius: 50%;
      margin-right: 5px;
    }

    .status-regular {
      background-color: var(--verde-sucesso);
    }

    .status-pendente {
      background-color: var(--vermelho-alerta);
    }

    .aluno-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px;
      background-color: var(--bg-input);
      border-radius: 8px;
      margin-bottom: 8px;
    }

    .barra-energia {
      display: flex;
      align-items: center;
      gap: 4px;
      margin-bottom: 10px;
    }

    .gota-azeite {
      font-size: 1.2rem;
      filter: grayscale(1);
    }

    .gota-azeite.cheia {
      filter: grayscale(0);
    }
  </style>
</head>
<body>

  <header class="header-app">
    <img src="https://raw.githubusercontent.com/[USUARIO_GITHUB]/[REPOSITORIO]/main/logo_projeto_sementes.png" 
         alt="Logo Projeto Sementes" 
         onerror="this.src='https://via.placeholder.com/140x50/141414/d4af37?text=SEMENTES+JIU-JITSU'">
  </header>

  <main class="container">
    <section id="tab-matricula" class="tab-conteudo active">
      <div class="bento-card destaque-ouro">
        <h2>📝 Cadastro de Matrícula</h2>
        <p class="texto-secundario">Preencha os dados e assine na tela. A biometria facial do menor será liberada após o envio desta etapa.</p>
      </div>

      <form id="form-cadastro" onsubmit="event.preventDefault(); processarMatricula();">
        <div class="bento-card">
          <h2>🥋 Dados do Aluno</h2>
          <div class="campo-grupo">
            <label for="nome-aluno">Nome Completo do Aluno</label>
            <input type="text" id="nome-aluno" class="input-app" required>
          </div>
          <div class="campo-grupo">
            <label for="data-nasc">Data de Nascimento</label>
            <input type="date" id="data-nasc" class="input-app" required>
          </div>
          <div class="campo-grupo">
            <label for="endereco">Endereço Residencial (Corumbá-MS)</label>
            <input type="text" id="endereco" class="input-app" placeholder="Bairro, Rua, Número" required>
          </div>
        </div>

        <div class="bento-card">
          <h2>✍️ Termo & Assinatura na Tela</h2>
          <p class="texto-secundario" style="margin-bottom: 15px;">
            Eu, responsável legal, autorizo o menor a participar das aulas e dou consentimento para uso institucional de imagem conforme LGPD.
          </p>
          <canvas id="painel-assinatura"></canvas>
          <button type="button" class="btn-app" style="background-color: var(--bg-input); color: #fff; margin-bottom: 10px;" onclick="limparAssinatura()">Limpar Desenho</button>
          <button type="submit" class="btn-app">Finalizar e Enviar para o Zap</button>
        </div>
      </form>
    </section>

    <section id="tab-cadastrados" class="tab-conteudo">
      <div class="bento-card">
        <h2>👥 Alunos Registrados</h2>
        <input type="text" class="input-app" placeholder="Buscar aluno..." style="margin-bottom: 15px;">

        <div class="aluno-item">
          <div>
            <span class="status-aluno status-regular"></span>
            <strong>Gabriel Amorim</strong>
            <div class="texto-secundario" style="font-size: 0.8rem;">Faixa Cinza - 2 Graus</div>
          </div>
          <button class="btn-app" style="width: auto; padding: 6px 12px; font-size: 0.8rem;" onclick="alert('Ficha de saúde sem restrições.')">Ficha</button>
        </div>

        <div class="aluno-item">
          <div>
            <span class="status-aluno status-pendente"></span>
            <strong>Lucas Silva (Experimental)</strong>
            <div class="texto-secundario" style="font-size: 0.8rem;">Faixa Branca - 0 Graus</div>
          </div>
          <button class="btn-app" style="width: auto; padding: 6px 12px; font-size: 0.8rem; background-color: var(--vermelho-alerta); color: #fff;">Pendente</button>
        </div>
      </div>
    </section>

    <section id="tab-tatame" class="tab-conteudo">
      <div class="bento-card destaque-ouro">
        <h2>🥋 Controle do Tatame</h2>
        <p id="data-chamada" class="texto-secundario" style="color: var(--ouro); font-weight: bold;"></p>
      </div>

      <div class="bento-card">
        <h2>📸 Chamada Coletiva Inteligente</h2>
        <p class="texto-secundario" style="margin-bottom: 12px;">Bata uma foto do fim do treino. Nossa IA cruzará a imagem com a biometria para registrar presença coletiva automática.</p>
        <button class="btn-app" onclick="alert('Abrindo câmera nativa em modo Sandbox seguro...')">Bater Foto da Turma</button>
      </div>
    </section>

    <section id="tab-fe" class="tab-conteudo">
      <div class="bento-card">
        <h2>🌿 Nossas Células Familiares</h2>
        <p class="texto-secundario" style="margin-bottom: 15px;">Escolha uma de nossas células da IEQ Guaicurus e participe conosco!</p>
        
        <div class="aluno-item" style="margin-bottom: 10px;">
          <div>
            <strong>Célula Videira</strong>
            <div class="texto-secundario">Líder Marcos</div>
          </div>
          <a href="https://wa.me/5567998513404?text=Quero%20participar%20da%20Celula%20Videira" class="btn-app" style="width: auto; text-decoration: none; font-size: 0.8rem; padding: 8px 12px;">Participar</a>
        </div>
      </div>

      <div class="bento-card">
        <h2>🙏 Pedido de Oração Confidencial</h2>
        <textarea class="input-app" rows="4" placeholder="Escreva aqui seu pedido, ele será enviado diretamente ao Pastor Joel Amorim..." style="margin-bottom: 10px; resize: none;"></textarea>
        <button class="btn-app" onclick="alert('Pedido enviado!')">Enviar Pedido para o Pastor</button>
      </div>
    </section>

    <section id="tab-game" class="tab-conteudo">
      <div class="bento-card destaque-ouro">
        <h2>🏆 Jornada da Sabedoria</h2>
        <div class="barra-energia">
          <span class="gota-azeite cheia">🫒</span>
          <span class="gota-azeite cheia">🫒</span>
          <span class="gota-azeite cheia">🫒</span>
          <span class="gota-azeite">🫒</span>
          <span class="gota-azeite">🫒</span>
          <span style="font-size: 0.8rem; margin-left: 10px; color: var(--texto-escuro);">3/5 Azeites</span>
        </div>
        <p class="texto-secundario">Consuma gotas para responder aos Quizzes Históricos das Eras Bíblicas e ganhar insígnias de ouro.</p>
      </div>

      <div class="bento-card">
        <h2>📖 Fase Ativa: Era 1 - Origens</h2>
        <p style="margin-bottom: 15px; font-size: 0.95rem;"><strong>Pergunta:</strong> O que Deus criou no primeiro dia, de acordo com Gênesis 1?</p>
        
        <button class="btn-app" style="background-color: var(--bg-input); color: #fff; text-align: left; margin-bottom: 8px; font-weight: normal;" onclick="alert('Incorreto!')">A) O Sol e a Lua</button>
        <button class="btn-app" style="background-color: var(--bg-input); color: #fff; text-align: left; margin-bottom: 8px; font-weight: normal;" onclick="alert('Incorreto!')">B) Os animais marinhos</button>
        <button class="btn-app" style="background-color: var(--bg-input); color: #fff; text-align: left; margin-bottom: 8px; font-weight: normal;" onclick="alert('Resposta Exata! +10 XP.')">C) A luz e a separação das trevas</button>
      </div>
    </section>
  </main>

  <nav class="nav-bottom">
    <button class="nav-item active" onclick="trocarTab('tab-matricula', this)">
      <span class="nav-icon">📝</span>
      <span>Matrícula</span>
    </button>
    <button class="nav-item" onclick="trocarTab('tab-cadastrados', this)">
      <span class="nav-icon">👥</span>
      <span>Alunos</span>
    </button>
    <button class="nav-item" onclick="trocarTab('tab-tatame', this)">
      <span class="nav-icon">🥋</span>
      <span>Tatame</span>
    </button>
    <button class="nav-item" onclick="trocarTab('tab-fe', this)">
      <span class="nav-icon">🙏</span>
      <span>Fé</span>
    </button>
    <button class="nav-item" onclick="trocarTab('tab-game', this)">
      <span class="nav-icon">🎮</span>
      <span>Game</span>
    </button>
  </nav>

  <script>
    function trocarTab(tabId, elemento) {
      document.querySelectorAll('.tab-conteudo').forEach(tab => tab.classList.remove('active'));
      document.querySelectorAll('.nav-item').forEach(item => item.classList.remove('active'));
      document.getElementById(tabId).classList.add('active');
      elemento.classList.add('active');
    }

    const canvas = document.getElementById('painel-assinatura');
    const ctx = canvas.getContext('2d');
    let desenhando = false;

    function redimensionarCanvas() {
      const rect = canvas.getBoundingClientRect();
      canvas.width = rect.width;
      canvas.height = rect.height || 150;
      ctx.strokeStyle = "#000000";
      ctx.lineWidth = 3;
      ctx.lineCap = "round";
    }

    window.addEventListener('load', redimensionarCanvas);

    function obterCoordenadas(e) {
      const rect = canvas.getBoundingClientRect();
      const clientX = e.touches ? e.touches[0].clientX : e.clientX;
      const clientY = e.touches ? e.touches[0].clientY : e.clientY;
      return { x: clientX - rect.left, y: clientY - rect.top };
    }

    function iniciarDesenho(e) {
      desenhando = true;
      const pos = obterCoordenadas(e);
      ctx.beginPath();
      ctx.moveTo(pos.x, pos.y);
      if (e.cancelable) e.preventDefault();
    }

    function desenhar(e) {
      if (!desenhando) return;
      const pos = obterCoordenadas(e);
      ctx.lineTo(pos.x, pos.y);
      ctx.stroke();
      if (e.cancelable) e.preventDefault();
    }

    function pararDesenho() {
      desenhando = false;
    }

    canvas.addEventListener('mousedown', iniciarDesenho);
    canvas.addEventListener('mousemove', desenhar);
    canvas.addEventListener('mouseup', pararDesenho);
    canvas.addEventListener('mouseleave', pararDesenho);

    canvas.addEventListener('touchstart', iniciarDesenho, { passive: false });
    canvas.addEventListener('touchmove', desenhar, { passive: false });
    canvas.addEventListener('touchend', pararDesenho);

    function limparAssinatura() {
      ctx.clearRect(0, 0, canvas.width, canvas.height);
    }

    document.getElementById('data-chamada').innerText = new Date().toLocaleDateString('pt-BR', {
      day: '2-digit',
      month: 'long',
      year: 'numeric'
    });

    function processarMatricula() {
      alert('Enviando via API para: (67) 99851-3404.');
      limparAssinatura();
      document.getElementById('form-cadastro').reset();
    }
  </script>
</body>
</html>
"""

# Renderização do componente HTML dentro do Streamlit
components.html(codigo_html, height=800, scrolling=True)
