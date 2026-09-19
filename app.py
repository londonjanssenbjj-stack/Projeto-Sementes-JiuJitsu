<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Projeto Sementes - IEQ Guaicurus</title>
  <style>
    /* ==========================================================================
       1. ARQUITETURA DE DESIGN PALETA PRETO & DOURADO (MODO ESCURO PREMIUM)
       ========================================================================== */
    :root {
      --bg-principal: #0a0a0a;   /* Preto profundo para economia de bateria */
      --bg-card: #141414;        /* Cinza escuro para blocos do Bento Grid */
      --bg-input: #1f1f1f;       /* Fundo dos campos de texto */
      --ouro: #d4af37;           /* Dourado clássico de alta qualidade */
      --ouro-brilho: rgba(212, 175, 55, 0.4);
      --texto-claro: #f5f5f5;    /* Branco fosco para leitura confortável */
      --texto-escuro: #888888;   /* Cinza para textos secundários */
      --verde-sucesso: #2e7d32;  /* Status: Cadastro Completo */
      --vermelho-alerta: #c62828; /* Status: Pendente / Alertas */
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
      padding-bottom: 80px; /* Espaço para a barra de navegação inferior */
      overflow-x: hidden;
    }

    /* ==========================================================================
       2. COMPONENTES VISUAIS COMUNS & BENTO GRID
       ========================================================================== */
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

    /* Formulários e Botões */
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

    /* ==========================================================================
       3. NAVEGAÇÃO INFERIOR (TABS)
       ========================================================================== */
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

    /* ==========================================================================
       4. CUSTOMIZAÇÕES ESPECÍFICAS DE TELAS
       ========================================================================== */
    /* Canvas de Assinatura */
    #painel-assinatura {
      background-color: #ffffff;
      border-radius: 8px;
      width: 100%;
      height: 150px;
      margin-bottom: 10px;
      cursor: crosshair;
      touch-action: none;
    }

    /* Status Pills (Bolinhas LGPD) */
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

    /* Lista de Alunos */
    .aluno-item {
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 12px;
      background-color: var(--bg-input);
      border-radius: 8px;
      margin-bottom: 8px;
    }

    /* Game Gospel Interface */
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

  <!-- CABEÇALHO GLOBAL COM A LOGO -->
  <header class="header-app">
    <img src="https://raw.githubusercontent.com/[USUARIO_GITHUB]/[REPOSITORIO]/main/logo_projeto_sementes.png" 
         alt="Logo Projeto Sementes" 
         onerror="this.src='https://via.placeholder.com/140x50/141414/d4af37?text=SEMENTES+JIU-JITSU'">
  </header>

  <main class="container">

    <!-- ======================================================================
         ABA 1: MATRÍCULA (INSCRIÇÃO E TERMOS LEGAIS)
         ====================================================================== -->
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

    
