//V1.1

(function () {
  Office.onReady(function () {
    const mainButtons = document.getElementsByClassName("main-btn");
    for (let i = 0; i < mainButtons.length; i++) {
      mainButtons[i].addEventListener("click", function () {
        const subMenu = this.nextElementSibling;
        subMenu.style.display = subMenu.style.display === "block" ? "none" : "block";
      });
    }

    const allSubButtons = document.getElementsByClassName("btn");
    for (let i = 0; i < allSubButtons.length; i++) {
      allSubButtons[i].addEventListener("click", handleButtonClick);
    }
    document.querySelectorAll(".main-btn").forEach((btn) => {
      btn.addEventListener("click", function () {
        const group = this.parentElement;
        group.classList.toggle("active");
      });
    });

    document.querySelectorAll(".subgroup-title").forEach((title) => {
      title.addEventListener("click", function () {
        const subgroup = this.parentElement;
        subgroup.classList.toggle("active");
      });
    });
  });

  async function handleButtonClick(event) {
    const action = event.target.getAttribute("data-action");
    let text = "";
    let color = "";

    switch (action) {
      case "pagina":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const p1 = range.insertParagraph("<página X>", Word.InsertLocation.before);
          p1.style = "arco_RECADO_ART";
          await context.sync();
        });
        return;

      case "paginaxx":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const p1 = range.insertParagraph("<página X e X>", Word.InsertLocation.before);
          p1.style = "arco_RECADO_ART";
          await context.sync();
        });
        return;

      case "linha":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const p1 = range.insertParagraph("<linha>", Word.InsertLocation.before);
          p1.style = "arco_RECADO_ART";
          await context.sync();
        });
        return;

      case "titulo":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<título>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "Al.correta":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<alternativa correta>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "cotas":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<abre cotas>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_IMAGEM_COTA_TEXTO";
          const pAfter = range.insertParagraph("<fecha cotas>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "coluna_f":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<abre coluna falsa>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha coluna falsa>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "lacuna":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          range.load("text");
          await context.sync();
          const abreLacuna = range.insertText("<abre lacuna>", Word.InsertLocation.before);
          abreLacuna.font.color = "#EE0000";
          abreLacuna.font.size = 10;
          const fechaLacuna = range.insertText("<fecha lacuna>", Word.InsertLocation.after);
          fechaLacuna.font.color = "#EE0000";
          fechaLacuna.font.size = 10;
          range.font.color = "#00B0F0";
          range.font.size = 10;

          await context.sync();
        });
        return;

      case "destaque":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const abre = range.insertText("<abre destaque>", Word.InsertLocation.before);
          abre.font.color = "#EE0000";
          abre.font.size = 10;
          const fecha = range.insertText("<fecha destaque>", Word.InsertLocation.after);
          fecha.font.color = "#EE0000";
          fecha.font.size = 10;
          await context.sync();
        });
        return;

      case "destaque.glos":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const abre = range.insertText("<abre destaque glossário>", Word.InsertLocation.before);
          abre.font.color = "#EE0000";
          abre.font.size = 10;
          const fecha = range.insertText("<fecha destaque glossário>", Word.InsertLocation.after);
          fecha.font.color = "#EE0000";
          fecha.font.size = 10;
          await context.sync();
        });
        return;

      case "destaque.hip":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const abre = range.insertText("<abre destaque hiperlink>", Word.InsertLocation.before);
          abre.font.color = "#EE0000";
          abre.font.size = 10;
          const fecha = range.insertText("<fecha destaque hiperlink>", Word.InsertLocation.after);
          fecha.font.color = "#EE0000";
          fecha.font.size = 10;
          await context.sync();
        });
        return;

      // TAGs SAE
      case "agora_voce":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção agora você já sabe>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conectando":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conectando os pontos>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conversa_vai":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conversa vai>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conversa_vem":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conversa vem>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "glossario":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção glossário>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "organizando":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção organizando o conhecimento>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "para_saber":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção para saber mais>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "saberes":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção saberes em ação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "testando":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção testando as ideias>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      // Icones SAE
      case "audio_sae":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone áudio>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "caderno":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone caderno>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "oralidade":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone oralidade>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      // Boxes SAE
      case "conjunto":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box conjunto>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "situacao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box situação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "observacao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box observação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "procedimento":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box procedimento>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "resolucao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box resolução>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "conceito":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box sistematização/conceito>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "trecho":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box trecho>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "exemplo":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box exemplo>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      // Seções SPE EM
      case "atividades":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção atividades>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conexoes":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conexões>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "foco":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção foco na aprovação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "gabarito":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção gabarito>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "organize":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção organize as ideias>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "objetivos":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção objetivos do capítulo>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "coisas_da_gente":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção coisas da gente>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "fazendo_ciencia":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção fazendo ciência>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "reflexao_em_acao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção reflexão em ação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "geografia_em_foco":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção geografia em foco>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "interpretando":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção interpretando documentos>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "voce_faz_historia":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção você faz história>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "oficina":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção oficina de texto>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "em_teste":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção matemática em teste>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "em_acao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção matemática em ação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "em_detalhes":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção matemática em detalhes>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "inicio_de_conversa":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção para início de conversa>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      // Seções SPE AF
      case "proposta_1":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 1>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_2":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 2>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_3":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 3>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_4":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 4>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_5":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 5>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_6":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 6>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_7":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 7>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "proposta_8":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<proposta 8>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      // Icones SPE
      case "audio_spe":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone áudio>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "atividades_facil":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone atividades fácil>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "atividades_medio":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone atividades médio>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "atividades_dificil":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone atividades difícil>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      // Boxes SPE
      case "citacao":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box citação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_CITACAO";

          await context.sync();
        });
        return;

      case "box_destaque":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box destaque>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "box_glossario":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box glossário>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "geral":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box geral>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      // Seções PIA
      case "atividades_PIA":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção atividades>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "construindo_PIA":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção construindo ideias>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "acao_PIA":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção em ação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "fazer_PIA":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção fazer e aprender>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "pesquisa":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção pesquisa>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      // Icones PIA
      case "interessante":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone interessante>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      // Seções CQT
      case "atividades_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção atividades>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conectado_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conectado>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conquistar":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção conquistar-se>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "curiosidade_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção curiosidade>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "emocoes":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção emoções em pauta>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "ideias_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção ideias em ação>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "conquistei":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção o que já conquistei>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "pesquisa_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção pesquisa>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "saiba_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção saiba mais>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "troca":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção troca de ideias>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "activities":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção activities>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "around":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção around the world>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "goals":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção goals>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "action":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção in action>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "know":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção know yourself>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "language":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção language in use>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "listen":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção lets listen>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "see":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção lets see it again>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "start":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção lets start>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "look":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção look it up>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "reading":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção reading>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "sing":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção sing along>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "speaking":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção speaking>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "thinking":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção thinking together>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "going_to_see":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção what you are going to see>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "writing":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção writing>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "brasileira":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção arte brasileira>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "festa":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção em festa>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "galeria":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção galeria>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "pratica_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção ciência em prática>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "cartografica":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção leitura cartográfica>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "geografico":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção olhar geográfico>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "documentos_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção interpretando documentos>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "historias_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção outras histórias>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "escreve_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção como se escreve>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "lingua_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção estudo da língua>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "texto_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção estudo do texto>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "texto_escrito":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção produção de texto escrito>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "texto_oral":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção produção de texto oral>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "problemas_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção criar e resolver problemas>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "olhar":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção ampliando o olhar>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "partida":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção ponto de partida>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      case "final":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<seção reflexão final>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          const pAfter = range.insertParagraph("<fecha seção>", Word.InsertLocation.after);
          pAfter.style = "arco_RECADO_ART";

          await context.sync();
        });
        return;

      // Icones CQT
      case "audio_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone áudio>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "caderno_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone caderno>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "calculadora_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone calculadora>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "ciencias_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone ciências>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "cuidado":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone cuidado>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "desafio_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone desafio>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "grupo_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone grupo>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "material_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone material extra>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      case "voz_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const inserted = range.insertText("<ícone voz>", Word.InsertLocation.before);
          inserted.font.color = "#EE0000";
          inserted.font.size = 10;
          await context.sync();
        });
        return;

      // Boxes CQT
      case "destaque_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box destaque>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "exercicio_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box exercício>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "geral_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box geral>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "glossario_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box glossário>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "hiperlink_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box hiperlink>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "internet_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box internet>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "jornal_CQT":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box jornal>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;

      case "explicativo":
        await Word.run(async (context) => {
          const range = context.document.getSelection();
          const pBefore = range.insertParagraph("<box texto explicativo>", Word.InsertLocation.before);
          pBefore.style = "arco_RECADO_ART";
          range.style = "arco_BOX_GERAL_TEXTO";

          await context.sync();
        });
        return;
    }

    // Para os demais casos, mantém o comportamento padrão
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      const paragraph = range.insertParagraph(text, Word.InsertLocation.after);
      paragraph.style = "arco_TEXTO";
      paragraph.font.color = color;
      await context.sync();
    });
  }
})();
