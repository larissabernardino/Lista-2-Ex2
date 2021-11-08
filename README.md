# Lista-2-Ex2
descrição : > -
  Layout do suplemento: uma caixa de entrada de seleção de cores e um botão.
  Comportamento do suplemento: ao clicar no botão, garantir que há uma cor
  atualizado na caixa de entrada. Em caso de sucesso, localizar o conjunto de
  células atualmente selecionadas na planilha do Excel ativa e colorir o fundo
  essas células com a cor definida (ou uma cor mais próxima aceita pelo
  Excel, já que nem toda a escala de núcleos RGB existente é aceita).
  Tempo estimado para conclusão: 40 minutos.
anfitrião : EXCEL
api_set : {}
script :
  conteúdo : |
    $ ("# show"). click (() => tryCatch (run));
    função assíncrona run () {
      aguardar Excel.run (assíncrono (contexto) => {
        intervalo const = context.workbook.getSelectedRange ();
        range.format.fill.color = optcolor.value;
        range.load ("endereço");
        aguarde context.sync ();
      });
    }
    função assíncrona tryCatch (callback) {
      Experimente {
        aguardar retorno de chamada ();
      } catch (erro) {
        console.error (erro);
      }
    }
  linguagem : texto datilografado
modelo :
  conteúdo : " <div> \ n \ t <select id = \" optcolor \ " name = \" select \ " > \ n     <option value = \" black \ " selected> Preto </option> \ n     <valor da opção = \ " red \" > Vermelho </option> \ n     <option value = \ " blue \" > Azul </option> \ n     <option value = \ " green \" > Verde </option> \ n   </ selecione>\ n </div> \ n <div>\ n \ t <button id = \ " show \" > Colorir Seleção </button> \ n </div> "
  linguagem : html
estilo :
  conteúdo : | -
    section.samples {
        margem superior: 20px;
    }
    section.samples .ms-Button, section.setup .ms-Button {
        display: bloco;
        margin-bottom: 5px;
        margem esquerda: 20px;
        largura mínima: 80px;
    }
  idioma : css
bibliotecas : |
  https://appsforoffice.microsoft.com/lib/1/hosted/office.js
  @ types / office-js
  office-ui-fabric-js@1.4.0/dist/css/fabric.min.css
  office-ui-fabric-js@1.4.0/dist/css/fabric.components.min.css
  core-js@2.4.1/client/core.min.js
  @ types / core-js
  jquery@3.1.1
  @ types / jquery @ 3.3.1
