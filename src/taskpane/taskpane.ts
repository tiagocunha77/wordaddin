/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */
let json = (body: unknown, replacer?: (key: string, value: unknown) => unknown): string => {
  return JSON.stringify(body !== undefined ? body : {}, replacer);
};

let makeRequest = (url: string, method: string = "GET", body?: BodyInit): Promise<any> => {
  let headers = new Headers();

  if (body instanceof FormData || body instanceof URLSearchParams) {
    console.info("teste", body);
  } else {
    headers.append("Content-Type", "application/json");
  }
  let token = sessionStorage.getItem("AuthToken");
  if (token) {
    headers.append("Authorization", "Bearer " + token);
  }
  console.log("teste43", url, method, headers, body);

  return fetch("https://tiagocunhapc:8443" + url, {
    method: method,
    headers: headers,
    credentials: "include",
    body: body,
  })
    .then((response) => {
      if (response.redirected && response.url.includes("/login")) {
        //console.info("router", this.router);
        // this.router.load("login-page");
        // let form = new FormData();
        // form.append("username", "tiago");
        // form.append("password", "123456");
        // return makeRequest("/login", "POST", form);
      }
      console.log("response", response);
      response.headers.forEach(console.info);

      if (url == "/login") {
        return response.text();
      }
      return response.json();
    })
    .catch((erro) => {
      console.error("deu erro", erro);
    });
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.");
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("login").onclick = login;
    document.getElementById("select-box").onfocus = selectBox;
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
    document.getElementById("insert-table").onclick = insertTable;
    document.getElementById("insert-image").onclick = insertImage;
    document.getElementById("send-comment").onclick = sendComment;
  }

  function login() {
    var data = new URLSearchParams();

    this.username = (document.getElementById("username") as HTMLInputElement).value;
    this.password = (document.getElementById("password") as HTMLInputElement).value;
    data.append("username", this.username);
    data.append("password", this.password);
    console.log("dadoslogin", this.username, this.password);
    makeRequest("/token", "POST", data).then((resp) => {
      try {
        sessionStorage.setItem("AuthToken", resp.jwt);

        console.info("token", resp);
        document.querySelector(".need-login").classList.remove("need-login");
      } catch {}
    });
  }
  function selectBox() {
    Word.run(function (context) {
      return makeRequest("/desenhos").then((desenhos) => {
        let selectbox = document.getElementById("select-box") as HTMLSelectElement;
        document.querySelectorAll("#select-box option").forEach((option) => option.remove());
        desenhos.forEach((desenho) => {
          let opt = document.createElement("option");
          opt.value = desenho.id;
          opt.text = desenho.titulo;
          selectbox.add(opt);
        });

        return context.sync();
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function sendComment() {
    Word.run(function (context) {
      var selectedRange = context.document.getSelection();
      context.load(selectedRange, "text");

      return context.sync().then(function () {
        console.log("rangess", selectedRange.text);

        let comment = {
          desenho: {
            id: 1,
          },
          texto: selectedRange.text,
        };

        makeRequest("/comentario", "POST", json(comment));
        return context.sync();
      });
    });
  }

  function insertImage() {
    Word.run(function (context) {
      return makeRequest("/desenhos").then((img) => {
        img.forEach((desenhos) => {
          let index = desenhos.desenho.indexOf("base64,");
          let base64 = desenhos.desenho.substring(index + 7);

          context.document.body.insertInlinePictureFromBase64(base64, "Start");
        });
        return context.sync();
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function insertTable() {
    Word.run(function (context) {
      var body = context.document.body;

      return makeRequest("/autor").then((tabs) => {
        let tableData = [];
        let titulos = ["Autor", "Titulo Desenho"];
        tableData.push(titulos);
        console.info("tabela", tabs);
        tabs.forEach((autor) => {
          autor.desenhos.forEach((desenho) => {
            let linha = [autor.nome, desenho.titulo];
            tableData.push(linha);
          });
        });
        console.log("tabela 222", tableData);
        body.insertTable(tableData.length, 2, "End", tableData);

        return context.sync();
      });
      // var tableData = [
      //   ["Nome", "Desenho"],
      //   ["autor1", "desenho1"],
      //   ["autor2", "desenho2"],
      // ];

      // secondParagraph.insertTable(3, 2, "After", tableData);
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function insertParagraph() {
    Word.run(function (context) {
      var docBody = context.document.body;

      return makeRequest("/autor").then((resp) => {
        console.info("resp boy", resp);
        resp.forEach((autor) => {
          let p = docBody.insertParagraph(autor.nome, "Start");

          autor.desenhos.forEach((desenho) => {
            p.insertText("\n \t" + desenho.titulo, "End");
          });
        });

        return context.sync();
      });
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function applyStyle() {
    Word.run(function (context) {
      var firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.Style.intenseReference;

      return context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function changeFont() {
    Word.run(function (context) {
      var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Courier New",
        bold: true,
        size: 18,
      });

      return context.sync();
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }

  function insertTextIntoRange() {
    Word.run(function (context) {
      var doc = context.document;
      var originalRange = doc.getSelection();
      originalRange.insertText(" (C2R)", "End");

      originalRange.load("text");
      return context
        .sync()
        .then(function () {
          doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
        })
        .then(context.sync);
    }).catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
    });
  }
});
