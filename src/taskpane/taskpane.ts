/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

/* global document, Office, Word */

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
    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("apply-style").onclick = applyStyle;
    document.getElementById("change-font").onclick = changeFont;
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;
  }

  function login() {
    var data = new URLSearchParams();
    data.append("username", "tiago");
    data.append("password", "123456");
    makeRequest("/token", "POST", data).then((resp) => {
      try {
        sessionStorage.setItem("AuthToken", resp.jwt);

        console.info("token", resp);
        document.querySelector(".need-login").classList.remove("need-login");
      } catch {}
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
