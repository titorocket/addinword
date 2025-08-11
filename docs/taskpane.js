/* IE11/Word 2016 compatible JS (no ES6). */
(function () {
  "use strict";

  var InsertLocationReplace = "Replace";
  var modelName = "gpt-4o-mini";

  Office.initialize = function () {
    if (document && document.getElementById) {
      document.getElementById("btnRevisarSel").onclick = onRevisarSeleccion;
      document.getElementById("btnApplyPreview").onclick = onApplyPreview;
      document.getElementById("btnCancelPreview").onclick = closePreview;
      log("Listo. Selecciona texto en el documento y elige un nivel.");
    }
  };

  function log(msg) {
    var el = document.getElementById("log");
    if (el) {
      var ts = new Date().toLocaleTimeString();
      el.textContent += "[" + ts + "] " + msg + "\n";
      el.scrollTop = el.scrollHeight;
    }
  }

  function getApiKey() {
    var k = document.getElementById("apiKey").value;
    return (k || "").trim();
  }

  function postJson(url, headers, bodyObj, onSuccess, onError) {
    try {
      var xhr = new XMLHttpRequest();
      xhr.open("POST", url, true);
      xhr.setRequestHeader("Content-Type", "application/json");
      for (var h in headers) { if (headers.hasOwnProperty(h)) xhr.setRequestHeader(h, headers[h]); }
      xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
          if (xhr.status >= 200 && xhr.status < 300) {
            try { onSuccess(JSON.parse(xhr.responseText)); }
            catch (e) { onError(new Error("Respuesta no es JSON válido: " + e.message)); }
          } else {
            onError(new Error("HTTP " + xhr.status + ": " + xhr.responseText));
          }
        }
      };
      xhr.send(JSON.stringify(bodyObj));
    } catch (ex) {
      onError(ex);
    }
  }

  function buildSystemPrompt() {
    var sys =
      "Eres un asistente de redacción judicial penal en español (El Salvador). " +
      "Corriges y mejoras textos preservando el sentido fáctico y procesal. " +
      "Prohibido: saludos, cierres, explicaciones meta, inventar hechos o citas. " +
      "Respeta nombres, fechas, montos, números de expediente y estructura. " +
      "Estilo objetivo e impersonal; precisión terminológica penal.\n\n" +
      "Modos:\n" +
      "L1: solo ortografía, mayúsculas normativas, concordancias y puntuación; no cambies palabras ni el orden.\n" +
      "L2: L1 + mejora moderada de claridad; puedes sustituir algunas palabras por equivalentes jurídicos y ajustar conectores manteniendo el bloque argumental (±15% de longitud).\n" +
      "L3: voz de juez penal; reestructura para mayor técnica (hechos–prueba–norma–conclusión), refuerza estándares (suficiencia indiciaria, corroboración, cadena de custodia, sana crítica), sin añadir hechos ni citas inventadas.\n" +
      "Salida: solo el texto final, sin comentarios; conserva saltos de párrafo y numeraciones.";
    return sys;
  }

  function buildUserPromptSeleccion(seleccionTexto, nivel, enfoque) {
    var header = "MODO: " + nivel + "\nMATERIA: penal\n";
    if (enfoque) { header += "ENFOQUE_OPCIONAL: " + enfoque + "\n"; }
    var body = "\nTEXTO_SELECCIONADO:\n" + seleccionTexto;
    return header + body;
  }

  function callOpenAI_chat(apiKey, systemPrompt, userPrompt, onSuccess, onError) {
    var body = {
      "model": modelName,
      "messages": [
        { "role": "system", "content": systemPrompt },
        { "role": "user", "content": userPrompt }
      ],
      "temperature": 0
    };

    postJson("https://api.openai.com/v1/chat/completions",
      { "Authorization": "Bearer " + apiKey },
      body,
      function (data) {
        try {
          var content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
          if (!content) throw new Error("Sin contenido de modelo.");
          onSuccess(content);
        } catch (e) { onError(e); }
      },
      onError
    );
  }

  function openPreview(text, modeLabel) {
    var ta = document.getElementById("previewText");
    ta.value = text || "";
    document.getElementById("modalTitle").textContent = "Vista previa de cambios (" + modeLabel + ")";
    document.getElementById("backdrop").style.display = "block";
    document.getElementById("previewModal").style.display = "block";
  }
  function closePreview() {
    document.getElementById("backdrop").style.display = "none";
    document.getElementById("previewModal").style.display = "none";
  }
  function onApplyPreview() {
    var ta = document.getElementById("previewText");
    var finalText = ta.value;
    closePreview();
    applyToSelection(finalText);
  }

  function applyToSelection(text) {
    Word.run(function (ctx2) {
      var r = ctx2.document.getSelection();
      r.insertText(text, InsertLocationReplace);
      return ctx2.sync();
    }).then(function () { log("Cambios aplicados a la selección."); })
      .catch(function (e) { log("Error al aplicar: " + e.message); });
  }

  function onRevisarSeleccion() {
    var apiKey = getApiKey();
    if (!apiKey) { log("Pegue su API Key."); return; }

    var nivel = document.getElementById("nivel").value || "L1";
    var enfoque = document.getElementById("enfoque").value || "";

    Word.run(function (context) {
      var range = context.document.getSelection();
      range.load("text");
      return context.sync().then(function () {
        var t = range.text || "";
        if (!t) { log("No hay texto seleccionado."); return; }

        log("Solicitando sugerencia (" + nivel + ")...");
        var sys = buildSystemPrompt();
        var user = buildUserPromptSeleccion(t, nivel, enfoque);

        callOpenAI_chat(apiKey, sys, user, function (content) {
          if (nivel === "L1") {
            applyToSelection(content);
          } else {
            openPreview(content, nivel);
          }
        }, function (err) {
          log("Error: " + (err.message || err));
        });
      });
    }).catch(function (e) { log("Error Word.run: " + e.message); });
  }

})();