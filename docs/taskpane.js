/* IE11/Word 2016 compatible JS (no ES6). */
(function () {
  "use strict";

  var InsertLocationReplace = "Replace";
  var InsertLocationEnd = "End";
  var modelName = "gpt-4o-mini";

  Office.initialize = function () {
    if (document && document.getElementById) {
      document.getElementById("btnRevisarSel").onclick = onRevisarSeleccion;
      document.getElementById("btnApplyPreview").onclick = onApplyPreviewReplace;
      document.getElementById("btnInsertBelow").onclick = onApplyPreviewInsert;
      document.getElementById("btnCancelPreview").onclick = closePreview;
      document.getElementById("nivel").onchange = onNivelChange;
      onNivelChange(); // inicializa visibilidad de Idea central
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

  function onNivelChange() {
    var nivel = document.getElementById("nivel").value;
    var row = document.getElementById("ideaRow");
    // Mostrar "Idea central" solo para L4
    row.style.display = (nivel === "L4") ? "block" : "none";
  }

  function getApiKey() {
    var k = document.getElementById("apiKey").value;
    return (k || "").trim();
  }

  // XHR (evita fetch para IE11)
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

  // SYSTEM fijo (penal) con 4 niveles
  function buildSystemPrompt() {
    var sys =
      "Eres un asistente de redacción judicial penal en español (El Salvador). " +
      "Corriges y mejoras textos preservando el sentido fáctico y procesal. " +
      "Prohibido: saludos, cierres, explicaciones meta; inventar hechos; inventar citas. " +
      "Respeta nombres, fechas, montos, números de expediente y estructura. " +
      "Estilo objetivo e impersonal; precisión terminológica penal.\n\n" +
      "Modos:\n" +
      "L1_CORRECCIÓN_MÍNIMA: solo ortografía, mayúsculas normativas, concordancias y puntuación; no cambies palabras ni el orden.\n" +
      "L2_MEJORA_MODERADA: L1 + claridad; sustituye algunas palabras por equivalentes jurídicos y ajusta conectores manteniendo el bloque argumental (±15%).\n" +
      "L3_TÉCNICA_PENAL_REFORZADA: voz técnica (no juez); reordena levemente para coherencia (hechos–prueba–norma–conclusión), refuerza estándares (suficiencia indiciaria, corroboración, cadena de custodia, sana crítica), sin añadir hechos (±30%).\n" +
      "L4_JUEZ_PENAL_FUNDAMENTA: simula ser juez penal y elabora fundamentación a partir de IDEA_CENTRAL (si se provee) o de la selección; estructura sugerida: hechos relevantes – problema jurídico – calificación típica – valoración probatoria – derecho aplicable – conclusión. Sin inventar hechos ni citas.\n" +
      "Salida: solo el texto final, sin comentarios; conserva saltos de párrafo y numeraciones.";
    return sys;
  }

  // USER por selección + enfoque + idea (solo L4 la usa si viene)
  function buildUserPromptSeleccion(seleccionTexto, nivel, enfoque, ideaCentral) {
    var header = "MODO: " + nivel + "\nMATERIA: penal\n";
    if (enfoque) { header += "ENFOQUE_OPCIONAL: " + enfoque + "\n"; }
    if (nivel === "L4") {
      header += "IDEA_CENTRAL: " + (ideaCentral || "") + "\n";
    }
    var body = "\nTEXTO_SELECCIONADO:\n" + (seleccionTexto || "");
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

  // Modal
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
  function onApplyPreviewReplace() {
    var ta = document.getElementById("previewText");
    var finalText = ta.value;
    closePreview();
    applyToSelection(finalText, true);
  }
  function onApplyPreviewInsert() {
    var ta = document.getElementById("previewText");
    var finalText = ta.value;
    closePreview();
    applyToSelection(finalText, false);
  }

  // Aplicar a selección: replace o insertar debajo (útil en L4)
  function applyToSelection(text, replace) {
    Word.run(function (ctx2) {
      var r = ctx2.document.getSelection();
      if (replace) {
        r.insertText(text, InsertLocationReplace);
      } else {
        // Inserta un salto de párrafo y luego el texto sugerido
        r.insertText("\r" + text, InsertLocationEnd);
      }
      return ctx2.sync();
    }).then(function () { log("Cambios aplicados a la selección."); })
      .catch(function (e) { log("Error al aplicar: " + e.message); });
  }

  // Flujo principal: solo selección (L1 aplica directo; L2/L3/L4 vista previa)
  function onRevisarSeleccion() {
    var apiKey = getApiKey();
    if (!apiKey) { log("Pegue su API Key."); return; }

    var nivel = document.getElementById("nivel").value || "L1";
    var enfoque = document.getElementById("enfoque").value || "";
    var ideaCentral = document.getElementById("ideaCentral").value || "";

    Word.run(function (context) {
      var range = context.document.getSelection();
      range.load("text");
      return context.sync().then(function () {
        var t = range.text || "";
        if (!t && nivel !== "L4") { log("No hay texto seleccionado."); return; }
        // En L4 se permite que no haya selección si hay Idea central
        if (!t && nivel === "L4" && !ideaCentral) {
          log("L4: provea selección o una 'Idea central'.");
          return;
        }

        log("Solicitando sugerencia (" + nivel + ")...");
        var sys = buildSystemPrompt();
        var user = buildUserPromptSeleccion(t, nivel, enfoque, ideaCentral);

        callOpenAI_chat(apiKey, sys, user, function (content) {
          if (nivel === "L1") {
            applyToSelection(content, true);
          } else {
            openPreview(content, nivel);
          }
        }, function (err) { log("Error: " + (err.message || err)); });
      });
    }).catch(function (e) { log("Error Word.run: " + e.message); });
  }

})();
