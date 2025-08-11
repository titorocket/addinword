/* IE11/Word 2016 compatible JS (no ES6). */
(function () {
  "use strict";

  var InsertLocationReplace = "Replace";
  var modelName = "gpt-4o-mini";

  Office.initialize = function () {
    if (document && document.getElementById) {
      document.getElementById("btnRevisarDoc").onclick = onRevisarDocumentoCompleto;
      document.getElementById("btnRevisarSel").onclick = onRevisarSeleccion;
      log("Listo. Abra un documento y use uno de los botones.");
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
      for (var h in headers) {
        if (headers.hasOwnProperty(h)) {
          xhr.setRequestHeader(h, headers[h]);
        }
      }
      xhr.onreadystatechange = function () {
        if (xhr.readyState === 4) {
          if (xhr.status >= 200 && xhr.status < 300) {
            try {
              var data = JSON.parse(xhr.responseText);
              onSuccess(data);
            } catch (e) {
              onError(new Error("Respuesta no es JSON válido: " + e.message));
            }
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
      "Eres un asistente especializado en redacción judicial en español. " +
      "Corrige ortografía, gramática, sintaxis y estilo con técnica jurídica, " +
      "manteniendo tono formal y objetivo propio de resoluciones judiciales de El Salvador." +
      "\n\nREGLAS:" +
      "\n1) No agregues saludos, cierres, sugerencias ni comentarios meta. Solo el texto solicitado." +
      "\n2) No inventes citas ni jurisprudencia. Si hay citas existentes, respétalas; solo mejora puntuación y formato." +
      "\n3) Respeta nombres, fechas, montos y números de expediente." +
      "\n4) Mantén el sentido jurídico. Mejora coherencia, cohesión y precisión terminológica." +
      "\n5) No cambies la estructura numerada ni referencias internas." +
      "\n6) Puntuación y formato sobrios; evita oraciones excesivamente largas; usa conectores jurídicos típicos." +
      "\n7) Neutralidad y objetividad; evita adjetivación innecesaria.";
    return sys;
  }

  function buildUserPromptDocumento(parrafosArr, contexto) {
    var payload = {
      "parrafos": parrafosArr,
      "contexto_documento": contexto
    };

    var user =
      "TAREA: REVISAR_DOCUMENTO_COMPLETO\n\n" +
      "DEVOLUCIÓN ESTRICTA: JSON con la forma:\n" +
      '{\n  "version": "1.0",\n  "parrafos": [\n' +
      '    {\"index\": 0, \"texto\": \"<párrafo 0 corregido>\"},\n' +
      '    {\"index\": 1, \"texto\": \"<párrafo 1 corregido>\"}\n' +
      "  ]\n}\n\n" +
      "REGLAS DE SALIDA:\n" +
      "- Mismo número de párrafos y mismos índices.\n" +
      "- No agregar ni eliminar párrafos.\n" +
      "- No incluir comentarios ni texto fuera del JSON.\n" +
      "- En cada 'texto', solo contenido del párrafo, sin estilos.\n\n" +
      "ENTRADA JSON:\n" + JSON.stringify(payload);

    return user;
  }

  function buildUserPromptSeleccion(seleccionTexto) {
    var user =
      "TAREA: MEJORAR_SELECCION\n\n" +
      "DEVOLUCIÓN ESTRICTA:\n" +
      "- Entrega únicamente el texto revisado, listo para sustituir la selección.\n" +
      "- No agregues comentarios ni prefijos/sufijos.\n" +
      "- Conserva los saltos de párrafo existentes.\n\n" +
      "CRITERIOS:\n" +
      "- Refuerza claridad, coherencia y técnica argumentativa.\n" +
      "- Usa terminología jurídica precisa según el contenido.\n" +
      "- Si hay numerales/listas, respeta su estructura.\n" +
      "- No inventes jurisprudencia ni doctrina; mejora la formulación técnica.\n\n" +
      "ENTRADA:\n" + seleccionTexto;
    return user;
  }

  function callOpenAI_chat(apiKey, systemPrompt, userPrompt, forceJson, onSuccess, onError) {
    var body = {
      "model": modelName,
      "messages": [
        { "role": "system", "content": systemPrompt },
        { "role": "user", "content": userPrompt }
      ],
      "temperature": 0
    };
    if (forceJson) {
      body["response_format"] = { "type": "json_object" };
    }

    postJson("https://api.openai.com/v1/chat/completions",
      { "Authorization": "Bearer " + apiKey },
      body,
      function (data) {
        try {
          var content = data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
          if (!content) throw new Error("Sin contenido de modelo.");
          onSuccess(content);
        } catch (e) {
          onError(e);
        }
      },
      onError
    );
  }

  function parseJsonSafe(text) {
    try {
      var t = (text || "").trim();
      if (t.indexOf("```") === 0) {
        t = t.replace(/^```(json)?/i, "").replace(/```$/i, "").trim();
      }
      return JSON.parse(t);
    } catch (e) {
      return null;
    }
  }

  function recogerContexto() {
    var ctx = {
      "jurisdiccion": document.getElementById("jurisdiccion").value || "El Salvador",
      "tipo": document.getElementById("tipoDoc").value || "",
      "materia": document.getElementById("materia").value || ""
    };
    return ctx;
  }

  function onRevisarDocumentoCompleto() {
    var apiKey = getApiKey();
    if (!apiKey) { log("Pegue su API Key."); return; }

    log("Recolectando párrafos...");
    Word.run(function (context) {
      var paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      return context.sync().then(function () {
        var items = paragraphs.items;
        for (var i = 0; i < items.length; i++) {
          items[i].load("text");
        }
        return context.sync().then(function () {
          var arr = [];
          for (var j = 0; j < items.length; j++) {
            var t = items[j].text || "";
            arr.push({ "index": j, "texto": t });
          }

          if (arr.length === 0) { log("No se encontraron párrafos."); return; }

          var contexto = recogerContexto();
          var batches = makeBatches(arr, 15000);
          log("Párrafos: " + arr.length + ". Lotes: " + batches.length + ".");

          processBatchSequential(apiKey, context, items, batches, contexto, 0, function () {
            log("Documento revisado.");
          }, function (err) {
            log("Error: " + (err.message || err));
          });
        });
      });
    }).catch(function (e) { log("Error Word.run: " + e.message); });
  }

  function makeBatches(parrafosArr, limitChars) {
    var batches = [];
    var current = [];
    var size = 0;
    for (var i = 0; i < parrafosArr.length; i++) {
      var p = parrafosArr[i];
      var add = JSON.stringify(p).length;
      if (size + add > limitChars && current.length > 0) {
        batches.push(current);
        current = [p];
        size = add;
      } else {
        current.push(p);
        size += add;
      }
    }
    if (current.length > 0) batches.push(current);
    return batches;
  }

  function processBatchSequential(apiKey, context, paragraphItems, batches, contexto, idx, onDone, onErr) {
    if (idx >= batches.length) { onDone(); return; }

    var batch = batches[idx];
    log("Lote " + (idx + 1) + " de " + batches.length + ": enviando a OpenAI...");
    var sys = buildSystemPrompt();
    var user = buildUserPromptDocumento(batch, contexto);

    callOpenAI_chat(apiKey, sys, user, true, function (content) {
      var json = parseJsonSafe(content);
      if (!json || !json.parrafos) {
        onErr(new Error("El modelo no devolvió JSON válido para el lote " + (idx + 1) + "."));
        return;
      }
      Word.run(function (ctx2) {
        for (var k = 0; k < json.parrafos.length; k++) {
          var item = json.parrafos[k];
          var i = item.index;
          var nuevo = item.texto || "";
          if (paragraphItems[i]) {
            paragraphItems[i].insertText(nuevo, InsertLocationReplace);
          }
        }
        return ctx2.sync();
      }).then(function () {
        log("Lote " + (idx + 1) + " aplicado.");
        processBatchSequential(apiKey, context, paragraphItems, batches, contexto, idx + 1, onDone, onErr);
      }).catch(function (e) { onErr(e); });
    }, onErr);
  }

  function onRevisarSeleccion() {
    var apiKey = getApiKey();
    if (!apiKey) { log("Pegue su API Key."); return; }

    Word.run(function (context) {
      var range = context.document.getSelection();
      range.load("text");
      return context.sync().then(function () {
        var t = range.text || "";
        if (!t) { log("No hay texto seleccionado."); return; }

        log("Enviando selección a OpenAI...");
        var sys = buildSystemPrompt();
        var user = buildUserPromptSeleccion(t);

        callOpenAI_chat(apiKey, sys, user, false, function (content) {
          Word.run(function (ctx2) {
            var r = ctx2.document.getSelection();
            r.insertText(content, InsertLocationReplace);
            return ctx2.sync();
          }).then(function () { log("Selección mejorada."); })
            .catch(function (e) { log("Error al aplicar selección: " + e.message); });
        }, function (err) { log("Error: " + (err.message || err)); });
      });
    }).catch(function (e) { log("Error Word.run: " + e.message); });
  }

})();