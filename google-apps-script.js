/**
 * SOLICITUD AVP UFRO - BACKEND GOOGLE APPS SCRIPT
 * Copia y pega este código en Extensiones > Apps Script de tu Google Sheet.
 * Este script mapea las columnas DINÁMICAMENTE por sus encabezados en la fila 1.
 */

// ==========================================
// FUNCIÓN PARA PROBAR CONEXIÓN Y AUTORIZAR PERMISOS
// ==========================================
// Selecciona esta función en el menú desplegable y presiona "Ejecutar"
function testConnection() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      Logger.log("ERROR: El script no está vinculado a una hoja. Asegúrate de abrir el script desde Extensiones > Apps Script dentro de tu Google Sheet.");
      return;
    }
    Logger.log("CONEXIÓN EXITOSA: Conectado a la hoja '" + ss.getName() + "'");
    
    const sheets = ["Inventario", "Alumnos", "Préstamos"];
    sheets.forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (sh) {
        Logger.log("Pestaña encontrada: '" + sheetName + "'");
        
        // Validar columnas
        const values = sh.getDataRange().getValues();
        if (values.length > 0) {
          const headers = values[0].map(normalizeHeader);
          Logger.log("  Columnas en '" + sheetName + "': " + JSON.stringify(values[0]));
          
          if (sheetName === "Inventario") {
            const indices = getInventoryHeaderIndices(headers);
            Logger.log("  Mapeo Inventario -> Categoria: " + indices.categoryIdx + ", Equipo: " + indices.nameIdx + ", Cantidad: " + indices.totalIdx + ", Disponible: " + indices.availableIdx + ", Códigos: " + indices.codesIdx);
            if (indices.categoryIdx === -1 || indices.nameIdx === -1 || indices.totalIdx === -1 || indices.availableIdx === -1 || indices.codesIdx === -1) {
              Logger.log("  ADVERTENCIA: Faltan columnas en 'Inventario'. Se requiere Categoría, Equipo, Cantidad, Disponible y Códigos.");
            }
          } else if (sheetName === "Alumnos") {
            const indices = getStudentHeaderIndices(headers);
            Logger.log("  Mapeo Alumnos -> Rut: " + indices.rutIdx + ", Nombre: " + indices.nameIdx + ", Apellido: " + indices.lastnameIdx + ", Fono: " + indices.fonoIdx + ", Email: " + indices.emailIdx + ", Observaciones: " + indices.obsIdx);
            if (indices.rutIdx === -1 || indices.nameIdx === -1 || indices.lastnameIdx === -1 || indices.emailIdx === -1) {
              Logger.log("  ADVERTENCIA: Faltan columnas en 'Alumnos'. Se requiere Rut, Nombre, Apellido y E-mail.");
            }
          } else if (sheetName === "Préstamos") {
            const indices = getLoanHeaderIndices(headers);
            Logger.log("  Mapeo Préstamos -> ID: " + indices.idIdx + ", RUT: " + indices.rutIdx + ", Nombre: " + indices.nameIdx + ", Email: " + indices.emailIdx + ", Equipo: " + indices.itemIdx + ", Código: " + indices.codeIdx + ", Estado: " + indices.statusIdx);
          }
        } else {
          Logger.log("  ADVERTENCIA: La pestaña '" + sheetName + "' está vacía.");
        }
      } else {
        Logger.log("ADVERTENCIA: No se encontró la pestaña '" + sheetName + "'. Asegúrate de crearla con ese nombre exacto.");
      }
    });
    
    // Prueba de envío de correo diagnóstico
    try {
      const myEmail = Session.getActiveUser().getEmail();
      if (myEmail) {
        Logger.log("PROBANDO ENVÍO DE CORREO DIAGNÓSTICO A: " + myEmail);
        MailApp.sendEmail({
          to: myEmail,
          subject: "Prueba de Permisos - AVP UFRO",
          htmlBody: "<h3>¡Conexión de correo exitosa!</h3><p>Este es un correo automático de prueba para confirmar que los permisos de envío de correos están habilitados correctamente en tu cuenta de Google.</p>"
        });
        Logger.log("¡CORREO DE PRUEBA ENVIADO CON ÉXITO! Revisa tu bandeja de entrada (" + myEmail + ") para confirmarlo.");
      } else {
        Logger.log("ADVERTENCIA: No se pudo determinar tu dirección de correo electrónico activa para la prueba de envío.");
      }
    } catch (mailErr) {
      Logger.log("ERROR CRÍTICO AL ENVIAR CORREO DIAGNÓSTICO: " + mailErr.toString());
      Logger.log("Sugerencia: Si el error dice 'Se requiere autorización' o similar, vuelve a ejecutar el script y asegúrate de otorgar todos los permisos solicitados.");
    }
  } catch (e) {
    Logger.log("ERROR al probar conexión: " + e.toString());
  }
}

function autoUpgradeHeaders(ss) {
  if (!ss) return;
  
  const loanSheet = ss.getSheetByName("Préstamos");
  if (loanSheet) {
    const values = loanSheet.getDataRange().getValues();
    if (values.length > 0) {
      const currentHeaders = values[0].map(h => h.toString().trim().toLowerCase());
      const expected = [
        { name: "ID Préstamo", match: "id" },
        { name: "RUT Alumno", match: "rut" },
        { name: "Nombre Alumno", match: "nombre" },
        { name: "Email Alumno", match: "email" },
        { name: "Equipo", match: "equipo" },
        { name: "Código Inventario", match: "codigo" },
        { name: "Fecha Registro", match: "registro" },
        { name: "Fecha Retiro", match: "retiro" },
        { name: "Fecha Devolución", match: "devolucion" },
        { name: "Estado", match: "estado" },
        { name: "Observaciones", match: "observacio" },
        { name: "Fecha Retiro Programada", match: "retiro programada" },
        { name: "Fecha Devolución Programada", match: "devolucion programada" },
        { name: "Asignatura", match: "asignatura" },
        { name: "Observación Devolución", match: "observacion devolucion" },
        { name: "Días Atraso", match: "atraso" }
      ];
      
      let modified = false;
      const headersToAppend = [];
      
      expected.forEach(item => {
        const found = currentHeaders.some(h => {
          const norm = h.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
          const itemMatch = item.match.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
          if (itemMatch === "retiro" && norm.includes("programad")) return false;
          if (itemMatch === "devolucion" && norm.includes("programad")) return false;
          return norm.includes(itemMatch);
        });
        
        if (!found) {
          headersToAppend.push(item.name);
          modified = true;
        }
      });
      
      if (modified && headersToAppend.length > 0) {
        const lastCol = loanSheet.getLastColumn();
        const range = loanSheet.getRange(1, lastCol + 1, 1, headersToAppend.length);
        range.setValues([headersToAppend]);
        SpreadsheetApp.flush();
      }
    }
  }

  const studentSheet = ss.getSheetByName("Alumnos");
  if (studentSheet) {
    const values = studentSheet.getDataRange().getValues();
    if (values.length > 0) {
      const currentHeaders = values[0].map(h => h.toString().trim().toLowerCase());
      const hasBlockDate = currentHeaders.some(h => {
        const norm = h.normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
        return norm.includes("bloqueo") || norm.includes("sancion");
      });
      if (!hasBlockDate) {
        const lastCol = studentSheet.getLastColumn();
        studentSheet.getRange(1, lastCol + 1, 1, 1).setValue("Fecha Bloqueo");
        SpreadsheetApp.flush();
      }
    }
  }
}

// ==========================================
// CONTROLADOR DE PETICIONES GET (LECTURA)
// ==========================================
function doGet(e) {
  const action = (e && e.parameter) ? e.parameter.action : null;
  
  // Si no se especifica ninguna acción (ej. cuando el usuario hace clic en el enlace directamente en su navegador)
  if (!action) {
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>AVP UFRO - Conexión Exitosa</title>
        <style>
          body { font-family: 'Plus Jakarta Sans', 'Segoe UI', -apple-system, sans-serif; background: #f8fafc; color: #1e293b; display: flex; align-items: center; justify-content: center; height: 100vh; margin: 0; padding: 20px; text-align: center; }
          .card { background: white; padding: 40px; border-radius: 16px; box-shadow: 0 10px 30px rgba(15, 23, 42, 0.05); max-width: 500px; border: 1px solid #e2e8f0; text-align: center; }
          .icon { font-size: 56px; margin-bottom: 20px; }
          h1 { color: #4f46e5; margin-top: 0; font-size: 24px; font-weight: 800; letter-spacing: -0.5px; }
          p { font-size: 14px; line-height: 1.6; color: #475569; margin-bottom: 15px; }
          .url { background: #f1f5f9; padding: 12px; border-radius: 8px; font-family: monospace; font-size: 12px; color: #0f172a; word-break: break-all; margin: 20px 0; border: 1px solid #cbd5e1; user-select: all; cursor: pointer; font-weight: bold; }
          .badge { display: inline-block; background: #dcfce7; color: #15803d; font-weight: bold; padding: 6px 14px; border-radius: 9999px; font-size: 11px; margin-bottom: 15px; letter-spacing: 0.5px; }
          .btn-copy { display: inline-block; background: #4f46e5; color: white; border: none; padding: 10px 20px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; text-decoration: none; transition: background 0.2s; }
          .btn-copy:hover { background: #4338ca; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="icon">⚡</div>
          <div class="badge">CONEXIÓN EXITOSA</div>
          <h1>¡Backend AVP UFRO Activo!</h1>
          <p>El script se ha conectado correctamente a tu Google Sheet y la Web App está funcionando.</p>
          <p>Copia esta URL e ingrésala en el <strong>Panel Admin > Configurar Equipos y Alumnos</strong> (abajo del todo) en tu aplicación visual (index.html):</p>
          <div class="url" id="script-url-box" title="Haz clic para seleccionar todo" onclick="selectText(this)">https://script.google.com/macros/s/.../exec</div>
          <button class="btn-copy" onclick="copyUrl()">Copiar URL al Portapapeles</button>
          <p style="font-size:11px; color:#94a3b8; margin-top:25px; line-height: 1.4;">Esta página informativa aparece al acceder al enlace directamente desde un navegador. El frontend visual (index.html) se comunicará con ella automáticamente por detrás.</p>
        </div>
        <script>
          const cleanUrl = window.location.href.split('?')[0].replace('/dev', '/exec');
          document.getElementById('script-url-box').textContent = cleanUrl;
          
          function selectText(element) {
            var range = document.createRange();
            range.selectNodeContents(element);
            var selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
          }
          
          function copyUrl() {
            const el = document.createElement('textarea');
            el.value = cleanUrl;
            document.body.appendChild(el);
            el.select();
            document.execCommand('copy');
            document.body.removeChild(el);
            alert('¡URL copiada con éxito!');
          }
        </script>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(html);
  }

  let responseData = {};

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    autoUpgradeHeaders(ss);
    
    if (action === "getInitData") {
      const isFresh = (e && e.parameter && (e.parameter.fresh === "1" || e.parameter._t));
      responseData = getCachedInitData(ss, isFresh);
    } 
    else if (action === "checkStudent") {
      const rut = e.parameter.rut;
      const student = findStudentByRut(ss, rut);
      if (student) {
        responseData = { status: "success", student: student };
      } else {
        responseData = { status: "error", message: "RUT no registrado en la base de datos de Periodismo." };
      }
    } 
    else if (action === "searchLoans") {
      const query = e.parameter.query.toUpperCase();
      const loans = searchPendingLoans(ss, query);
      responseData = { status: "success", loans: loans };
    } 
    else {
      responseData = { status: "error", message: "Acción GET no válida." };
    }
  } catch (error) {
    responseData = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

function clearInitDataCache() {
  try {
    const cache = CacheService.getScriptCache();
    const keys = ["avp_inv", "avp_st", "avp_ln", "avp_sb", "avp_cat", "avp_url"];
    cache.removeAll(keys);
    keys.forEach(k => {
      try { cache.remove(k); } catch(err) {}
    });
  } catch (e) {
    Logger.log("Error clearing cache: " + e.toString());
  }
}

function getCachedInitData(ss, isFresh) {
  const cache = CacheService.getScriptCache();
  
  // Extraer información de diagnóstico del Inventario
  let debugInfo = null;
  const invSheet = ss.getSheetByName("Inventario");
  if (invSheet) {
    const v = invSheet.getDataRange().getValues();
    if (v.length > 0) {
      const headers = v[0].map(normalizeHeader);
      const indices = getInventoryHeaderIndices(headers);
      const sampleRows = [];
      for (let i = 1; i < Math.min(v.length, 15); i++) {
        if (!v[i][indices.nameIdx]) continue;
        const totalNum = parseInt(v[i][indices.totalIdx]) || 0;
        const rawD = v[i][indices.availableIdx];
        const dispNum = (rawD === "" || rawD === undefined || rawD === null) ? totalNum : (parseInt(rawD) || 0);
        sampleRows.push({
          row: i + 1,
          name: v[i][indices.nameIdx].toString(),
          total: totalNum,
          disponible: dispNum
        });
      }
      debugInfo = {
        detectedHeaders: v[0],
        normalizedHeaders: headers,
        indices: indices,
        sampleRows: sampleRows
      };
    }
  }

  // SIEMPRE leer préstamos en vivo directamente de la hoja de cálculo (no se cachean)
  const loans = getLoansData(ss);

  if (!isFresh) {
    try {
      const cachedInv = cache.get("avp_inv");
      const cachedSt = cache.get("avp_st");
      const cachedSb = cache.get("avp_sb");
      const cachedCat = cache.get("avp_cat");
      const cachedUrl = cache.get("avp_url");
      
      if (cachedInv && cachedSt && cachedSb) {
        let catParsed = cachedCat ? JSON.parse(cachedCat) : null;
        if (!catParsed || catParsed.length === 0) {
          catParsed = getCategoriesData(ss, JSON.parse(cachedInv));
        }
        return {
          status: "success",
          inventory: JSON.parse(cachedInv),
          students: JSON.parse(cachedSt),
          loans: loans, // Préstamos frescos 100% en tiempo real
          subjects: JSON.parse(cachedSb),
          categories: catParsed,
          sheetUrl: cachedUrl || ss.getUrl(),
          debugInfo: debugInfo,
          fromCache: true
        };
      }
    } catch (e) {
      Logger.log("Cache read error, falling back to sheet: " + e.toString());
    }
  }
  
  const inventory = getInventoryData(ss);
  const students = getStudentsData(ss);
  const subjects = getSubjectsData(ss);
  const categories = getCategoriesData(ss, inventory);
  const sheetUrl = ss.getUrl();
  
  try {
    const invStr = JSON.stringify(inventory);
    const stStr = JSON.stringify(students);
    const sbStr = JSON.stringify(subjects);
    const catStr = JSON.stringify(categories);
    
    const entries = {};
    if (invStr.length < 95000) entries["avp_inv"] = invStr;
    if (stStr.length < 95000) entries["avp_st"] = stStr;
    if (sbStr.length < 95000) entries["avp_sb"] = sbStr;
    if (catStr.length < 95000) entries["avp_cat"] = catStr;
    entries["avp_url"] = sheetUrl;
    
    cache.putAll(entries, 300); // 5 minutos de caché para inventario, alumnos y categorías
  } catch (e) {
    Logger.log("Cache write error: " + e.toString());
  }
  
  return {
    status: "success",
    inventory: inventory,
    students: students,
    loans: loans, // Préstamos en vivo
    subjects: subjects,
    categories: categories,
    sheetUrl: sheetUrl,
    debugInfo: debugInfo,
    fromCache: false
  };
}

// ==========================================
// CONTROLADOR DE PETICIONES POST (ESCRITURA)
// ==========================================
function doPost(e) {
  let responseData = {};
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    autoUpgradeHeaders(ss);
    const action = e.parameter.action;
    const postData = JSON.parse(e.postData.contents);
    
    if (action === "createLoan") {
      responseData = executeCreateLoan(ss, postData);
    } 
    else if (action === "deliverLoan") {
      responseData = executeDeliverLoan(ss, postData);
    } 
    else if (action === "returnLoan") {
      responseData = executeReturnLoan(ss, postData);
    } 
    else if (action === "cancelLoan") {
      responseData = executeCancelLoan(ss, postData);
    } 
    else if (action === "addEquipment") {
      responseData = executeAddEquipment(ss, postData);
    }
    else if (action === "deleteEquipment") {
      responseData = executeDeleteEquipment(ss, postData);
    }
    else if (action === "addStudent") {
      responseData = executeAddStudent(ss, postData);
    }
    else if (action === "deleteStudent") {
      responseData = executeDeleteStudent(ss, postData);
    }
    else if (action === "applySanction") {
      responseData = executeApplySanction(ss, postData);
    }
    else if (action === "removeSanction") {
      responseData = executeRemoveSanction(ss, postData);
    }
    else if (action === "sendContactEmail") {
      responseData = executeSendContactEmail(ss, postData);
    }
    else if (action === "addCategory") {
      responseData = executeAddCategory(ss, postData);
    }
    else {
      responseData = { status: "error", message: "Acción POST no válida." };
    }
    
    if (responseData && responseData.status === "success") {
      clearInitDataCache();
    }
  } catch (error) {
    responseData = { status: "error", message: error.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(responseData))
    .setMimeType(ContentService.MimeType.JSON);
}

// ==========================================
// UTILIDAD DE NORMALIZACIÓN DE CABECERAS
// ==========================================
function normalizeHeader(str) {
  return str.toString()
            .trim()
            .toLowerCase()
            .normalize("NFD")
            .replace(/[\u0300-\u036f]/g, ""); // Remueve tildes y diacríticos
}

// ==========================================
// BUSCADORES FLEXIBLES DE COLUMNAS POR ALIAS
// ==========================================
function getInventoryHeaderIndices(headers) {
  const categoryIdx = headers.findIndex(h => h.includes("categoria") || h.includes("tipo"));
  const nameIdx = headers.findIndex(h => h.includes("equipo") || h === "nombre" || h === "articulo" || h.includes("elemento"));
  const availableIdx = headers.findIndex(h => h.includes("disponible") || h.includes("dispo"));
  const totalIdx = headers.findIndex((h, idx) => {
    if (idx === availableIdx) return false;
    return h.includes("cantidad") || h.includes("total") || h.includes("stock") || h === "cant";
  });
  const codesIdx = headers.findIndex((h, idx) => {
    if (idx === categoryIdx || idx === nameIdx || idx === totalIdx || idx === availableIdx) return false;
    return h.includes("codigo") || h.includes("inventario") || h.includes("nº") || h.includes("nro") || h.includes("serial") || h.includes("id");
  });
  const imageIdx = headers.findIndex(h => h.includes("imagen") || h.includes("foto") || h.includes("url") || h.includes("image"));
  const descriptionIdx = headers.findIndex(h => h.includes("descripcion") || h.includes("descripci") || h.includes("detalle") || h.includes("info"));
  return { categoryIdx, nameIdx, totalIdx, availableIdx, codesIdx, imageIdx, descriptionIdx };
}

function getStudentHeaderIndices(headers) {
  const rutIdx = headers.findIndex(h => h.includes("rut") || h.includes("run"));
  const nameIdx = headers.findIndex(h => h === "nombre" || h === "nombres" || h.includes("primer nombre"));
  const lastnameIdx = headers.findIndex(h => h.includes("apellido") || h.includes("apellidos"));
  const fonoIdx = headers.findIndex(h => h.includes("fono") || h.includes("telefono") || h.includes("celular") || h.includes("movil") || h.includes("contacto"));
  const emailIdx = headers.findIndex(h => h.includes("email") || h.includes("e-mail") || h.includes("correo") || h.includes("mail"));
  const blockDateIdx = headers.findIndex(h => (h.includes("fecha") && (h.includes("bloqueo") || h.includes("sancion"))) || h.includes("bloqueado el") || h.includes("bloqueado desde") || h.includes("fecha bloqueo"));
  const obsIdx = headers.findIndex((h, idx) => {
    if (idx === blockDateIdx) return false;
    return h.includes("observacio") || h.includes("deuda") || h.includes("bloqueo") || h.includes("comentario") || h.includes("detalle");
  });
  return { rutIdx, nameIdx, lastnameIdx, fonoIdx, emailIdx, obsIdx, blockDateIdx };
}

function getLoanHeaderIndices(headers) {
  const idIdx = headers.findIndex(h => h.includes("id"));
  const rutIdx = headers.findIndex(h => h.includes("rut") || h.includes("run"));
  const nameIdx = headers.findIndex(h => h.includes("nombre") && !h.includes("equipo") && !h.includes("articulo"));
  const emailIdx = headers.findIndex(h => h.includes("email") || h.includes("e-mail") || h.includes("correo") || h.includes("mail"));
  const itemIdx = headers.findIndex(h => h.includes("equipo") || h.includes("articulo") || h.includes("item"));
  const codeIdx = headers.findIndex((h, idx) => {
    if (idx === idIdx) return false;
    return h.includes("codigo") || h.includes("inventario") || h.includes("nº") || h.includes("nro") || h.includes("serial");
  });
  
  // Fecha Registro / Solicitud original (reserva)
  let dateOutIdx = headers.findIndex(h => h.includes("registro") || h.includes("solicitud"));
  if (dateOutIdx === -1) {
    dateOutIdx = headers.findIndex(h => {
      const norm = h.toString().toLowerCase();
      if (norm.includes("programad") || norm.includes("retiro") || norm.includes("devolucion") || norm.includes("entrega")) return false;
      return norm.includes("fecha");
    });
  }
  
  // Fecha Retiro / Entrega Real (cuando se retira físicamente)
  const dateDeliverIdx = headers.findIndex(h => {
    const norm = h.toString().toLowerCase();
    if (norm.includes("registro") || norm.includes("solicitud")) return false;
    if (norm.includes("devolucion") || norm.includes("retorno") || norm.includes("entrada")) return false;
    return (norm.includes("retiro") && !norm.includes("programad") && !norm.includes("previst") && !norm.includes("estimad")) ||
           (norm.includes("entrega") && !norm.includes("programad") && !norm.includes("previst") && !norm.includes("estimad"));
  });
  
  // Fecha Devolución Real (excluyendo programadas y observaciones)
  const dateInIdx = headers.findIndex(h => {
    const norm = h.toString().toLowerCase();
    if (norm.includes("programad") || norm.includes("previst") || norm.includes("estimad") || norm.includes("planificad")) return false;
    if (norm.includes("observaci") || norm.includes("obs") || norm.includes("nota")) return false;
    return norm.includes("devolucion") || norm.includes("retorno") || norm.includes("entrada");
  });
  const statusIdx = headers.findIndex(h => h.includes("estado") || h.includes("status"));
  const progRetiroIdx = headers.findIndex(h => (h.includes("programad") || h.includes("previst") || h.includes("estimad") || h.includes("planificad")) && (h.includes("retiro") || h.includes("salida")));
  const progDevolucionIdx = headers.findIndex(h => (h.includes("programad") || h.includes("previst") || h.includes("estimad") || h.includes("planificad")) && (h.includes("devolucion") || h.includes("retorno")));
  const subjectIdx = headers.findIndex(h => h.includes("asignatura") || h.includes("catedra") || h.includes("clase") || h.includes("materia") || h.includes("curso"));
  const obsReturnIdx = headers.findIndex(h => {
    const norm = h.toString().toLowerCase();
    return (norm.includes("observacio") || norm.includes("obs") || norm.includes("nota") || norm.includes("comentario")) && 
           (norm.includes("devolucion") || norm.includes("retorno") || norm.includes("recib"));
  });
  const obsIdx = headers.findIndex((h, idx) => {
    if (idx === obsReturnIdx) return false;
    return h.includes("observacio") || h.includes("obs") || h.includes("nota") || h.includes("comentario");
  });
  const daysOverdueIdx = headers.findIndex(h => h.includes("atraso") || h.includes("retraso") || h.includes("mora") || h.includes("dias atraso"));
  return { idIdx, rutIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateOutIdx, dateDeliverIdx, dateInIdx, statusIdx, progRetiroIdx, progDevolucionIdx, subjectIdx, obsIdx, daysOverdueIdx, obsReturnIdx };
}

function getSubjectsData(ss) {
  const sheet = ss.getSheetByName("Asignaturas");
  if (!sheet) {
    return ["Periodismo Escrito", "Periodismo Radial", "Periodismo Televisivo", "Fotoperiodismo", "Comunicación Digital", "Cine y Documental", "Proyecto de Título", "Ninguna (Proyecto Personal)"];
  }
  
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  const subjects = [];
  for (let i = 1; i < values.length; i++) {
    const val = values[i][0];
    if (val && val.toString().trim() !== "") {
      subjects.push(val.toString().trim());
    }
  }
  return subjects;
}

function cleanCategoryGAS(cat) {
  if (!cat) return "";
  const clean = cat.toString().trim().replace(/\s+/g, ' ');
  const norm = clean.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  if (norm === "camaras" || norm === "camara") return "Cámaras";
  if (norm === "tripodes" || norm === "tripode") return "Trípodes";
  if (norm === "microfonos" || norm === "microfono") return "Micrófonos";
  if (norm === "iluminacion" || norm === "iluminaciones" || norm === "luces") return "Iluminación";
  return clean;
}

function getCategoriesData(ss, inventory) {
  let categories = [];
  const catSheet = ss.getSheetByName("Categorías") || ss.getSheetByName("Categorias");
  if (catSheet) {
    const values = catSheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      const val = values[i][0];
      if (val && val.toString().trim() !== "") {
        categories.push(cleanCategoryGAS(val.toString()));
      }
    }
  }
  
  // Extraer también las categorías existentes en la pestaña Inventario
  if (inventory && Array.isArray(inventory)) {
    inventory.forEach(item => {
      if (item.category && item.category.toString().trim() !== "") {
        categories.push(cleanCategoryGAS(item.category.toString()));
      }
    });
  }
  
  // Categorías base por defecto si no hubiese ninguna
  if (categories.length === 0) {
    categories = ["Cámaras", "Trípodes", "Audio", "Luces"];
  }
  
  // Desduplicar preservando formato canónico
  const unique = [];
  const seen = new Set();
  categories.forEach(cat => {
    const norm = cat.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    if (!seen.has(norm) && norm !== "") {
      seen.add(norm);
      unique.push(cat);
    }
  });

  // Asegurar que Cámaras esté siempre como la primera pestaña si existe
  const camIdx = unique.findIndex(c => c.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "") === "camaras");
  if (camIdx > 0) {
    const camCat = unique.splice(camIdx, 1)[0];
    unique.unshift(camCat);
  }

  return unique;
}

// ==========================================
// LECTURA DINÁMICA DE HOJAS
// ==========================================

function getInventoryData(ss) {
  const sheet = ss.getSheetByName("Inventario");
  if (!sheet) throw new Error("No se encontró la pestaña 'Inventario'");
  
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  const headers = values[0].map(normalizeHeader);
  const { categoryIdx, nameIdx, totalIdx, availableIdx, codesIdx, imageIdx, descriptionIdx } = getInventoryHeaderIndices(headers);
  
  if (categoryIdx === -1 || nameIdx === -1 || totalIdx === -1 || availableIdx === -1 || codesIdx === -1) {
    throw new Error("Pestaña 'Inventario' requiere columnas de: Categoría, Equipo, Cantidad, Disponible y Códigos de Inventario.");
  }
  
  const inventory = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[nameIdx]) continue;
    
    const totalVal = parseInt(row[totalIdx]) || 0;
    const rawDisp = row[availableIdx];
    const availableVal = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
    
    inventory.push({
      id: i,
      category: cleanCategoryGAS(row[categoryIdx]),
      name: row[nameIdx].toString(),
      total: totalVal,
      available: availableVal,
      codes: row[codesIdx] ? row[codesIdx].toString().split(",").map(c => c.trim()) : [],
      image: imageIdx !== -1 && row[imageIdx] ? row[imageIdx].toString().trim() : "",
      description: descriptionIdx !== -1 && row[descriptionIdx] ? row[descriptionIdx].toString().trim() : ""
    });
  }
  return inventory;
}

function getStudentsData(ss) {
  const sheet = ss.getSheetByName("Alumnos");
  if (!sheet) throw new Error("No se encontró la pestaña 'Alumnos'");
  
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  const headers = values[0].map(normalizeHeader);
  const { rutIdx, nameIdx, lastnameIdx, fonoIdx, emailIdx, obsIdx, blockDateIdx } = getStudentHeaderIndices(headers);
  
  if (rutIdx === -1 || nameIdx === -1 || lastnameIdx === -1 || emailIdx === -1) {
    throw new Error("Pestaña 'Alumnos' requiere columnas de: Rut, Nombre, Apellido y E-mail.");
  }
  
  const students = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[rutIdx]) continue;
    
    const obs = obsIdx !== -1 && row[obsIdx] ? row[obsIdx].toString().trim() : "";
    const isBlocked = obs !== "";
    
    let blockDate = "";
    if (blockDateIdx !== -1 && row[blockDateIdx]) {
      const rawDate = row[blockDateIdx];
      if (rawDate instanceof Date) {
        blockDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone() || "America/Santiago", "yyyy-MM-dd HH:mm:ss");
      } else {
        blockDate = rawDate.toString().trim();
      }
    }
    
    students.push({
      rut: row[rutIdx].toString(),
      name: row[nameIdx].toString(),
      lastname: row[lastnameIdx].toString(),
      fono: fonoIdx !== -1 ? row[fonoIdx].toString() : "",
      email: row[emailIdx].toString(),
      status: isBlocked ? "Bloqueado" : "Activo",
      debt: obs,
      blockDate: isBlocked ? blockDate : ""
    });
  }
  return students;
}

function splitCellValues(val, expectedCount) {
  if (val === undefined || val === null) return new Array(expectedCount).fill("");
  if (val instanceof Date) {
    const list = [val];
    while (list.length < expectedCount) {
      list.push("");
    }
    return list;
  }
  
  let list = val.toString().split("\n").map(x => x.trim());
  if (list.length < expectedCount && val.toString().includes(",")) {
    list = val.toString().split(",").map(x => x.trim());
  }
  
  while (list.length < expectedCount) {
    list.push("");
  }
  return list;
}

function getLoansData(ss) {
  const sheet = ss.getSheetByName("Préstamos");
  if (!sheet) return [];
  
  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];
  
  const headers = values[0].map(normalizeHeader);
  const { idIdx, rutIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateOutIdx, dateDeliverIdx, dateInIdx, statusIdx, progRetiroIdx, progDevolucionIdx, subjectIdx, obsIdx, daysOverdueIdx, obsReturnIdx } = getLoanHeaderIndices(headers);
  
  if (idIdx === -1 || rutIdx === -1 || nameIdx === -1 || itemIdx === -1 || codeIdx === -1 || statusIdx === -1) {
    return [];
  }
  
  const loans = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[idIdx]) continue;
    
    const rawItems = row[itemIdx].toString();
    const itemsList = rawItems.split("\n").map(x => x.trim()).filter(Boolean);
    const expectedCount = itemsList.length;
    if (expectedCount === 0) continue;
    
    const codesList = splitCellValues(row[codeIdx], expectedCount);
    const obsList = splitCellValues(obsIdx !== -1 ? row[obsIdx] : "", expectedCount);
    const obsReturnList = splitCellValues(obsReturnIdx !== -1 ? row[obsReturnIdx] : "", expectedCount);
    
    // Leer estado general o lista de estados (compatibilidad retroactiva)
    const rawStatus = statusIdx !== -1 ? row[statusIdx].toString().trim() : "Solicitado";
    let statusesList = [];
    if (rawStatus.includes("\n")) {
      statusesList = splitCellValues(row[statusIdx], expectedCount);
    } else {
      statusesList = new Array(expectedCount).fill(rawStatus);
    }
    
    // Leer fecha devolución general o lista de fechas (compatibilidad retroactiva)
    const rawDateIn = dateInIdx !== -1 && row[dateInIdx] ? row[dateInIdx] : "";
    let datesInList = [];
    if (rawDateIn instanceof Date) {
      datesInList = new Array(expectedCount).fill(rawDateIn);
    } else if (rawDateIn.toString().includes("\n")) {
      datesInList = splitCellValues(row[dateInIdx], expectedCount);
    } else {
      datesInList = new Array(expectedCount).fill(rawDateIn);
    }
    
    for (let k = 0; k < expectedCount; k++) {
      let itemCode = codesList[k] || "Pte. Entrega";
      let itemStatus = statusesList[k] || "Solicitado";
      let itemDateIn = datesInList[k] ? formatDate(datesInList[k]) : "";
      
      // Si el equipo fue anulado a nivel de código, su estado es "Anulado" y no tiene fecha de devolución
      if (itemCode.toUpperCase() === "ANULADO") {
        itemStatus = "Anulado";
        itemDateIn = "";
      }
      
      loans.push({
        id: row[idIdx].toString(),
        rut: row[rutIdx].toString(),
        name: row[nameIdx].toString(),
        email: emailIdx !== -1 ? row[emailIdx].toString() : "",
        item: itemsList[k],
        code: itemCode,
        dateOut: dateOutIdx !== -1 ? formatDate(row[dateOutIdx]) : "",
        dateDeliver: dateDeliverIdx !== -1 && row[dateDeliverIdx] ? formatDate(row[dateDeliverIdx]) : "",
        dateIn: itemDateIn,
        status: itemStatus,
        progRetiro: progRetiroIdx !== -1 && row[progRetiroIdx] ? formatDate(row[progRetiroIdx]) : "",
        progDevolucion: progDevolucionIdx !== -1 && row[progDevolucionIdx] ? formatDate(row[progDevolucionIdx]) : "",
        subject: subjectIdx !== -1 ? row[subjectIdx].toString() : "",
        obs: obsList[k] || "",
        obsReturn: obsReturnList[k] || "",
        daysOverdue: daysOverdueIdx !== -1 && row[daysOverdueIdx] ? parseInt(row[daysOverdueIdx]) || 0 : 0
      });
    }
  }
  return loans;
}

function findStudentByRut(ss, rut) {
  const students = getStudentsData(ss);
  const cleanSearch = rut.replace(/[^0-9kK]/g, '').toLowerCase();
  
  return students.find(s => {
    const cleanStudentRut = s.rut.replace(/[^0-9kK]/g, '').toLowerCase();
    return cleanStudentRut === cleanSearch;
  });
}

function searchPendingLoans(ss, query) {
  const loans = getLoansData(ss);
  const cleanQuery = query.replace(/[^0-9kK]/g, '').toLowerCase();
  
  return loans.filter(loan => {
    const isRetirado = loan.status === "Retirado";
    const cleanStudentRut = loan.rut.replace(/[^0-9kK]/g, '').toLowerCase();
    const matchRut = cleanStudentRut === cleanQuery;
    const matchCode = loan.code.toUpperCase() === query.toUpperCase();
    return isRetirado && (matchRut || matchCode);
  });
}

// ==========================================
// ESCRITURA DINÁMICA
// ==========================================

function executeCreateLoan(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado. Intenta de nuevo." };
  }
  
  try {
    const invSheet = ss.getSheetByName("Inventario");
    let loanSheet = ss.getSheetByName("Préstamos");
    
    if (!loanSheet) {
      loanSheet = ss.insertSheet("Préstamos");
      loanSheet.appendRow(["ID Préstamo", "RUT Alumno", "Nombre Alumno", "Email Alumno", "Equipo", "Código Inventario", "Fecha Registro", "Fecha Devolución", "Estado", "Observaciones"]);
    }
    
    const loanValues = loanSheet.getDataRange().getValues();
    const loanHeaders = loanValues[0].map(normalizeHeader);
    const { idIdx, rutIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateOutIdx, dateInIdx, statusIdx, progRetiroIdx, progDevolucionIdx, subjectIdx, obsIdx } = getLoanHeaderIndices(loanHeaders);
    
    let nextIdNumber = 1001;
    if (loanValues.length > 1 && idIdx !== -1) {
      const lastIdVal = loanValues[loanValues.length - 1][idIdx].toString();
      const numMatch = lastIdVal.match(/\d+/);
      if (numMatch) {
        nextIdNumber = parseInt(numMatch[0]) + 1;
      }
    }
    
    const student = payload.student;
    const items = payload.items;
    const timestamp = payload.timestamp;
    const progRetiro = payload.progRetiro;
    const progDevolucion = payload.progDevolucion;
    const subject = payload.subject;
    
    // Obtener y mapear Inventario
    const invValues = invSheet.getDataRange().getValues();
    const invHeaders = invValues[0].map(normalizeHeader);
    const { nameIdx: invNameIdx, totalIdx: invTotalIdx, availableIdx: invDispIdx } = getInventoryHeaderIndices(invHeaders);
    
    const updates = [];
    for (let k = 0; k < items.length; k++) {
      const item = items[k];
      let nameMatched = false;
      let hasStock = false;
      let currentStock = 0;
      
      const cleanSearchName = item.name.toString().trim().toLowerCase();
      
      for (let i = 1; i < invValues.length; i++) {
        const rowName = invValues[i][invNameIdx] ? invValues[i][invNameIdx].toString().trim().toLowerCase() : "";
        if (rowName === cleanSearchName) {
          nameMatched = true;
          const totalVal = parseInt(invValues[i][invTotalIdx]) || 0;
          const rawDisp = invValues[i][invDispIdx];
          const disponible = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
          currentStock = disponible;
          
          if (disponible > 0) {
            hasStock = true;
            updates.push({ rowIndex: i + 1, currentDisp: disponible });
            // Reducir stock temporalmente en memoria para esta transacción si hay duplicados
            invValues[i][invDispIdx] = disponible - 1;
            break;
          }
        }
      }
      
      if (!nameMatched) {
        throw new Error("El equipo '" + item.name + "' no se encontró en la pestaña 'Inventario' de Google Sheets. Verifica que el nombre esté escrito exactamente igual en la hoja.");
      }
      if (!hasStock) {
        throw new Error("El equipo '" + item.name + "' no tiene stock disponible en Google Sheets (Disponible: " + currentStock + "). Actualiza la columna 'Disponible' en la hoja 'Inventario' para poder prestarlo.");
      }
    }
    
    // Descontar stock
    for (let k = 0; k < items.length; k++) {
      const updateInfo = updates[k];
      invSheet.getRange(updateInfo.rowIndex, invDispIdx + 1).setValue(updateInfo.currentDisp - 1);
    }
    
    const studentFullName = student.name + " " + (student.lastname || "");
    const loanId = "L-" + nextIdNumber;
    
    // Crear fila dinámica única agrupando todos los ítems por saltos de línea (\n)
    const newRow = new Array(loanHeaders.length).fill("");
    if (idIdx !== -1) newRow[idIdx] = loanId;
    if (rutIdx !== -1) newRow[rutIdx] = student.rut;
    if (nameIdx !== -1) newRow[nameIdx] = studentFullName;
    if (emailIdx !== -1) newRow[emailIdx] = student.email;
    if (itemIdx !== -1) newRow[itemIdx] = items.map(it => it.name).join("\n");
    if (codeIdx !== -1) newRow[codeIdx] = items.map(it => it.code || "Pte. Entrega").join("\n");
    if (dateOutIdx !== -1) newRow[dateOutIdx] = timestamp;
    if (dateInIdx !== -1) newRow[dateInIdx] = "";
    if (statusIdx !== -1) newRow[statusIdx] = "Solicitado";
    if (progRetiroIdx !== -1) newRow[progRetiroIdx] = progRetiro || "";
    if (progDevolucionIdx !== -1) newRow[progDevolucionIdx] = progDevolucion || "";
    if (subjectIdx !== -1) newRow[subjectIdx] = subject || "";
    if (obsIdx !== -1) newRow[obsIdx] = items.map(() => "").join("\n");
    
    loanSheet.appendRow(newRow);
    SpreadsheetApp.flush();
    
    try {
      sendSolicitudEmail(student, items, timestamp, progRetiro, progDevolucion, subject, loanId);
    } catch (emailError) {
      Logger.log("ERROR al enviar email de solicitud: " + emailError.toString());
      return { status: "success", loanId: loanId, message: "Solicitud registrada con éxito. (Nota: No se pudo enviar el correo de comprobación por falta de destinatario o configuración)." };
    }
    return { status: "success", loanId: loanId, message: "Solicitud registrada con éxito." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 2. Registrar Retiro Físico (Estado: Retirado, Email 2)
function executeDeliverLoan(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const loanSheet = ss.getSheetByName("Préstamos");
    const invSheet = ss.getSheetByName("Inventario");
    
    const loanId = payload.loanId;
    const items = payload.items; // Array of { name, code, obs }
    const timestamp = payload.timestamp;
    
    const values = loanSheet.getDataRange().getValues();
    const headers = values[0].map(normalizeHeader);
    const { idIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateOutIdx, dateDeliverIdx, statusIdx, obsIdx } = getLoanHeaderIndices(headers);
    
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][idIdx].toString() === loanId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("No se encontró la solicitud con ID " + loanId);
    
    const studentName = values[rowIndex - 1][nameIdx].toString();
    const studentEmail = emailIdx !== -1 ? values[rowIndex - 1][emailIdx].toString() : "";
    
    // Split current values
    const currentItems = values[rowIndex - 1][itemIdx].toString().split("\n").map(x => x.trim()).filter(Boolean);
    const expectedCount = currentItems.length;
    
    const currentCodes = splitCellValues(codeIdx !== -1 ? values[rowIndex - 1][codeIdx] : "", expectedCount);
    const currentStatuses = splitCellValues(statusIdx !== -1 ? values[rowIndex - 1][statusIdx] : "", expectedCount);
    const currentObs = splitCellValues(obsIdx !== -1 ? values[rowIndex - 1][obsIdx] : "", expectedCount);
    
    const invValues = invSheet.getDataRange().getValues();
    const invHeaders = invValues[0].map(normalizeHeader);
    const { nameIdx: invNameIdx, totalIdx: invTotalIdx, availableIdx: invDispIdx } = getInventoryHeaderIndices(invHeaders);
    
    let updatedRowsCount = 0;
    const matchedIndices = new Set();
    const extraItems = [];
    
    // First, process matches for existing requested items
    items.forEach(payloadItem => {
      let matched = false;
      for (let k = 0; k < expectedCount; k++) {
        if (!matchedIndices.has(k) && currentItems[k].toUpperCase().trim() === payloadItem.name.toUpperCase().trim() && (currentStatuses[k] === "Solicitado" || currentStatuses[k] === "")) {
          currentCodes[k] = payloadItem.code;
          currentStatuses[k] = "Retirado";
          currentObs[k] = payloadItem.obs || "";
          matchedIndices.add(k);
          matched = true;
          updatedRowsCount++;
          break;
        }
      }
      if (!matched) {
        // It's a new item added on the fly!
        extraItems.push(payloadItem);
      }
    });
    
    // Second, process requested items that were NOT delivered (cancelled/excluded in modal)
    const newlyCancelledItems = [];
    for (let k = 0; k < expectedCount; k++) {
      if (!matchedIndices.has(k) && (currentStatuses[k] === "Solicitado" || currentStatuses[k] === "")) {
        const itemNameToCancel = currentItems[k];
        currentStatuses[k] = "Anulado";
        currentCodes[k] = "Anulado";
        matchedIndices.add(k);
        updatedRowsCount++;
        newlyCancelledItems.push(itemNameToCancel);
        
        // Devolver stock al Inventario (incrementar disponible +1)
        for (let j = 1; j < invValues.length; j++) {
          if (invValues[j][invNameIdx].toString().toUpperCase().trim() === itemNameToCancel.toUpperCase().trim()) {
            const totalVal = parseInt(invValues[j][invTotalIdx]) || 0;
            const rawDisp = invValues[j][invDispIdx];
            const currentDisp = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
            const newDisp = Math.min(totalVal, currentDisp + 1);
            invSheet.getRange(j + 1, invDispIdx + 1).setValue(newDisp);
            invValues[j - 1][invDispIdx] = newDisp; // Sync memory array
            break;
          }
        }
      }
    }
    
    // Third, process extra items added on the fly
    for (let k = 0; k < extraItems.length; k++) {
      const extraItem = extraItems[k];
      
      let invRowIndex = -1;
      let availableDisp = 0;
      for (let j = 1; j < invValues.length; j++) {
        if (invValues[j][invNameIdx].toString().toUpperCase().trim() === extraItem.name.toUpperCase().trim()) {
          const totalVal = parseInt(invValues[j][invTotalIdx]) || 0;
          const rawDisp = invValues[j][invDispIdx];
          const disponible = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
          
          if (disponible > 0) {
            invRowIndex = j + 1;
            availableDisp = disponible;
            break;
          }
        }
      }
      
      if (invRowIndex !== -1) {
        invSheet.getRange(invRowIndex, invDispIdx + 1).setValue(availableDisp - 1);
        invValues[invRowIndex - 1][invDispIdx] = availableDisp - 1; // Sync memory array
      } else {
        throw new Error("No hay stock disponible para agregar '" + extraItem.name + "'.");
      }
      
      currentItems.push(extraItem.name);
      currentCodes.push(extraItem.code);
      currentStatuses.push("Retirado");
      currentObs.push(extraItem.obs || "");
      updatedRowsCount++;
    }
    
    if (updatedRowsCount === 0) throw new Error("No se encontraron registros de solicitud pendientes de entrega.");
    
    // Save back
    let overallStatus = "Solicitado";
    if (currentStatuses.every(s => s === "Anulado")) {
      overallStatus = "Anulado";
    } else if (currentStatuses.some(s => s === "Retirado")) {
      overallStatus = "Retirado";
    }
    
    if (itemIdx !== -1) loanSheet.getRange(rowIndex, itemIdx + 1).setValue(currentItems.join("\n"));
    if (codeIdx !== -1) loanSheet.getRange(rowIndex, codeIdx + 1).setValue(currentCodes.join("\n"));
    if (statusIdx !== -1) loanSheet.getRange(rowIndex, statusIdx + 1).setValue(currentStatuses.join("\n"));
    if (obsIdx !== -1) loanSheet.getRange(rowIndex, obsIdx + 1).setValue(currentObs.join("\n"));
    
    if (dateDeliverIdx !== -1) {
      loanSheet.getRange(rowIndex, dateDeliverIdx + 1).setValue(timestamp); // Date of delivery in dedicated column
    } else if (dateOutIdx !== -1) {
      loanSheet.getRange(rowIndex, dateOutIdx + 1).setValue(timestamp); // Fallback: overwrite request date
    }
    SpreadsheetApp.flush();
    
    const extrasInfo = extraItems.length > 0 ? extraItems.map(function(x) { return x.name; }).join(", ") : "";
    const cancInfo = newlyCancelledItems.length > 0 ? newlyCancelledItems.join(", ") : "";
    let msg = "Retiro físico confirmado.";
    if (items.length > 0) {
      msg += " Entregados: " + items.map(function(x) { return x.name; }).join(", ");
    }
    if (cancInfo) {
      msg += " (Anulados y reincorporados al stock: " + cancInfo + ")";
    }
    
    if (items.length > 0) {
      try {
        sendRetiroFisicoEmail(studentName, studentEmail, items, timestamp, loanId);
      } catch (emailError) {
        Logger.log("ERROR al enviar email de retiro: " + emailError.toString());
        return { status: "success", message: msg + ". (Nota: No se pudo enviar el correo de comprobación)." };
      }
    } else if (newlyCancelledItems.length > 0) {
      try {
        sendAnulacionEmail(studentName, studentEmail, newlyCancelledItems.map(function(x) { return { name: x, code: "Anulado" }; }), timestamp, loanId);
      } catch (emailError) {
        Logger.log("ERROR al enviar email de anulación: " + emailError.toString());
        return { status: "success", message: msg + ". (Nota: No se pudo enviar el correo de notificación)." };
      }
    }
    return { status: "success", message: msg };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function executeReturnLoan(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const invSheet = ss.getSheetByName("Inventario");
    const loanSheet = ss.getSheetByName("Préstamos");
    
    const loanId = payload.loanId;
    const timestamp = payload.timestamp;
    
    const loanValues = loanSheet.getDataRange().getValues();
    const loanHeaders = loanValues[0].map(normalizeHeader);
    const { idIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateInIdx, statusIdx, progDevolucionIdx, daysOverdueIdx, obsReturnIdx, obsIdx } = getLoanHeaderIndices(loanHeaders);
    
    const invValues = invSheet.getDataRange().getValues();
    const invHeaders = invValues[0].map(normalizeHeader);
    const { nameIdx: invNameIdx, totalIdx: invTotalIdx, availableIdx: invDispIdx, codesIdx: invCodesIdx } = getInventoryHeaderIndices(invHeaders);
    
    let rowIndex = -1;
    for (let i = 1; i < loanValues.length; i++) {
      if (loanValues[i][idIdx].toString() === loanId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("No se encontró el préstamo con ID " + loanId);
    
    const studentName = loanValues[rowIndex - 1][nameIdx].toString();
    const studentEmail = emailIdx !== -1 ? loanValues[rowIndex - 1][emailIdx].toString() : "";
    
    const currentItems = loanValues[rowIndex - 1][itemIdx].toString().split("\n").map(x => x.trim()).filter(Boolean);
    const expectedCount = currentItems.length;
    
    const currentCodes = splitCellValues(codeIdx !== -1 ? loanValues[rowIndex - 1][codeIdx] : "", expectedCount);
    
    const rawStatus = statusIdx !== -1 ? loanValues[rowIndex - 1][statusIdx].toString().trim() : "Solicitado";
    const currentStatuses = rawStatus.includes("\n") ? splitCellValues(rawStatus, expectedCount) : new Array(expectedCount).fill(rawStatus);
    
    const rawDateIn = dateInIdx !== -1 && loanValues[rowIndex - 1][dateInIdx] ? loanValues[rowIndex - 1][dateInIdx] : "";
    let currentDatesIn = [];
    if (rawDateIn instanceof Date) {
      currentDatesIn = new Array(expectedCount).fill(rawDateIn);
    } else if (rawDateIn.toString().includes("\n")) {
      currentDatesIn = splitCellValues(rawDateIn.toString(), expectedCount);
    } else {
      currentDatesIn = new Array(expectedCount).fill(rawDateIn);
    }
    
    const returnedItems = [];
    
    for (let k = 0; k < expectedCount; k++) {
      if (currentStatuses[k] === "Retirado") {
        currentStatuses[k] = "Devuelto";
        currentDatesIn[k] = timestamp;
        
        const itemName = currentItems[k];
        const itemCode = currentCodes[k] || "";
        
        returnedItems.push({ name: itemName, code: itemCode });
        
        // Restituir stock en inventario
        let invRowIndex = -1;
        for (let j = 1; j < invValues.length; j++) {
          if (invValues[j][invNameIdx].toString() === itemName) {
            const codesCell = invCodesIdx !== -1 ? invValues[j][invCodesIdx].toString() : "";
            const rowCodes = codesCell.split(",").map(c => c.trim().toUpperCase());
            
            if (itemCode && rowCodes.includes(itemCode.toUpperCase().trim())) {
              invRowIndex = j + 1;
              break;
            } else if (invRowIndex === -1) {
              invRowIndex = j + 1;
            }
          }
        }
        
        if (invRowIndex !== -1) {
          const rawDisp = invValues[invRowIndex - 1][invDispIdx];
          const totalVal = parseInt(invValues[invRowIndex - 1][invTotalIdx]) || 0;
          const currentDisp = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
          const newDisp = Math.min(totalVal, currentDisp + 1);
          
          invSheet.getRange(invRowIndex, invDispIdx + 1).setValue(newDisp);
          invValues[invRowIndex - 1][invDispIdx] = newDisp; // Sync in-memory array for next iterations
        }
      }
    }
    
    if (returnedItems.length === 0) throw new Error("No se encontraron registros de préstamo activos para devolver.");
    
    // Save back
    const formattedDatesIn = currentDatesIn.map(function(d) { return d ? formatDate(d) : ""; }).join("\n");
    if (statusIdx !== -1) loanSheet.getRange(rowIndex, statusIdx + 1).setValue(currentStatuses.join("\n"));
    if (dateInIdx !== -1) loanSheet.getRange(rowIndex, dateInIdx + 1).setValue(formattedDatesIn);
    
    // Guardar observaciones de devolución
    if (obsReturnIdx !== -1) {
      loanSheet.getRange(rowIndex, obsReturnIdx + 1).setValue(payload.obsReturn || "");
    } else if (obsIdx !== -1 && payload.obsReturn) {
      const existingObs = loanSheet.getRange(rowIndex, obsIdx + 1).getValue().toString().trim();
      const newObs = existingObs ? existingObs + " | Dev: " + payload.obsReturn : "Dev: " + payload.obsReturn;
      loanSheet.getRange(rowIndex, obsIdx + 1).setValue(newObs);
    }
    
    // Calcular y guardar días de atraso en columna histórica si existe
    if (daysOverdueIdx !== -1) {
      let daysOverdue = 0;
      const progDevolucionVal = progDevolucionIdx !== -1 ? loanValues[rowIndex - 1][progDevolucionIdx] : "";
      if (progDevolucionVal) {
        const progDate = parseDateString(progDevolucionVal);
        const realDate = parseDateString(timestamp);
        
        if (progDate && realDate) {
          progDate.setHours(0,0,0,0);
          realDate.setHours(0,0,0,0);
          
          const diffTime = realDate.getTime() - progDate.getTime();
          const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
          if (diffDays > 0) {
            daysOverdue = diffDays;
          }
        }
      }
      loanSheet.getRange(rowIndex, daysOverdueIdx + 1).setValue(daysOverdue);
    }
    SpreadsheetApp.flush();
    
    const returnedInfo = returnedItems.map(function(x) { return x.name + " (" + x.code + ")"; }).join(" | ");
    const msg = "Devolución procesada. Equipos: " + returnedInfo;
    
    try {
      sendDevolucionEmail(studentName, studentEmail, returnedItems, timestamp, loanId);
    } catch (emailError) {
      Logger.log("ERROR al enviar email de devolución: " + emailError.toString());
      return { status: "success", message: msg + ". (Nota: No se pudo enviar el correo de comprobación)." };
    }
    return { status: "success", message: msg };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function executeCancelLoan(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const invSheet = ss.getSheetByName("Inventario");
    const loanSheet = ss.getSheetByName("Préstamos");
    
    const loanId = payload.loanId;
    const targetItemName = payload.itemName; // Opcional: para anular un ítem específico de la solicitud
    const timestamp = payload.timestamp;
    
    const loanValues = loanSheet.getDataRange().getValues();
    const loanHeaders = loanValues[0].map(normalizeHeader);
    const { idIdx, nameIdx, emailIdx, itemIdx, codeIdx, dateInIdx, statusIdx } = getLoanHeaderIndices(loanHeaders);
    
    const invValues = invSheet.getDataRange().getValues();
    const invHeaders = invValues[0].map(normalizeHeader);
    const { nameIdx: invNameIdx, totalIdx: invTotalIdx, availableIdx: invDispIdx, codesIdx: invCodesIdx } = getInventoryHeaderIndices(invHeaders);
    
    let rowIndex = -1;
    for (let i = 1; i < loanValues.length; i++) {
      if (loanValues[i][idIdx].toString() === loanId) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("No se encontró el préstamo con ID " + loanId);
    
    const studentName = loanValues[rowIndex - 1][nameIdx].toString();
    const studentEmail = emailIdx !== -1 ? loanValues[rowIndex - 1][emailIdx].toString() : "";
    
    const currentItems = loanValues[rowIndex - 1][itemIdx].toString().split("\n").map(x => x.trim()).filter(Boolean);
    const expectedCount = currentItems.length;
    
    const currentCodes = splitCellValues(codeIdx !== -1 ? loanValues[rowIndex - 1][codeIdx] : "", expectedCount);
    
    const rawStatus = statusIdx !== -1 ? loanValues[rowIndex - 1][statusIdx].toString().trim() : "Solicitado";
    const currentStatuses = rawStatus.includes("\n") ? splitCellValues(rawStatus, expectedCount) : new Array(expectedCount).fill(rawStatus);
    
    const cancelledItems = [];
    
    for (let k = 0; k < expectedCount; k++) {
      const currentStatus = currentStatuses[k];
      if (currentStatus === "Solicitado" || currentStatus === "Retirado" || currentStatus === "") {
        const itemName = currentItems[k];
        
        if (targetItemName && itemName.toUpperCase().trim() !== targetItemName.toUpperCase().trim()) {
          continue;
        }
        
        const itemCode = currentCodes[k] || "";
        currentStatuses[k] = "Anulado";
        if (!itemCode || itemCode.toUpperCase().trim() === "PTE. ENTREGA") {
          currentCodes[k] = "Anulado";
        }
        
        let invRowIndex = -1;
        for (let j = 1; j < invValues.length; j++) {
          if (invValues[j][invNameIdx].toString() === itemName) {
            const codesCell = invCodesIdx !== -1 ? invValues[j][invCodesIdx].toString() : "";
            const rowCodes = codesCell.split(",").map(c => c.trim().toUpperCase());
            
            if (itemCode && rowCodes.includes(itemCode.toUpperCase().trim())) {
              invRowIndex = j + 1;
              break;
            } else if (invRowIndex === -1) {
              invRowIndex = j + 1;
            }
          }
        }
        
        if (invRowIndex !== -1) {
          const rawDisp = invValues[invRowIndex - 1][invDispIdx];
          const totalVal = parseInt(invValues[invRowIndex - 1][invTotalIdx]) || 0;
          const currentDisp = (rawDisp === "" || rawDisp === undefined || rawDisp === null) ? totalVal : (parseInt(rawDisp) || 0);
          const newDisp = Math.min(totalVal, currentDisp + 1);
          
          invSheet.getRange(invRowIndex, invDispIdx + 1).setValue(newDisp);
          invValues[invRowIndex - 1][invDispIdx] = newDisp; // Sync in-memory array for next iterations
        }
        
        cancelledItems.push({ name: itemName, code: itemCode });
      }
    }
    
    if (cancelledItems.length === 0) throw new Error("No se encontraron solicitudes o préstamos activos para anular.");
    
    // Save back
    let overallStatus = "Solicitado";
    if (currentStatuses.every(s => s === "Anulado")) {
      overallStatus = "Anulado";
    } else if (currentStatuses.some(s => s === "Retirado")) {
      overallStatus = "Retirado";
    } else if (currentStatuses.some(s => s === "Devuelto")) {
      overallStatus = "Devuelto";
    }
    
    if (statusIdx !== -1) loanSheet.getRange(rowIndex, statusIdx + 1).setValue(currentStatuses.join("\n"));
    if (codeIdx !== -1) loanSheet.getRange(rowIndex, codeIdx + 1).setValue(currentCodes.join("\n"));
    if (overallStatus === "Anulado" && dateInIdx !== -1) {
      loanSheet.getRange(rowIndex, dateInIdx + 1).setValue("");
    }
    SpreadsheetApp.flush();
    
    try {
      sendAnulacionEmail(studentName, studentEmail, cancelledItems, timestamp, loanId);
    } catch (emailError) {
      Logger.log("ERROR al enviar email de anulación: " + emailError.toString());
      return { status: "success", message: "Préstamo anulado. (Nota: No se pudo enviar el correo de notificación)." };
    }
    return { status: "success", message: "Préstamo anulado." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// CONFIGURACIÓN DE REGISTROS DINÁMICOS
// ==========================================

function executeAddEquipment(ss, payload) {
  const sheet = ss.getSheetByName("Inventario");
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  const { categoryIdx, nameIdx, totalIdx, availableIdx, codesIdx, imageIdx, descriptionIdx } = getInventoryHeaderIndices(headers);
  
  const cleanCat = cleanCategoryGAS(payload.category);
  const newRow = new Array(headers.length).fill("");
  if (categoryIdx !== -1) newRow[categoryIdx] = cleanCat;
  if (nameIdx !== -1) newRow[nameIdx] = payload.name;
  if (totalIdx !== -1) newRow[totalIdx] = payload.total;
  if (availableIdx !== -1) newRow[availableIdx] = payload.total;
  if (codesIdx !== -1) newRow[codesIdx] = payload.codes.join(", ");
  if (imageIdx !== -1) newRow[imageIdx] = payload.image || "";
  if (descriptionIdx !== -1) newRow[descriptionIdx] = payload.description || "";
  
  sheet.appendRow(newRow);

  // Si existe la pestaña Categorías, verificar si esta categoría ya está registrada; si no, agregarla
  try {
    const catSheet = ss.getSheetByName("Categorías") || ss.getSheetByName("Categorias");
    if (catSheet && cleanCat) {
      const catVals = catSheet.getDataRange().getValues();
      const normNew = cleanCat.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      let exists = false;
      for (let c = 1; c < catVals.length; c++) {
        const valNorm = catVals[c][0] ? catVals[c][0].toString().trim().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "") : "";
        if (valNorm === normNew) {
          exists = true;
          break;
        }
      }
      if (!exists) {
        catSheet.appendRow([cleanCat]);
      }
    }
  } catch (err) {
    Logger.log("Aviso al actualizar Categorías: " + err.toString());
  }

  return { status: "success", message: "Equipo guardado con éxito." };
}

function executeAddCategory(ss, payload) {
  let catSheet = ss.getSheetByName("Categorías") || ss.getSheetByName("Categorias");
  if (!catSheet) {
    catSheet = ss.insertSheet("Categorías");
    catSheet.appendRow(["Categoría"]);
  }
  const newCat = (payload.category || "").toString().trim();
  if (!newCat) throw new Error("Nombre de categoría no válido.");
  
  const values = catSheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] && values[i][0].toString().trim().toLowerCase() === newCat.toLowerCase()) {
      return { status: "success", message: "La categoría ya existe." };
    }
  }
  catSheet.appendRow([newCat]);
  return { status: "success", message: "Categoría registrada con éxito." };
}

function executeDeleteEquipment(ss, payload) {
  const sheet = ss.getSheetByName("Inventario");
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  const { nameIdx } = getInventoryHeaderIndices(headers);
  
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    if (values[i][nameIdx].toString() === payload.name) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Equipo no encontrado.");
  sheet.deleteRow(rowIndex);
  return { status: "success", message: "Eliminado." };
}

function executeAddStudent(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const sheet = ss.getSheetByName("Alumnos");
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(normalizeHeader);
    const { rutIdx, nameIdx, lastnameIdx, fonoIdx, emailIdx, obsIdx, blockDateIdx } = getStudentHeaderIndices(headers);
    
    const cleanSearch = payload.rut.replace(/[^0-9kK]/g, '').toLowerCase();
    for (let i = 1; i < values.length; i++) {
      const cleanStudentRut = values[i][rutIdx].toString().replace(/[^0-9kK]/g, '').toLowerCase();
      if (cleanStudentRut === cleanSearch) {
        throw new Error("El alumno con el RUT " + payload.rut + " ya está registrado.");
      }
    }
    
    const newRow = new Array(headers.length).fill("");
    if (rutIdx !== -1) newRow[rutIdx] = payload.rut;
    if (nameIdx !== -1) newRow[nameIdx] = payload.name;
    if (lastnameIdx !== -1) newRow[lastnameIdx] = payload.lastname;
    if (fonoIdx !== -1) newRow[fonoIdx] = payload.fono;
    if (emailIdx !== -1) newRow[emailIdx] = payload.email;
    if (obsIdx !== -1) newRow[obsIdx] = payload.status === "Bloqueado" ? payload.debt : "";
    if (blockDateIdx !== -1) {
      const now = new Date();
      const defaultDate = Utilities.formatDate(now, Session.getScriptTimeZone() || "America/Santiago", "yyyy-MM-dd HH:mm:ss");
      newRow[blockDateIdx] = payload.status === "Bloqueado" ? (payload.blockDate || defaultDate) : "";
    }
    
    sheet.appendRow(newRow);
    return { status: "success", message: "Alumno guardado." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function executeDeleteStudent(ss, payload) {
  const sheet = ss.getSheetByName("Alumnos");
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(normalizeHeader);
  const { rutIdx } = getStudentHeaderIndices(headers);
  
  const cleanSearch = payload.rut.replace(/[^0-9kK]/g, '').toLowerCase();
  let rowIndex = -1;
  for (let i = 1; i < values.length; i++) {
    const cleanStudentRut = values[i][rutIdx].toString().replace(/[^0-9kK]/g, '').toLowerCase();
    if (cleanStudentRut === cleanSearch) {
      rowIndex = i + 1;
      break;
    }
  }
  if (rowIndex === -1) throw new Error("Alumno no encontrado.");
  sheet.deleteRow(rowIndex);
  return { status: "success", message: "Eliminado." };
}

function executeApplySanction(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const sheet = ss.getSheetByName("Alumnos");
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(normalizeHeader);
    const { rutIdx, obsIdx, blockDateIdx } = getStudentHeaderIndices(headers);
    
    if (rutIdx === -1 || obsIdx === -1) {
      throw new Error("No se encontraron las columnas de RUT u Observaciones en la pestaña Alumnos.");
    }
    
    const cleanSearch = payload.rut.replace(/[^0-9kK]/g, '').toLowerCase();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const stRutVal = values[i][rutIdx];
      const cleanStudentRut = stRutVal ? stRutVal.toString().replace(/[^0-9kK]/g, '').toLowerCase() : "";
      if (cleanStudentRut === cleanSearch) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Alumno no encontrado.");
    
    const reasonText = payload.reason ? payload.reason.trim() : "Sancionado por el administrador";
    sheet.getRange(rowIndex, obsIdx + 1).setValue(reasonText);
    
    if (blockDateIdx !== -1) {
      const now = new Date();
      const dateStr = payload.date || Utilities.formatDate(now, Session.getScriptTimeZone() || "America/Santiago", "yyyy-MM-dd HH:mm:ss");
      sheet.getRange(rowIndex, blockDateIdx + 1).setValue(dateStr);
    }
    
    SpreadsheetApp.flush();
    return { status: "success", message: "Alumno bloqueado con éxito." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function executeRemoveSanction(ss, payload) {
  const lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) {
    return { status: "error", message: "Servidor ocupado." };
  }
  
  try {
    const sheet = ss.getSheetByName("Alumnos");
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(normalizeHeader);
    const { rutIdx, obsIdx, blockDateIdx } = getStudentHeaderIndices(headers);
    
    if (rutIdx === -1 || obsIdx === -1) {
      throw new Error("No se encontraron las columnas de RUT u Observaciones en la pestaña Alumnos.");
    }
    
    const cleanSearch = payload.rut.replace(/[^0-9kK]/g, '').toLowerCase();
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      const stRutVal = values[i][rutIdx];
      const cleanStudentRut = stRutVal ? stRutVal.toString().replace(/[^0-9kK]/g, '').toLowerCase() : "";
      if (cleanStudentRut === cleanSearch) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) throw new Error("Alumno no encontrado.");
    
    sheet.getRange(rowIndex, obsIdx + 1).setValue("");
    if (blockDateIdx !== -1) {
      sheet.getRange(rowIndex, blockDateIdx + 1).setValue("");
    }
    
    SpreadsheetApp.flush();
    return { status: "success", message: "Bloqueo levantado. Alumno ahora está Activo." };
  } catch (error) {
    return { status: "error", message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// CORREOS ELECTRÓNICOS
// ==========================================

function sendSolicitudEmail(student, items, timestamp, progRetiro, progDevolucion, subject, loanId) {
  if (!student || !student.email || student.email.trim() === "") {
    throw new Error("No hay destinatario registrado.");
  }
  let itemsHtml = "";
  items.forEach(item => {
    itemsHtml += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 12px; font-weight: bold; color: #1e293b;">${item.name}</td>
        <td style="padding: 12px; color: #475569;">${item.category}</td>
      </tr>
    `;
  });
  
  const studentFullName = student.name + " " + (student.lastname || "");
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05); border: 1px solid #e2e8f0;">
        <div style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); padding: 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Comprobante de Solicitud</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 13px;">Reserva ID: ${loanId}</p>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 15px; color: #1e293b; margin-top: 0;">Estimado(a) <strong>${studentFullName}</strong>,</p>
          <p style="font-size: 14px; color: #475569; line-height: 1.5;">Hemos recibido tu solicitud de reserva realizada el <strong>${timestamp}</strong>:</p>
          
          <table style="width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 13px;">
            <thead>
              <tr style="background-color: #f8fafc; border-bottom: 2px solid #e2e8f0; text-align: left; color: #475569; font-weight: bold;">
                <th style="padding: 12px;">Equipo</th>
                <th style="padding: 12px;">Categoría</th>
              </tr>
            </thead>
            <tbody>
              ${itemsHtml}
            </tbody>
          </table>
          
          <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-top: 15px; font-size: 13px;">
            <table style="width: 100%; border-collapse: collapse;">
              <tr>
                <td style="padding: 6px 0; color: #64748b; width: 45%;"><strong>Fecha Retiro:</strong></td>
                <td style="padding: 6px 0; color: #1e293b; font-weight: bold;">${progRetiro || timestamp}</td>
              </tr>
              <tr>
                <td style="padding: 6px 0; color: #64748b;"><strong>Fecha Devolución:</strong></td>
                <td style="padding: 6px 0; color: #1e293b; font-weight: bold;">${progDevolucion || "-"}</td>
              </tr>
            </table>
          </div>
          
          <div style="margin-top: 20px; background-color: #fffbeb; border: 1px solid #fef3c7; border-radius: 8px; padding: 15px; font-size: 13px;">
            <h4 style="margin: 0 0 5px 0; color: #d97706; font-weight: bold;">Reserva Temporal (ID: ${loanId}):</h4>
            <p style="margin: 0; color: #92400e; line-height: 1.45;">Tienes un plazo máximo de <strong>12 horas</strong> desde el envío de esta solicitud para presentarte en el pañol AVP y realizar el retiro físico. Pasado este plazo, tu reserva expirará automáticamente.</p>
          </div>
        </div>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: student.email,
    subject: "Comprobante de Solicitud - ID: " + loanId,
    htmlBody: htmlBody
  });
}

function sendRetiroFisicoEmail(studentName, studentEmail, items, timestamp, loanId) {
  if (!studentEmail || studentEmail.trim() === "") {
    throw new Error("No hay destinatario registrado.");
  }
  
  let itemsTableRows = "";
  items.forEach(item => {
    itemsTableRows += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 10px 0; color: #1e293b; font-weight: bold;">${item.name}</td>
        <td style="padding: 10px 0; color: #4f46e5; font-weight: bold; text-align: right;"><code>${item.code}</code></td>
      </tr>
    `;
  });

  let obsText = "";
  const firstItemWithObs = items.find(item => item.obs && item.obs.trim() !== "");
  if (firstItemWithObs) {
    obsText = `
      <div style="margin-top: 15px; background-color: #fffbeb; border: 1px solid #fef3c7; border-radius: 8px; padding: 15px;">
        <h4 style="margin: 0 0 5px 0; color: #b45309; font-size: 13px; font-weight: 800;">Observaciones consignadas:</h4>
        <p style="margin: 0; color: #78350f; font-size: 12px; font-style: italic; line-height: 1.4;">${firstItemWithObs.obs}</p>
      </div>
    `;
  }
  
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05); border: 1px solid #e2e8f0;">
        <div style="background: linear-gradient(135deg, #4f46e5 0%, #7c3aed 100%); padding: 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Retiro Físico Confirmado</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 13px;">ID Préstamo: ${loanId}</p>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 15px; color: #1e293b; margin-top: 0;">Estimado(a) <strong>${studentName}</strong>,</p>
          <p style="font-size: 14px; color: #475569; line-height: 1.5;">Confirmamos la entrega física de los equipos el día <strong>${timestamp}</strong>:</p>
          
          <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-top: 15px;">
            <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
              <thead>
                <tr style="border-bottom: 2px solid #e2e8f0;">
                  <th style="text-align: left; padding-bottom: 8px; color: #64748b;">Equipo</th>
                  <th style="text-align: right; padding-bottom: 8px; color: #64748b;">Nº Inventario</th>
                </tr>
              </thead>
              <tbody>
                ${itemsTableRows}
              </tbody>
            </table>
          </div>
          ${obsText}
          
          <div style="margin-top: 20px; background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 15px; font-size: 13px;">
            <h4 style="margin: 0 0 5px 0; color: #166534; font-weight: bold;">Recordatorio del Protocolo (Préstamo: ${loanId}):</h4>
            <p style="margin: 0; color: #14532d; line-height: 1.45;">Eres legalmente responsable de la custodia, integridad y retorno de este equipo UFRO en la fecha programada. Cualquier daño debe ser reportado inmediatamente a la oficina audiovisual.</p>
          </div>
        </div>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: studentEmail,
    subject: "Retiro Físico Confirmado - ID: " + loanId,
    htmlBody: htmlBody
  });
}

function sendDevolucionEmail(studentName, studentEmail, items, timestamp, loanId) {
  if (!studentEmail || studentEmail.trim() === "") {
    throw new Error("No hay destinatario registrado.");
  }
  
  let itemsTableRows = "";
  items.forEach(item => {
    itemsTableRows += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 10px 0; color: #1e293b; font-weight: bold;">${item.name}</td>
        <td style="padding: 10px 0; color: #64748b; text-align: right;"><code>${item.code}</code></td>
      </tr>
    `;
  });
  
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05); border: 1px solid #e2e8f0;">
        <div style="background: linear-gradient(135deg, #10b981 0%, #059669 100%); padding: 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Devolución Exitosa</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 13px;">ID Préstamo: ${loanId}</p>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 15px; color: #1e293b; margin-top: 0;">Estimado(a) <strong>${studentName}</strong>,</p>
          <p style="font-size: 14px; color: #475569; line-height: 1.5;">Confirmamos la devolución física de los equipos realizada el <strong>${timestamp}</strong>:</p>
          
          <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-top: 15px;">
            <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
              <thead>
                <tr style="border-bottom: 2px solid #e2e8f0;">
                  <th style="text-align: left; padding-bottom: 8px; color: #64748b;">Equipo</th>
                  <th style="text-align: right; padding-bottom: 8px; color: #64748b;">Nº Inventario</th>
                </tr>
              </thead>
              <tbody>
                ${itemsTableRows}
              </tbody>
            </table>
          </div>
          
          <div style="margin-top: 20px; background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 15px; font-size: 13px;">
            <h4 style="margin: 0 0 5px 0; color: #166534; font-weight: bold;">Devolución Exitosa (Préstamo: ${loanId}):</h4>
            <p style="margin: 0; color: #14532d; line-height: 1.45;">El equipo ha sido verificado físicamente y reingresado al stock activo. Tu estado de préstamos se encuentra al día.</p>
          </div>
        </div>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: studentEmail,
    subject: "Devolución Exitosa - ID: " + loanId,
    htmlBody: htmlBody
  });
}

function sendAnulacionEmail(studentName, studentEmail, items, timestamp, loanId) {
  if (!studentEmail || studentEmail.trim() === "") {
    throw new Error("No hay destinatario registrado.");
  }
  
  let itemsTableRows = "";
  items.forEach(item => {
    itemsTableRows += `
      <tr style="border-bottom: 1px solid #e2e8f0;">
        <td style="padding: 10px 0; color: #1e293b; font-weight: bold;">${item.name}</td>
        <td style="padding: 10px 0; color: #64748b; text-align: right;"><code>${item.code}</code></td>
      </tr>
    `;
  });
  
  const htmlBody = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px;">
      <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05); border: 1px solid #e2e8f0;">
        <div style="background: linear-gradient(135deg, #ef4444 0%, #dc2626 100%); padding: 30px; text-align: center; color: white;">
          <h1 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Préstamo Anulado</h1>
          <p style="margin: 5px 0 0 0; opacity: 0.9; font-size: 13px;">ID Préstamo: ${loanId}</p>
        </div>
        <div style="padding: 30px;">
          <p style="font-size: 15px; color: #1e293b; margin-top: 0;">Estimado(a) <strong>${studentName}</strong>,</p>
          <p style="font-size: 14px; color: #475569; line-height: 1.5;">Tu préstamo con ID <strong>${loanId}</strong> ha sido <strong>Anulado</strong> el día <strong>${timestamp}</strong>:</p>
          
          <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 20px; margin-top: 15px;">
            <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
              <thead>
                <tr style="border-bottom: 2px solid #e2e8f0;">
                  <th style="text-align: left; padding-bottom: 8px; color: #64748b;">Equipo</th>
                  <th style="text-align: right; padding-bottom: 8px; color: #64748b;">Nº Inventario</th>
                </tr>
              </thead>
              <tbody>
                ${itemsTableRows}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  `;
  
  MailApp.sendEmail({
    to: studentEmail,
    subject: "Préstamo Anulado - ID: " + loanId,
    htmlBody: htmlBody
  });
}

function executeSendContactEmail(ss, payload) {
  try {
    const studentEmail = payload.studentEmail ? payload.studentEmail.trim() : "";
    const studentName = payload.studentName ? payload.studentName.trim() : "Estudiante";
    const loanId = payload.loanId ? payload.loanId.trim() : "Préstamo";
    const subject = payload.subject ? payload.subject.trim() : ("Notificación Oficial - Pañol AVP UFRO - ID: " + loanId);
    const message = payload.message ? payload.message.trim() : "";
    const items = payload.items || [];
    const progRetiro = payload.progRetiro || "";
    const progDevolucion = payload.progDevolucion || "";
    const subjectName = payload.subjectName || "";
    const loanStatus = payload.loanStatus || "Activo";

    if (!studentEmail) {
      return { status: "error", message: "El estudiante no tiene un correo electrónico válido registrado." };
    }
    if (!message) {
      return { status: "error", message: "El mensaje no puede estar vacío." };
    }

    let itemsTableRows = "";
    if (items.length > 0) {
      items.forEach(item => {
        const itemCode = (item.code && item.code.trim() !== "") ? item.code : "Pte. Entrega";
        itemsTableRows += `
          <tr style="border-bottom: 1px solid #e2e8f0;">
            <td style="padding: 10px 12px; color: #1e293b; font-weight: bold;">${item.name}</td>
            <td style="padding: 10px 12px; color: #0284c7; font-weight: bold; text-align: right;"><code>${itemCode}</code></td>
          </tr>
        `;
      });
    } else {
      itemsTableRows = `
        <tr style="border-bottom: 1px solid #e2e8f0;">
          <td colspan="2" style="padding: 10px 12px; color: #64748b; text-align: center; font-style: italic;">Sin equipos asociados registrados</td>
        </tr>
      `;
    }

    const formattedMessage = message
      .split('&').join('&amp;')
      .split('<').join('&lt;')
      .split('>').join('&gt;')
      .split('"').join('&quot;')
      .split('\n').join('<br>');

    const htmlBody = `
      <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8fafc; padding: 30px;">
        <div style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 15px rgba(15, 23, 42, 0.05); border: 1px solid #e2e8f0;">
          <div style="background: linear-gradient(135deg, #0284c7 0%, #2563eb 100%); padding: 28px 30px; text-align: center; color: white;">
            <h1 style="margin: 0; font-size: 22px; font-weight: 800; letter-spacing: 0.5px;">Comunicación del Pañol AVP</h1>
            <p style="margin: 5px 0 0 0; opacity: 0.95; font-size: 13px; font-weight: 500;">Préstamo ID: ${loanId} • Carrera de Periodismo UFRO</p>
          </div>
          
          <div style="padding: 30px;">
            <p style="font-size: 15px; color: #1e293b; margin-top: 0;">Estimado(a) <strong>${studentName}</strong>,</p>
            <p style="font-size: 14px; color: #475569; line-height: 1.5; margin-bottom: 18px;">
              Te contactamos desde el <strong>Pañol Audiovisual (AVP)</strong> en relación al registro de préstamo con identificador <strong>${loanId}</strong>:
            </p>
            
            <div style="background-color: #eff6ff; border-left: 4px solid #0284c7; border-radius: 6px; padding: 16px 20px; margin-bottom: 24px;">
              <h4 style="margin: 0 0 8px 0; color: #0369a1; font-size: 13px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">Mensaje de Administración:</h4>
              <div style="margin: 0; color: #1e3a8a; font-size: 14px; line-height: 1.6;">${formattedMessage}</div>
            </div>

            <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 18px; margin-top: 15px;">
              <h4 style="margin: 0 0 12px 0; color: #475569; font-size: 12px; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px;">Equipos Comprendidos en el Préstamo:</h4>
              <table style="width: 100%; font-size: 13px; border-collapse: collapse;">
                <thead>
                  <tr style="border-bottom: 2px solid #e2e8f0; text-align: left;">
                    <th style="padding: 8px 0; color: #64748b; font-weight: 700;">Equipo</th>
                    <th style="padding: 8px 0; color: #64748b; font-weight: 700; text-align: right;">Nº Inventario</th>
                  </tr>
                </thead>
                <tbody>
                  ${itemsTableRows}
                </tbody>
              </table>
            </div>

            <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 8px; padding: 16px 20px; margin-top: 15px; font-size: 13px;">
              <table style="width: 100%; border-collapse: collapse;">
                <tr>
                  <td style="padding: 5px 0; color: #64748b; width: 45%;"><strong>Estado del Préstamo:</strong></td>
                  <td style="padding: 5px 0; color: #1e293b; font-weight: bold;">${loanStatus}</td>
                </tr>
                ${progRetiro ? `<tr><td style="padding: 5px 0; color: #64748b;"><strong>Fecha Retiro Programada:</strong></td><td style="padding: 5px 0; color: #1e293b; font-weight: bold;">${progRetiro}</td></tr>` : ''}
                ${progDevolucion ? `<tr><td style="padding: 5px 0; color: #64748b;"><strong>Fecha Devolución Programada:</strong></td><td style="padding: 5px 0; color: #1e293b; font-weight: bold;">${progDevolucion}</td></tr>` : ''}
                ${subjectName ? `<tr><td style="padding: 5px 0; color: #64748b;"><strong>Asignatura / Cátedra:</strong></td><td style="padding: 5px 0; color: #1e293b; font-weight: bold;">${subjectName}</td></tr>` : ''}
              </table>
            </div>
            
            <div style="margin-top: 20px; background-color: #fffbeb; border: 1px solid #fef3c7; border-radius: 8px; padding: 15px; font-size: 13px;">
              <h4 style="margin: 0 0 5px 0; color: #b45309; font-weight: bold;">Atención y Respuestas:</h4>
              <p style="margin: 0; color: #92400e; line-height: 1.45;">Si requieres responder a este comunicado o justificar alguna situación particular, puedes responder directamente a este correo o acercarte a la oficina del pañol AVP durante los horarios de atención.</p>
            </div>
            
            <p style="margin-top: 25px; margin-bottom: 0; font-size: 12px; color: #94a3b8; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 15px;">
              Pañol Audiovisual (AVP) • Carrera de Periodismo • Universidad de La Frontera
            </p>
          </div>
        </div>
      </div>
    `;

    MailApp.sendEmail({
      to: studentEmail,
      subject: subject,
      htmlBody: htmlBody
    });

    return { status: "success", message: "Correo oficial enviado con éxito a " + studentEmail };
  } catch (err) {
    Logger.log("Error al enviar correo de contacto: " + err.toString());
    return { status: "error", message: "Error al enviar correo: " + err.toString() };
  }
}

// ==========================================
// FORMATEADOR DE FECHA
// ==========================================
function formatDate(dateObj) {
  if (dateObj instanceof Date) {
    const pad = (n) => n.toString().padStart(2, '0');
    return `${dateObj.getFullYear()}-${pad(dateObj.getMonth() + 1)}-${pad(dateObj.getDate())} ${pad(dateObj.getHours())}:${pad(dateObj.getMinutes())}`;
  }
  return dateObj ? dateObj.toString() : "";
}

function parseDateString(dateStr) {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return dateStr;
  
  const cleanStr = dateStr.toString().trim().split(" ")[0];
  const parts = cleanStr.includes("-") ? cleanStr.split("-") : cleanStr.split("/");
  if (parts.length < 3) return null;
  
  let year, month, day;
  if (parts[0].length === 4) {
    year = parseInt(parts[0], 10);
    month = parseInt(parts[1], 10) - 1;
    day = parseInt(parts[2], 10);
  } else if (parts[2].length === 4) {
    year = parseInt(parts[2], 10);
    month = parseInt(parts[1], 10) - 1;
    day = parseInt(parts[0], 10);
  } else {
    return null;
  }
  
  if (isNaN(year) || isNaN(month) || isNaN(day)) return null;
  return new Date(year, month, day);
}
