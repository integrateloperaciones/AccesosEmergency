const API_URL = "https://script.google.com/macros/s/AKfycbyExK3IUlkLtjAtrLICp0x8ie3ehs414Z_WOtUz3hwZkSK2wGHJ5SuOR8FUIHtor-zR/exec";

const btnAgregarFila = document.getElementById("btnAgregarFila");
const tbodyPersonal = document.getElementById("tbodyPersonal");
const formAcceso = document.getElementById("formAcceso");
const btnLimpiar = document.getElementById("btnLimpiar");
const btnDescargarPlantilla = document.getElementById("btnDescargarPlantilla");
const inputPlantilla = document.getElementById("inputPlantilla");

const buscarRelacion = document.getElementById("buscarRelacion");
const searchResults = document.getElementById("searchResults");

const popupExito = document.getElementById("popupExito");
const btnCerrarPopup = document.getElementById("btnCerrarPopup");

const loaderOverlay = document.getElementById("loaderOverlay");

const horaIngresoInput = document.getElementById("horaIngreso");
const horaSalidaInput = document.getElementById("horaSalida");

let guardando = false;

// =========================
// LISTADO DE PERSONAL
// =========================
btnAgregarFila.addEventListener("click", () => {
  agregarFilaPersonal("", "");
});

function agregarFilaPersonal(nombre = "", dni = "") {
  const tr = document.createElement("tr");

  tr.innerHTML = `
    <td>
      <input type="text" name="nombreFuncionario[]" placeholder="Ingrese nombre completo" value="${escapeHtml(nombre)}" />
    </td>
    <td>
      <input type="text" name="dni[]" placeholder="Ingrese DNI" maxlength="8" value="${escapeHtml(dni)}" />
    </td>
    <td class="text-center">
      <button type="button" class="btn btn-danger btnEliminarFila">Eliminar</button>
    </td>
  `;

  tbodyPersonal.appendChild(tr);
}

tbodyPersonal.addEventListener("click", (e) => {
  if (e.target.classList.contains("btnEliminarFila")) {
    const filas = tbodyPersonal.querySelectorAll("tr");

    if (filas.length === 1) {
      alert("Debe quedar al menos una fila en el listado de personal.");
      return;
    }

    e.target.closest("tr").remove();
  }
});

tbodyPersonal.addEventListener("input", (e) => {
  if (e.target.name === "dni[]") {
    e.target.value = e.target.value.replace(/\D/g, "").slice(0, 8);
  }
});

// =========================
// HORAS HH:MM 24 HORAS
// =========================
forzarFormatoHoraSinSegundos(horaIngresoInput);
forzarFormatoHoraSinSegundos(horaSalidaInput);

function forzarFormatoHoraSinSegundos(input) {
  if (!input) return;

  input.addEventListener("input", () => {
    input.value = normalizarHoraParcial(input.value);
  });

  input.addEventListener("blur", () => {
    input.value = normalizarHoraFinal(input.value);
  });
}

function normalizarHoraParcial(valor) {
  let limpio = String(valor || "").replace(/\D/g, "").slice(0, 4);

  if (limpio.length >= 3) {
    limpio = limpio.slice(0, 2) + ":" + limpio.slice(2);
  }

  return limpio;
}

function normalizarHoraFinal(valor) {
  const limpio = String(valor || "").trim();

  if (!limpio) return "";

  const match = limpio.match(/^(\d{1,2}):?(\d{1,2})$/);
  if (!match) return limpio;

  let hh = parseInt(match[1], 10);
  let mm = parseInt(match[2], 10);

  if (isNaN(hh) || isNaN(mm)) return limpio;

  if (hh < 0) hh = 0;
  if (hh > 23) hh = 23;
  if (mm < 0) mm = 0;
  if (mm > 59) mm = 59;

  return String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0");
}

function esHoraValida(valor) {
  return /^([01]\d|2[0-3]):([0-5]\d)$/.test(String(valor || "").trim());
}

function esCorreoValido(valor) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(valor || "").trim());
}

// =========================
// DESCARGAR / CARGAR PLANTILLA
// =========================
btnDescargarPlantilla.addEventListener("click", () => {
  descargarPlantillaExcel();
});

inputPlantilla.addEventListener("change", async (e) => {
  const archivo = e.target.files[0];
  if (!archivo) return;

  const extension = archivo.name.split(".").pop().toLowerCase();

  try {
    if (extension === "csv") {
      await cargarCSV(archivo);
    } else if (extension === "xlsx" || extension === "xls") {
      await cargarExcel(archivo);
    } else {
      alert("Formato no permitido. Solo se admite .xlsx, .xls o .csv");
    }
  } catch (error) {
    console.error("Error al cargar plantilla:", error);
    alert("Ocurrió un error al leer la plantilla.");
  }

  e.target.value = "";
});

function descargarPlantillaExcel() {
  const data = [
    ["Nombre Funcionario", "DNI"],
    ["JUAN PEREZ DIAZ", "12345678"],
    ["MARIA LOPEZ TORRES", "87654321"]
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);

  XLSX.utils.book_append_sheet(wb, ws, "Personal");
  XLSX.writeFile(wb, "Plantilla_Listado_Personal.xlsx");
}

async function cargarExcel(archivo) {
  const arrayBuffer = await archivo.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const primeraHoja = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[primeraHoja];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

  procesarFilasPlantilla(rows);
}

async function cargarCSV(archivo) {
  const texto = await archivo.text();
  const filas = texto
    .split(/\r?\n/)
    .map((linea) => linea.split(","));

  procesarFilasPlantilla(filas);
}

function procesarFilasPlantilla(rows) {
  if (!rows || rows.length < 2) {
    alert("La plantilla no contiene datos.");
    return;
  }

  const encabezados = rows[0].map((h) => String(h).trim().toLowerCase());

  const indexNombre = encabezados.findIndex(
    (h) => h === "nombre funcionario" || h === "nombre" || h === "funcionario"
  );

  const indexDni = encabezados.findIndex(
    (h) => h === "dni" || h === "documento"
  );

  if (indexNombre === -1 || indexDni === -1) {
    alert('La plantilla debe tener las columnas "Nombre Funcionario" y "DNI".');
    return;
  }

  let registrosAgregados = 0;

  rows.slice(1).forEach((row) => {
    const nombre = String(row[indexNombre] ?? "").trim();
    let dni = String(row[indexDni] ?? "").trim();

    dni = dni.replace(/\D/g, "").slice(0, 8);

    if (nombre || dni) {
      agregarFilaPersonal(nombre, dni);
      registrosAgregados++;
    }
  });

  if (registrosAgregados === 0) {
    alert("No se encontraron registros válidos en la plantilla.");
    return;
  }

  limpiarFilaInicialVacia();
  alert(`Se cargaron ${registrosAgregados} registros correctamente.`);
}

function limpiarFilaInicialVacia() {
  const filas = tbodyPersonal.querySelectorAll("tr");
  if (filas.length <= 1) return;

  const primeraFila = filas[0];
  const nombre = primeraFila.querySelector('input[name="nombreFuncionario[]"]').value.trim();
  const dni = primeraFila.querySelector('input[name="dni[]"]').value.trim();

  if (!nombre && !dni) {
    primeraFila.remove();
  }
}

// =========================
// BUSCADOR DE SITE
// =========================
buscarRelacion.addEventListener("input", () => {
  const termino = buscarRelacion.value.trim().toLowerCase();

  if (!termino) {
    ocultarResultados();
    return;
  }

  const resultados = sitesData.filter((item) => {
    const codigo = String(item.codigoUnico || "").toLowerCase();
    const site = String(item.site || "").toLowerCase();
    const direccion = String(item.direccion || "").toLowerCase();
    const torrera = String(item.tipoTorrera || item.torrera || "").toLowerCase();

    return (
      codigo.includes(termino) ||
      site.includes(termino) ||
      direccion.includes(termino) ||
      torrera.includes(termino)
    );
  }).slice(0, 8);

  renderizarResultados(resultados);
});

function renderizarResultados(resultados) {
  if (!resultados.length) {
    searchResults.innerHTML = `
      <div class="search-result-item">
        <div class="search-result-title">Sin resultados</div>
        <div class="search-result-subtitle">No se encontró información relacionada.</div>
      </div>
    `;
    searchResults.style.display = "block";
    return;
  }

  searchResults.innerHTML = resultados.map((item, index) => `
    <div class="search-result-item" data-index="${index}">
      <div class="search-result-title">${escapeHtml(item.codigoUnico || "")} - ${escapeHtml(item.site || "")}</div>
      <div class="search-result-subtitle">${escapeHtml(item.direccion || "")}</div>
    </div>
  `).join("");

  searchResults.style.display = "block";

  const items = searchResults.querySelectorAll(".search-result-item");
  items.forEach((element, index) => {
    element.addEventListener("click", () => {
      seleccionarSite(resultados[index]);
    });
  });
}

function seleccionarSite(item) {
  document.getElementById("codigoUnico").value = item.codigoUnico || "";
  document.getElementById("nombreLocal").value = item.site || "";
  document.getElementById("direccion").value = item.direccion || "";
  document.getElementById("ubigeo").value = item.ubigeo || "";
  document.getElementById("departamento").value = item.departamento || "";
  document.getElementById("provincia").value = item.provincia || "";
  document.getElementById("distrito").value = item.distrito || "";
  document.getElementById("tipoLocal").value = item.tipoLocal || "";

  if (document.getElementById("tipoTorrera")) {
    document.getElementById("tipoTorrera").value = item.tipoTorrera || item.torrera || "";
  }

  buscarRelacion.value = `${item.codigoUnico || ""} - ${item.site || ""}`;
  ocultarResultados();
}

function ocultarResultados() {
  searchResults.style.display = "none";
  searchResults.innerHTML = "";
}

document.addEventListener("click", (e) => {
  if (!e.target.closest(".site-search-wrapper")) {
    ocultarResultados();
  }
});

// =========================
// LIMPIAR
// =========================
btnLimpiar.addEventListener("click", () => {
  if (guardando) return;
  limpiarFormulario();
});

function limpiarFormulario() {
  formAcceso.reset();

  document.getElementById("buscarRelacion").value = "";
  ocultarResultados();

  if (horaIngresoInput) horaIngresoInput.value = "";
  if (horaSalidaInput) horaSalidaInput.value = "";

  tbodyPersonal.innerHTML = `
    <tr>
      <td>
        <input type="text" name="nombreFuncionario[]" placeholder="Ingrese nombre completo" />
      </td>
      <td>
        <input type="text" name="dni[]" placeholder="Ingrese DNI" maxlength="8" />
      </td>
      <td class="text-center">
        <button type="button" class="btn btn-danger btnEliminarFila">Eliminar</button>
      </td>
    </tr>
  `;
}

// =========================
// POPUP
// =========================
btnCerrarPopup.addEventListener("click", () => {
  popupExito.classList.remove("active");
});

popupExito.addEventListener("click", (e) => {
  if (e.target === popupExito) {
    popupExito.classList.remove("active");
  }
});

function mostrarPopupExito() {
  popupExito.classList.add("active");
}

// =========================
// LOADER
// =========================
function mostrarLoader() {
  loaderOverlay.classList.add("active");
}

function ocultarLoader() {
  loaderOverlay.classList.remove("active");
}

// =========================
// RESPUESTA API SEGURA
// =========================
async function obtenerRespuestaApiSegura(response) {
  const rawText = await response.text();
  console.log("Respuesta cruda API:", rawText);

  try {
    return JSON.parse(rawText);
  } catch (error) {
    const texto = String(rawText || "").trim();

    if (
      response.ok &&
      (
        texto.startsWith("<!DOCTYPE html") ||
        texto.startsWith("<html") ||
        texto.startsWith("<!doctype html")
      )
    ) {
      return {
        ok: true,
        message: "Registro realizado y notificado"
      };
    }

    throw new Error("La API devolvió una respuesta no válida.");
  }
}

// =========================
// GUARDAR
// =========================
formAcceso.addEventListener("submit", async (e) => {
  e.preventDefault();

  if (guardando) return;

  try {
    guardando = true;
    mostrarLoader();

    validarFormulario();

    const data = await obtenerDatosFormularioParaEnvio();

    const response = await fetch(API_URL, {
      method: "POST",
      headers: {
        "Content-Type": "text/plain;charset=utf-8"
      },
      body: JSON.stringify(data),
      redirect: "follow"
    });

    const result = await obtenerRespuestaApiSegura(response);
    console.log("Respuesta API:", result);

    if (!result.ok) {
      throw new Error(result.message || "No se pudo registrar.");
    }

    ocultarLoader();
    mostrarPopupExito();

  } catch (error) {
    console.error(error);
    ocultarLoader();
    alert("Ocurrió un error al guardar: " + error.message);
  } finally {
    guardando = false;
  }
});

function validarFormulario() {
  const codigoUnico = document.getElementById("codigoUnico").value.trim();
  const nombreLocal = document.getElementById("nombreLocal").value.trim();
  const tipoTorrera = document.getElementById("tipoTorrera").value.trim();
  const tipoAcceso = document.getElementById("tipoAcceso").value.trim();
  const fechaInicio = document.getElementById("fechaInicio").value.trim();
  const horaIngreso = normalizarHoraFinal(document.getElementById("horaIngreso").value.trim());
  const horaSalida = normalizarHoraFinal(document.getElementById("horaSalida").value.trim());

  const solicitanteNombre = document.getElementById("solicitanteNombre").value.trim();
  const solicitanteCorreo = document.getElementById("solicitanteCorreo").value.trim();

  document.getElementById("horaIngreso").value = horaIngreso;
  document.getElementById("horaSalida").value = horaSalida;

  if (!codigoUnico) throw new Error("Ingrese o seleccione el código único.");
  if (!nombreLocal) throw new Error("Ingrese o seleccione el nombre del local.");
  if (!tipoTorrera) throw new Error("Seleccione la torrera.");
  if (!tipoAcceso) throw new Error("Seleccione el tipo de acceso.");
  if (!fechaInicio) throw new Error("Seleccione la fecha de inicio.");
  if (!horaIngreso) throw new Error("Ingrese la hora de ingreso.");
  if (!esHoraValida(horaIngreso)) throw new Error("La hora de ingreso debe estar entre 00:00 y 23:59.");
  if (horaSalida && !esHoraValida(horaSalida)) throw new Error("La hora de salida debe estar entre 00:00 y 23:59.");

  if (!solicitanteNombre) throw new Error("Ingrese el nombre del solicitante.");
  if (!solicitanteCorreo) throw new Error("Ingrese el correo del solicitante.");
  if (!esCorreoValido(solicitanteCorreo)) throw new Error("Ingrese un correo válido para el solicitante.");
}

async function obtenerDatosFormularioParaEnvio() {
  const personal = [];

  const filas = tbodyPersonal.querySelectorAll("tr");
  filas.forEach((fila) => {
    const nombre = fila.querySelector('input[name="nombreFuncionario[]"]').value.trim();
    const dni = fila.querySelector('input[name="dni[]"]').value.trim();

    if (nombre || dni) {
      personal.push({
        nombreFuncionario: nombre,
        dni: dni
      });
    }
  });

  const archivoPdfInput = document.getElementById("archivoPdf");
  const archivo = archivoPdfInput.files[0] || null;

  let archivoPdf = null;

  if (archivo) {
    archivoPdf = await convertirArchivoABase64(archivo);
  }

  return {
    datosSite: {
      codigoUnico: document.getElementById("codigoUnico").value.trim(),
      nombreLocal: document.getElementById("nombreLocal").value.trim(),
      direccion: document.getElementById("direccion").value.trim(),
      ubigeo: document.getElementById("ubigeo").value.trim(),
      departamento: document.getElementById("departamento").value.trim(),
      provincia: document.getElementById("provincia").value.trim(),
      distrito: document.getElementById("distrito").value.trim(),
      tipoLocal: document.getElementById("tipoLocal").value.trim()
    },
    trabajo: {
      tipoTorrera: document.getElementById("tipoTorrera").value,
      tipoAcceso: document.getElementById("tipoAcceso").value,
      fechaInicio: document.getElementById("fechaInicio").value,
      fechaFin: document.getElementById("fechaFin").value,
      horaIngreso: document.getElementById("horaIngreso").value,
      horaSalida: document.getElementById("horaSalida").value
    },
    solicitante: {
      nombre: document.getElementById("solicitanteNombre").value.trim(),
      correo: document.getElementById("solicitanteCorreo").value.trim()
    },
    archivoPdf,
    listadoPersonal: personal
  };
}

function convertirArchivoABase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const result = reader.result || "";
      const base64 = String(result).split(",")[1];

      resolve({
        nombre: file.name,
        tipo: file.type,
        base64: base64
      });
    };

    reader.onerror = () => reject(new Error("No se pudo leer el PDF."));
    reader.readAsDataURL(file);
  });
}

function escapeHtml(valor) {
  return String(valor || "")
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
