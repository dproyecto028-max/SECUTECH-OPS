
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwoQhHRRIhnITU28Pk2ZqVw8-ap5jUE05HiBlMvKPCWnYV-dT_RS4oxDpLQj48VxZs/exec";
const CODIGOS_COLABORADORES = ["ATLAS2025", "ATLASLIMPIEZA", "ADMINATLAS"];

const PROVINCIAS_Y_LOCALIDADES = {
  Mendoza: [
    { nombre: "Capital", cp: "5500" },
    { nombre: "Godoy Cruz", cp: "5501" },
    { nombre: "Luján de Cuyo", cp: "5507" },
    { nombre: "Maipú", cp: "5515" },
    { nombre: "Las Heras", cp: "5539" },
    { nombre: "San Martín", cp: "5540" },
    { nombre: "General San Martín", cp: "5521" },
    { nombre: "Rivadavia", cp: "5600" },
    { nombre: "Guaymallén", cp: "5519" },
    { nombre: "Junín", cp: "5570" },
    { nombre: "Malargüe", cp: "5613" },
    { nombre: "General Alvear", cp: "5620" },
    { nombre: "Santa Rosa", cp: "5605" },
    { nombre: "San Rafael", cp: "5600" },
    { nombre: "Tunuyán", cp: "5603" },
    { nombre: "Mendoza", cp: "5500" },
    { nombre: "Villa Nueva", cp: "5533" },
    { nombre: "Lavalle", cp: "5598" },
    { nombre: "La Paz", cp: "5597" },
    { nombre: "Tupungato", cp: "5561" }
  ],
  "San Juan": [
    { nombre: "Capital", cp: "5400" },
    { nombre: "Caucete", cp: "5430" },
    { nombre: "Santa Lucía", cp: "5401" },
    { nombre: "Rawson", cp: "5427" }
  ],
  "San Luis": [
    { nombre: "San Luis", cp: "5700" },
    { nombre: "Villa Mercedes", cp: "5730" },
    { nombre: "La Punta", cp: "5701" },
    { nombre: "Merlo", cp: "5881" }
  ],
  "La Rioja": [
    { nombre: "La Rioja", cp: "5300" },
    { nombre: "Chilecito", cp: "5360" },
    { nombre: "Aimogasta", cp: "5311" },
    { nombre: "Anillaco", cp: "5387" }
  ]
};

let usuarioActual = {
  nombre: "",
  email: "",
  tipo: ""
};

function poblarProvincias() {
  const provinciaSelect = document.getElementById("regProvincia");
  if (!provinciaSelect) return;

  const provincias = Object.keys(PROVINCIAS_Y_LOCALIDADES).sort();
  provincias.forEach((provincia) => {
    const option = document.createElement("option");
    option.value = provincia;
    option.textContent = provincia;
    provinciaSelect.appendChild(option);
  });

  provinciaSelect.addEventListener("change", () => {
    actualizarLocalidadesYCP(provinciaSelect.value);
  });
}

function actualizarLocalidadesYCP(provincia) {
  const localidadSelect = document.getElementById("regLocalidad");
  const cpInput = document.getElementById("regCp");
  if (!localidadSelect || !cpInput) return;

  localidadSelect.innerHTML = '<option value="">Seleccioná una localidad</option>';
  cpInput.value = "";

  if (!provincia || !PROVINCIAS_Y_LOCALIDADES[provincia]) {
    return;
  }

  PROVINCIAS_Y_LOCALIDADES[provincia].forEach((item) => {
    const option = document.createElement("option");
    option.value = item.nombre;
    option.textContent = item.nombre;
    localidadSelect.appendChild(option);
  });

  localidadSelect.addEventListener("change", () => {
    const seleccionada = PROVINCIAS_Y_LOCALIDADES[provincia].find((item) => item.nombre === localidadSelect.value);
    cpInput.value = seleccionada ? seleccionada.cp : "";
  }, { once: true });
}

function initRegistroUbicacion() {
  poblarProvincias();
  actualizarLocalidadesYCP(document.getElementById("regProvincia")?.value || "");
}

function openModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    modal.style.display = "flex";
  }
}

function closeModal(id) {
  const modal = document.getElementById(id);
  if (modal) {
    modal.style.display = "none";
  }
}

function setTipoRegistro(tipo) {
  document.getElementById("regTipo").value = tipo;

  const btnCliente = document.getElementById("btnTipoCliente");
  const btnColab = document.getElementById("btnTipoColab");
  const colabFields = document.getElementById("colabFields");
  const codigoInput = document.getElementById("regCodigoAdmin");

  if (tipo === "cliente") {
    btnCliente.classList.add("active");
    btnColab.classList.remove("active");
    colabFields.classList.add("hidden");
    if (codigoInput) {
      codigoInput.classList.add("hidden");
      codigoInput.required = false;
    }
  } else {
    btnColab.classList.add("active");
    btnCliente.classList.remove("active");
    colabFields.classList.remove("hidden");
    if (codigoInput) {
      codigoInput.classList.remove("hidden");
      codigoInput.required = true;
    }
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    if (!file) {
      resolve(null);
      return;
    }

    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

function isScriptReady() {
  return typeof SCRIPT_URL === "string" && SCRIPT_URL.trim() !== "";
}

async function submitRegistro(e) {
  e.preventDefault();

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App de Apps Script en app.js antes de registrar.");
    return;
  }

  const tipo = document.getElementById("regTipo").value;
  const btnSubmit = document.getElementById("btnRegistrar");

  if (tipo === "colaborador") {
    const codigoIngresado = (document.getElementById("regCodigoAdmin").value || "").trim().toUpperCase();
    if (!CODIGOS_COLABORADORES.includes(codigoIngresado)) {
      alert("Código de acceso inválido. Solo puede registrarse el personal autorizado por administración.");
      return;
    }
  }

  btnSubmit.disabled = true;
  btnSubmit.textContent = "Procesando...";

  let cvBase64 = null;
  let cvNombre = "";
  let fotoBase64 = null;
  let fotoNombre = "";

  if (tipo === "colaborador") {
    const cvInput = document.getElementById("regCvFile");
    const fotoInput = document.getElementById("regFotoFile");

    if (cvInput.files.length > 0) {
      cvBase64 = await fileToBase64(cvInput.files[0]);
      cvNombre = cvInput.files[0].name;
    }

    if (fotoInput.files.length > 0) {
      fotoBase64 = await fileToBase64(fotoInput.files[0]);
      fotoNombre = fotoInput.files[0].name;
    }
  }

  const payload = {
    action: tipo === "cliente" ? "registro_cliente" : "registro_colaborador",
    nombre: document.getElementById("regNombre").value,
    dni: document.getElementById("regDni").value,
    password: document.getElementById("regPass").value,
    email: document.getElementById("regEmail").value,
    provincia: document.getElementById("regProvincia").value,
    localidad: document.getElementById("regLocalidad").value,
    calle: document.getElementById("regCalle").value,
    altura: document.getElementById("regAltura").value,
    cp: document.getElementById("regCp").value,
    celular: document.getElementById("regCelular").value,
    codigoAdmin: document.getElementById("regCodigoAdmin").value,
    cvBase64,
    cvNombre,
    fotoBase64,
    fotoNombre
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    alert(data.message || "Registro enviado correctamente.");

    if (data.status === "success") {
      document.getElementById("formRegistro").reset();
      closeModal("registerModal");
    }
  } catch (error) {
    alert("No se pudo completar el registro. Revisá la URL del Apps Script y la conexión.");
  } finally {
    btnSubmit.disabled = false;
    btnSubmit.textContent = "Registrarse";
  }
}

async function submitLogin(e) {
  e.preventDefault();

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App de Apps Script en app.js antes de iniciar sesión.");
    return;
  }

  const payload = {
    action: "login",
    dni: document.getElementById("loginDni").value,
    password: document.getElementById("loginPass").value
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();

    if (data.status === "success") {
      usuarioActual = {
        nombre: data.name,
        email: data.email,
        tipo: data.type
      };

      alert(`Bienvenido/a ${data.name}`);
      closeModal("loginModal");

      if (data.type === "cliente") {
        document.getElementById("dashboardCliente").classList.remove("hidden");
        document.getElementById("dashboardCliente").scrollIntoView({ behavior: "smooth", block: "start" });
      } else {
        document.getElementById("dashboardColaborador").classList.remove("hidden");
        document.getElementById("dashboardColaborador").scrollIntoView({ behavior: "smooth", block: "start" });
      }
    } else {
      alert(data.message || "Credenciales incorrectas.");
    }
  } catch (error) {
    alert("No se pudo iniciar sesión. Verificá la conexión y la publicación del Apps Script.");
  }
}

async function contratarServicio(plan) {
  if (!usuarioActual.email) {
    alert("Debés iniciar sesión antes de contratar un servicio.");
    openModal("loginModal");
    return;
  }

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App de Apps Script en app.js para confirmar el pago.");
    return;
  }

  const payload = {
    action: "confirmar_pago",
    nombre: usuarioActual.nombre,
    email: usuarioActual.email,
    plan
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    alert(data.message || `Pago confirmado para ${plan}.`);
  } catch (error) {
    alert("No se pudo confirmar el pago, pero el flujo visual quedó listo.");
  }
}

async function registrarJornada(tipo) {
  if (!usuarioActual.email) {
    alert("Necesitás iniciar sesión para registrar la jornada.");
    openModal("loginModal");
    return;
  }

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App para registrar la jornada.");
    return;
  }

  const status = document.getElementById("statusJornada");
  const hora = new Date().toLocaleTimeString();

  status.textContent = `Registrando ${tipo.toLowerCase()}...`;

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        action: "marcar_jornada",
        tipo,
        hora,
        email: usuarioActual.email
      })
    });

    const data = await res.json();
    status.textContent = data.message || `Jornada ${tipo.toLowerCase()} registrada a las ${hora}.`;
  } catch (error) {
    status.textContent = `No se pudo registrar el ${tipo.toLowerCase()} de jornada.`;
  }
}

function submitIncidente(e) {
  e.preventDefault();

  const tipo = document.getElementById("incidenteTipo").value;
  const desc = document.getElementById("incidenteDesc").value;

  if (!tipo || !desc.trim()) {
    alert("Completá el tipo de incidente y la descripción.");
    return;
  }

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App para guardar el reporte.");
    return;
  }

  fetch(SCRIPT_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      action: "guardar_reporte",
      tipo,
      descripcion: desc,
      email: usuarioActual.email || "sin_usuario"
    })
  })
    .then((res) => res.json())
    .then((data) => {
      alert(data.message || "Reporte registrado correctamente.");
      document.getElementById("formIncidente").reset();
    })
    .catch(() => {
      alert("No se pudo enviar el reporte en este momento.");
    });
}

async function submitTrabajoNosotros(e) {
  e.preventDefault();

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App para registrar la postulación.");
    return;
  }

  const cvFile = document.getElementById("postulanteCvFile").files[0];
  const cvBase64 = cvFile ? await fileToBase64(cvFile) : null;

  const payload = {
    action: "trabaja_con_nosotros",
    nombre: document.getElementById("postulanteNombre").value,
    email: document.getElementById("postulanteEmail").value,
    telefono: document.getElementById("postulanteTelefono").value,
    puesto: document.getElementById("postulantePuesto").value,
    mensaje: document.getElementById("postulanteMensaje").value,
    cvBase64,
    cvNombre: cvFile ? cvFile.name : ""
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    alert(data.message || "Postulación enviada con éxito.");

    if (data.status === "success") {
      document.getElementById("formTrabajaConNosotros").reset();
    }
  } catch (error) {
    alert("No se pudo enviar la postulación. Verificá la conexión.");
  }
}

async function submitContratoServicio(e) {
  e.preventDefault();

  if (!usuarioActual.email) {
    alert("Debés iniciar sesión para contratar el servicio.");
    openModal("loginModal");
    return;
  }

  if (!isScriptReady()) {
    alert("Configurá la URL de tu Web App para registrar el contrato.");
    return;
  }

  const comprobanteFile = document.getElementById("contratoComprobante").files[0];
  const comprobanteBase64 = comprobanteFile ? await fileToBase64(comprobanteFile) : null;

  const payload = {
    action: "contrato_servicio",
    nombre: usuarioActual.nombre,
    email: usuarioActual.email,
    horas: document.getElementById("contratoHoras").value,
    tipoTrabajo: document.getElementById("contratoTipoTrabajo").value,
    descripcion: document.getElementById("contratoDescripcion").value,
    comprobanteBase64,
    comprobanteNombre: comprobanteFile ? comprobanteFile.name : ""
  };

  try {
    const res = await fetch(SCRIPT_URL, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams(payload)
    });

    const data = await res.json();
    alert(data.message || "Contrato registrado correctamente.");

    if (data.status === "success") {
      document.getElementById("formContratoServicio").reset();
    }
  } catch (error) {
    alert("No se pudo registrar el contrato. Revisá la conexión.");
  }
}

document.getElementById("formRegistro").addEventListener("submit", submitRegistro);
document.getElementById("formLogin").addEventListener("submit", submitLogin);
document.getElementById("formIncidente").addEventListener("submit", submitIncidente);
const formTrabaja = document.getElementById("formTrabajaConNosotros");
if (formTrabaja) formTrabaja.addEventListener("submit", submitTrabajoNosotros);

const formContrato = document.getElementById("formContratoServicio");
if (formContrato) formContrato.addEventListener("submit", submitContratoServicio);

window.addEventListener("click", (event) => {
  const modals = document.querySelectorAll(".modal");
  modals.forEach((modal) => {
    if (event.target === modal) {
      modal.style.display = "none";
    }
  });
});

setTipoRegistro("cliente");
initRegistroUbicacion();
