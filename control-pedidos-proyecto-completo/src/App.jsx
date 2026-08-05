import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";
import { Plus, Trash2, Download, RefreshCw, ClipboardList, Truck, ChevronLeft, Check, AlertCircle, PackageSearch, Lock, LogOut, Mail, FileUp, X, Clock } from "lucide-react";

const SUPABASE_URL = "https://biujteotjtafzsmkbbqi.supabase.co";
const SUPABASE_KEY = "sb_publishable_TrdxYYtD9Kt4bD6VLCip_Q_J2mlpi6E";

const VENDEDORES = [
  { nombre: "Betzabeth", usuario: "Betzabeth", claveHash: "c3a7295163aba5a14d6f1bd05c78d584aedc26725b93dfe632489570071a05ea" },
  { nombre: "Ana", usuario: "ana", claveHash: "766e0fca25603c291748b46bf7ad33f895dbd557e4ed7823cd9faae78db652ae" },
  { nombre: "Linda", usuario: "Linda", claveHash: "1ec509fcb7b38763149d30e52f32aed9b69086ecbc466940d11468bc8cc0a582" },
  { nombre: "Joshua", usuario: "Joshua", claveHash: "24cfa93cff8b9bdf1806824dcf7f921f781cf1a630b63bf7081dba2a7076941f" },
  { nombre: "Marianela", usuario: "Marianela", claveHash: "aef7d988b521d7ab165d4ebcf5cc797299908398afb662d8fb4d499a47e283c4" },
  { nombre: "Itvan", usuario: "Itvan", claveHash: "5d6a8c626f97a82518e9eadd490afed125216374803c3dc1ec108a44d8217803" },
  { nombre: "Elsa", usuario: "Elsa", claveHash: "e789234cb6cdddcd54157fa6fc4777fba93de3d752cccab7eb9c036af1f1bd66" },
  { nombre: "Xochil", usuario: "xochil", claveHash: "53b3231e59d9620e76f00686ba92606fae271387d7cb5d9c4dcf1f6dc3f75a99" },
  { nombre: "Cathy", usuario: "cathy", claveHash: "a2f49e028dc9638a0ffada8e4b31b9e0c1179651298a5868c0202402287a7835" },
  { nombre: "Elizabeth", usuario: "Elizabeth", claveHash: "3fd1f049f6eb7e94a7dca10825b92699d726662df47fb4b571acc6b4e9c9b758" },
  { nombre: "Carlos", usuario: "carlos", claveHash: "09eae93143621a9e2eec842e2bfe83036ceeb8cc0ccc788023af99abe4e80634" },
  { nombre: "Boris", usuario: "boris", claveHash: "453c33c4281fcbe0984f265cc42387cb54ea459192dc413a362175e6aefd7885" },
  { nombre: "Gabriel", usuario: "Gabriel", claveHash: "10e7d12f50fd78dfc4e99d233f934914dd011015396c4ab0165c5b5655f036d9" },
  { nombre: "Omar", usuario: "omar", claveHash: "d8486754f3fe5c91a871f7f63cc20cd5a192126373703a109176801acd417227" },
  { nombre: "Maykeling", usuario: "Maykeling", claveHash: "717ea22b306250e77e77caf8b5664237df875ec3748f2f657ffa0c8d739d5984" },
  { nombre: "Katiana", usuario: "Katiana", claveHash: "4816c05b087aae470d5f03a40bc20ba9f0b27e14d3aa45402c1275424ca1512f" },
];

const COMPRAS = { nombre: "Zoraida", usuario: "Zoraida", claveHash: "5e3e50b1d01723ed0404c32b139229998821fcaae009c0a857495bcf6bc21fcd" };

// Genera un hash SHA-256 en hexadecimal de un texto, usando la Web Crypto
// del navegador (no depende de librerías externas).
async function sha256Hex(texto) {
  const datos = new TextEncoder().encode(texto);
  const buffer = await crypto.subtle.digest("SHA-256", datos);
  return Array.from(new Uint8Array(buffer))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");
}

// Notificaciones push gratuitas vía ntfy.sh (sin cuentas ni llaves).
// Zoraida se suscribe a este mismo "tema" desde la app ntfy o el navegador.
const NTFY_TEMA = "colon-pedidos-e059fa159caa";
const CORREO_COMPRAS = "zoraida@ilumitec.com";

function textoPedido(pedido) {
  const lineasTexto = pedido.lineas.map((l) => `${l.codigo}\t${l.cantidad}`).join("\n");
  return (
    `Fecha: ${pedido.fecha}\tPedido: ${pedido.pedidoRef}\tCliente: ${pedido.cliente}\n` +
    `Vendedor: ${pedido.vendedor}\tMétodo de pago: ${pedido.metodoPago === "credito" ? "Crédito" : "Contado"}\n\n` +
    `Códigos:\tCantidad\n${lineasTexto}\n`
  );
}

const CELDA_ESTILO = "border:1px solid #000000; padding:4px 10px; font-family:Arial,sans-serif; font-size:13px;";

function htmlPedido(pedido) {
  const celda = (contenido) => `<td style="${CELDA_ESTILO}">${contenido}</td>`;
  const filasCodigos = pedido.lineas
    .map((l) => `<tr>${celda(l.codigo)}${celda(l.cantidad)}${celda("")}</tr>`)
    .join("");
  return (
    `<p style="font-family:Arial,sans-serif;font-size:13px;">Para: ${CORREO_COMPRAS}</p>` +
    `<table style="border-collapse:collapse;">` +
    `<tr>${celda(`<b>Fecha:</b> ${pedido.fecha}`)}${celda(`<b>Pedido:</b> ${pedido.pedidoRef}`)}${celda(`<b>Cliente:</b> ${pedido.cliente}`)}</tr>` +
    `<tr>${celda(`<b>Vendedor:</b> ${pedido.vendedor}`)}${celda(`<b>Método de pago:</b> ${pedido.metodoPago === "credito" ? "Crédito" : "Contado"}`)}${celda("")}</tr>` +
    `<tr>${celda("")}${celda("")}${celda("")}</tr>` +
    `<tr>${celda("<b>Códigos:</b>")}${celda("<b>Cantidad</b>")}${celda("")}</tr>` +
    filasCodigos +
    `</table>`
  );
}

function construirMailtoPedido(pedido) {
  const asunto = encodeURIComponent(`Pedido ${pedido.pedidoRef} · ${pedido.cliente} · ${pedido.vendedor}`);
  const cuerpo = encodeURIComponent(textoPedido(pedido));
  return `mailto:${CORREO_COMPRAS}?subject=${asunto}&body=${cuerpo}`;
}

async function avisarNuevoPedido(pedido) {
  try {
    await fetch(`https://ntfy.sh/${NTFY_TEMA}`, {
      method: "POST",
      body: `Pedido ${pedido.pedidoRef} · ${pedido.cliente} · ${pedido.vendedor} · ${pedido.lineas.length} código${pedido.lineas.length !== 1 ? "s" : ""}`,
      headers: { Title: "Nuevo pedido de venta", Priority: "default", Tags: "package" },
    });
  } catch {
    // si falla el aviso push no debe afectar el guardado del pedido
  }
}

const uid = () => (crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(36).slice(2));

// --- Importador de cotizaciones en PDF (formato de cotización de Ilumitec) ---
// Carga pdf.js una sola vez desde cdnjs (los artifacts en React no permiten
// "import" librerías fuera de la lista fija, así que se inyecta como <script>).
let pdfjsCargando = null;
function cargarPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (pdfjsCargando) return pdfjsCargando;
  pdfjsCargando = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js";
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    script.onerror = () => reject(new Error("No se pudo cargar el lector de PDF. Revisa tu conexión."));
    document.head.appendChild(script);
  });
  return pdfjsCargando;
}

// Junta palabras cercanas en el eje Y como una sola "línea" visual de la tabla.
function agruparLineas(palabras, tolerancia = 3) {
  const ordenadas = [...palabras].sort((a, b) => b.y - a.y);
  const lineas = [];
  let actual = [];
  let referencia = null;
  for (const p of ordenadas) {
    if (referencia === null || Math.abs(p.y - referencia) <= tolerancia) {
      actual.push(p);
      if (referencia === null) referencia = p.y;
    } else {
      lineas.push(actual);
      actual = [p];
      referencia = p.y;
    }
  }
  if (actual.length) lineas.push(actual);
  return lineas.map((l) => l.sort((a, b) => a.x - b.x));
}

// Dadas las posiciones X de los encabezados de columna (No., Código, Cant., etc.),
// arma una función que dice a qué columna pertenece cualquier palabra según su X.
function limitesColumnas(anchors) {
  const entradas = Object.entries(anchors).filter(([, x]) => x !== undefined);
  entradas.sort((a, b) => a[1] - b[1]);
  const bounds = entradas.map(([nombre, x], i) => ({
    nombre,
    izq: i === 0 ? x - 30 : (entradas[i - 1][1] + x) / 2,
    der: i === entradas.length - 1 ? x + 30 : (x + entradas[i + 1][1]) / 2,
  }));
  return (x) => (bounds.find((b) => x >= b.izq && x < b.der) || {}).nombre || null;
}

async function extraerCotizacionPdf(archivo) {
  const pdfjsLib = await cargarPdfJs();
  const buffer = await archivo.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buffer }).promise;

  let cliente = null,
    vendedorCotizacion = null,
    referenciaTipo = null,
    referenciaNum = null,
    metodo = null;
  const items = [];

  for (let n = 1; n <= pdf.numPages; n++) {
    const page = await pdf.getPage(n);
    const contenido = await page.getTextContent();
    const palabras = contenido.items
      .map((it) => ({ str: (it.str || "").trim(), x: it.transform[4], y: it.transform[5] }))
      .filter((p) => p.str !== "");

    const lineas = agruparLineas(palabras);

    let anchors = null;
    for (const linea of lineas) {
      const mapa = {};
      linea.forEach((p) => (mapa[p.str] = p.x));
      if (mapa["Código"] !== undefined && mapa["Cant."] !== undefined) {
        anchors = {
          no: mapa["No."],
          codigo: mapa["Código"],
          descripcion: mapa["Descripción"],
          emp: mapa["Emp."],
          cxb: mapa["CxB"],
          numero: mapa["Numer"],
          cant: mapa["Cant."],
          precio: mapa["Precio"],
          total: mapa["Total"],
        };
        break;
      }
    }
    if (!anchors) continue;
    const clasificar = limitesColumnas(anchors);

    for (const linea of lineas) {
      const izquierda = linea.filter((p) => p.x < 400);
      const derecha = linea.filter((p) => p.x >= 400);
      const textoIzq = izquierda.map((p) => p.str).join(" ");
      const textoDer = derecha.map((p) => p.str).join(" ");

      if (!cliente && textoIzq.startsWith("Cliente")) {
        cliente = textoIzq.replace(/^Cliente\s*/, "").trim();
      }
      if (!vendedorCotizacion && textoDer.startsWith("Vendedor")) {
        vendedorCotizacion = textoDer.replace(/^Vendedor\s*/, "").trim();
      }
      if (!referenciaNum) {
        const combinado = `${textoIzq} ${textoDer}`;
        const mCot = combinado.match(/Cotizaci[oó]n:?\s*(\d+)/i);
        const mPed = combinado.match(/Pedido:?\s*(\d+)/i);
        if (mCot) {
          referenciaTipo = "Cotización";
          referenciaNum = mCot[1];
        } else if (mPed) {
          referenciaTipo = "Pedido";
          referenciaNum = mPed[1];
        }
      }
      if (!metodo) {
        const combinado = `${textoIzq} ${textoDer}`;
        if (/\bContado\b/.test(combinado)) metodo = "contado";
        else if (/\bCr[ée]dito\b/.test(combinado)) metodo = "credito";
      }
    }

    // Detección de ítems: se busca la línea con el número de fila (No.), y
    // si el código es tan largo que el PDF lo parte en dos líneas (ej.
    // "HW-BGU10-05PRO-2PK-7W-" seguido de "DIM4K" justo debajo), se junta
    // con la línea siguiente — pero solo esa, para no arrastrar texto de
    // íconos, características o el pie de página del último ítem.
    const PALABRAS_PLANTILLA = /^(Marca|Caracter[ií]sticas|Observaci[oó]n|Cliente|Elaborado|SubTotal|Sub Total|Total|DETALLE|Sucursal|Impuesto|COTIZACION|Cotizaci[oó]n)/i;

    lineas.forEach((linea, idx) => {
      if (!(linea.length && /^\d+$/.test(linea[0].str) && clasificar(linea[0].x) === "no")) return;

      const codigoTok = linea.find((p) => clasificar(p.x) === "codigo" && p.str !== "#");
      const cantTok = linea.find((p) => clasificar(p.x) === "cant" && /^[\d,]+(\.\d+)?$/.test(p.str));
      if (!codigoTok || !cantTok) return;

      let codigoCompleto = codigoTok.str;
      const siguiente = lineas[idx + 1];
      const siguienteEsOtraFila = siguiente && /^\d+$/.test(siguiente[0]?.str || "") && clasificar(siguiente[0].x) === "no";
      // Solo se considera que el código sigue en la línea de abajo si el
      // texto capturado termina en guion (señal clara de que la palabra se
      // cortó a la mitad, ej. "HW-BGU10-05PRO-2PK-7W-"). Un código completo
      // como "CFLS-1560-40WD-05WH-E1" nunca termina en guion, así que no se
      // le pega por error texto suelto (íconos, notas) que caiga en esa
      // misma columna en la línea siguiente.
      if (siguiente && !siguienteEsOtraFila && codigoCompleto.endsWith("-")) {
        const continuacion = siguiente.find((p) => clasificar(p.x) === "codigo" && p.str !== "#");
        if (continuacion && !PALABRAS_PLANTILLA.test(continuacion.str)) {
          codigoCompleto += continuacion.str;
        }
      }
      items.push({ codigo: codigoCompleto.toUpperCase(), cantidad: cantTok.str.replace(/,/g, "") });
    });
  }

  if (items.length === 0) {
    throw new Error("No se encontraron códigos en este PDF. ¿Es una cotización de Ilumitec en el formato usual?");
  }
  return { cliente, vendedorCotizacion, referenciaTipo, referenciaNum, metodo, items };
}
const hoy = () => new Date().toISOString().slice(0, 10);
const hace = (dias) => {
  const d = new Date();
  d.setDate(d.getDate() - dias);
  return d.toISOString().slice(0, 10);
};
const horaDe = (timestampMs) => {
  if (!timestampMs) return "";
  try {
    return new Date(timestampMs).toLocaleTimeString("es-PA", { hour: "2-digit", minute: "2-digit" });
  } catch {
    return "";
  }
};

const vacio = () => ({
  id: uid(),
  fecha: hoy(),
  pedidoRef: "",
  cliente: "",
  vendedor: "",
  metodoPago: "contado",
  lineas: [{ id: uid(), codigo: "", cantidad: "" }],
  estado: "pendiente",
  creadoEn: Date.now(),
  exportadoEn: null,
});

// --- Conexión con Supabase (reemplaza el almacenamiento de Claude) ---
// Cada pedido es una fila real en la tabla "pedidos", así que ya no hace
// falta leer-modificar-escribir un bloque completo: cada guardado es
// independiente y Postgres no permite que uno tape al otro.
const SUPABASE_HEADERS = {
  apikey: SUPABASE_KEY,
  Authorization: `Bearer ${SUPABASE_KEY}`,
  "Content-Type": "application/json",
};

function filaAPedido(fila) {
  return {
    id: fila.id,
    fecha: fila.fecha,
    pedidoRef: fila.pedido_ref,
    cliente: fila.cliente,
    vendedor: fila.vendedor,
    metodoPago: fila.metodo_pago,
    lineas: fila.lineas,
    estado: fila.estado,
    creadoEn: Number(fila.creado_en),
    exportadoEn: fila.exportado_en !== null ? Number(fila.exportado_en) : null,
  };
}

function pedidoAFila(pedido) {
  return {
    id: pedido.id,
    fecha: pedido.fecha,
    pedido_ref: pedido.pedidoRef,
    cliente: pedido.cliente,
    vendedor: pedido.vendedor,
    metodo_pago: pedido.metodoPago,
    lineas: pedido.lineas,
    estado: pedido.estado,
    creado_en: pedido.creadoEn,
    exportado_en: pedido.exportadoEn,
  };
}

async function leerPedidos() {
  try {
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/pedidos?select=*&order=creado_en.desc`, {
      headers: SUPABASE_HEADERS,
    });
    if (!resp.ok) throw new Error(`Supabase respondió ${resp.status}`);
    const filas = await resp.json();
    return filas.map(filaAPedido);
  } catch (err) {
    console.error("Error leyendo pedidos:", err);
    return [];
  }
}

// Inserta un pedido nuevo como fila propia. No hay "choque" posible entre
// dos vendedores guardando a la vez: cada insert es independiente.
async function insertarPedido(pedido) {
  try {
    const resp = await fetch(`${SUPABASE_URL}/rest/v1/pedidos`, {
      method: "POST",
      headers: { ...SUPABASE_HEADERS, Prefer: "return=minimal" },
      body: JSON.stringify(pedidoAFila(pedido)),
    });
    return resp.ok;
  } catch (err) {
    console.error("Error guardando pedido:", err);
    return false;
  }
}

// Marca una lista de pedidos como exportados, actualizando cada fila por su id.
async function marcarPedidosExportados(pedidos) {
  try {
    const resultados = await Promise.all(
      pedidos.map((p) =>
        fetch(`${SUPABASE_URL}/rest/v1/pedidos?id=eq.${encodeURIComponent(p.id)}`, {
          method: "PATCH",
          headers: { ...SUPABASE_HEADERS, Prefer: "return=minimal" },
          body: JSON.stringify({
            estado: "exportado",
            exportado_en: Date.now(),
            lineas: p.lineas.map((l) => ({ ...l, precio: PRECIO_FIJO })),
          }),
        })
      )
    );
    return resultados.every((r) => r.ok);
  } catch (err) {
    console.error("Error marcando pedidos como exportados:", err);
    return false;
  }
}

export default function App() {
  const [sesion, setSesion] = useState(null); // { rol: 'vendedor', nombre } | { rol: 'compras' }
  const [modoLogin, setModoLogin] = useState(null); // 'vendedor' | 'compras' | null

  const cerrarSesion = () => {
    setSesion(null);
    setModoLogin(null);
  };

  return (
    <div style={{ fontFamily: "'Georgia', 'Times New Roman', serif", background: "#F0EDE4", minHeight: "100%", color: "#20241F" }}>
      <style>{`
        .mono { font-family: 'SFMono-Regular','Menlo','Consolas',monospace; }
        .sans { font-family: -apple-system, 'Segoe UI', Roboto, sans-serif; }
        input, select { font-family: inherit; }
        input:focus, select:focus { outline: 2px solid #8A6E4B; outline-offset: 1px; }
        button:focus-visible { outline: 2px solid #8A6E4B; outline-offset: 2px; }
        ::placeholder { color: #A6A090; }
        .stamp { transform: rotate(-3deg); }
      `}</style>
      {!sesion && !modoLogin && <Inicio onElegir={setModoLogin} />}
      {!sesion && modoLogin && (
        <PantallaLogin
          rol={modoLogin}
          onVolver={() => setModoLogin(null)}
          onExito={(datos) => {
            setSesion(datos);
            setModoLogin(null);
          }}
        />
      )}
      {sesion?.rol === "vendedor" && <VistaVendedor vendedor={sesion.nombre} onSalir={cerrarSesion} />}
      {sesion?.rol === "compras" && <VistaCompras onSalir={cerrarSesion} />}
    </div>
  );
}

function Inicio({ onElegir }) {
  return (
    <div style={{ minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center", padding: "48px 20px" }}>
      <div style={{ maxWidth: 640, width: "100%" }}>
        <div className="sans" style={{ fontSize: 11, letterSpacing: 3, textTransform: "uppercase", color: "#8A6E4B", marginBottom: 10, textAlign: "center" }}>
          Control de pedidos · Zona Libre de Colón
        </div>
        <h1 style={{ fontSize: 40, textAlign: "center", margin: "0 0 8px", color: "#20241F", lineHeight: 1.1 }}>
          Manifiesto de pedidos
        </h1>
        <p className="sans" style={{ textAlign: "center", color: "#5C5748", fontSize: 15, marginBottom: 40 }}>
          Un solo punto de entrada para los pedidos de venta.
        </p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 20 }}>
          <TarjetaRol
            icono={<Truck size={28} strokeWidth={1.5} />}
            titulo="Vendedor"
            desc="Llenar formulario de pedidos a Zona Libre."
            onClick={() => onElegir("vendedor")}
          />
          <TarjetaRol
            icono={<ClipboardList size={28} strokeWidth={1.5} />}
            titulo="Compras"
            desc="Ver pedidos realizados por Ventas PTY."
            onClick={() => onElegir("compras")}
          />
        </div>
      </div>
    </div>
  );
}

function TarjetaRol({ icono, titulo, desc, onClick }) {
  const [hover, setHover] = useState(false);
  return (
    <button
      onClick={onClick}
      onMouseEnter={() => setHover(true)}
      onMouseLeave={() => setHover(false)}
      style={{
        textAlign: "left",
        border: "1px solid #C9C2AE",
        borderRadius: 4,
        padding: 26,
        background: hover ? "#20241F" : "#FBFAF6",
        color: hover ? "#F0EDE4" : "#20241F",
        cursor: "pointer",
        transition: "background 0.15s, color 0.15s",
      }}
    >
      <div style={{ marginBottom: 14, color: hover ? "#D9C58A" : "#8A6E4B" }}>{icono}</div>
      <div style={{ fontSize: 19, marginBottom: 6 }}>{titulo}</div>
      <div className="sans" style={{ fontSize: 13, lineHeight: 1.4, color: hover ? "#C9C2AE" : "#5C5748" }}>{desc}</div>
    </button>
  );
}

function PantallaLogin({ rol, onVolver, onExito }) {
  const [usuario, setUsuario] = useState("");
  const [clave, setClave] = useState("");
  const [error, setError] = useState("");
  const [verificando, setVerificando] = useState(false);

  const intentarEntrar = async () => {
    if (verificando) return;
    setVerificando(true);
    const u = usuario.trim().toLowerCase();
    const hashIngresado = await sha256Hex(clave);
    if (rol === "compras") {
      if (u === COMPRAS.usuario.toLowerCase() && hashIngresado === COMPRAS.claveHash) {
        onExito({ rol: "compras" });
        return;
      }
    } else {
      const encontrado = VENDEDORES.find((v) => v.usuario.toLowerCase() === u && v.claveHash === hashIngresado);
      if (encontrado) {
        onExito({ rol: "vendedor", nombre: encontrado.nombre });
        return;
      }
    }
    setVerificando(false);
    setError("Usuario o clave incorrectos.");
  };

  const validar = (e) => {
    e.preventDefault();
    intentarEntrar();
  };

  const enterDirecto = (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    intentarEntrar();
  };

  return (
    <div style={{ minHeight: "100%", display: "flex", alignItems: "center", justifyContent: "center", padding: "48px 20px" }}>
      <form onSubmit={validar} style={{ maxWidth: 360, width: "100%", background: "#FBFAF6", border: "1px solid #C9C2AE", borderRadius: 4, padding: 30 }}>
        <Volver onVolver={onVolver} />
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4, color: "#8A6E4B" }}>
          <Lock size={18} />
          <h1 style={{ fontSize: 22, margin: 0 }}>{rol === "compras" ? "Acceso de compras" : "Acceso de vendedor"}</h1>
        </div>
        <p className="sans" style={{ fontSize: 13, color: "#5C5748", margin: "0 0 20px" }}>
          Ingresa tu usuario y clave para continuar.
        </p>
        <Campo label="Usuario">
          <input value={usuario} onChange={(e) => setUsuario(e.target.value)} onKeyDown={enterDirecto} style={estiloInput} autoFocus />
        </Campo>
        <div style={{ height: 12 }} />
        <Campo label="Clave">
          <input type="password" value={clave} onChange={(e) => setClave(e.target.value)} onKeyDown={enterDirecto} style={estiloInput} />
        </Campo>
        {error && (
          <div className="sans" style={{ marginTop: 14, fontSize: 13, padding: "9px 11px", borderRadius: 4, background: "#F3DCD3", color: "#8A3B22", display: "flex", gap: 8, alignItems: "flex-start" }}>
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            {error}
          </div>
        )}
        <button type="submit" disabled={verificando} className="sans" style={{ marginTop: 18, width: "100%", background: "#20241F", color: "#F0EDE4", border: "none", borderRadius: 4, padding: "12px 0", fontSize: 14, letterSpacing: 0.5, cursor: verificando ? "default" : "pointer", opacity: verificando ? 0.6 : 1 }}>
          {verificando ? "Verificando…" : "Entrar"}
        </button>
      </form>
    </div>
  );
}

function Volver({ onVolver }) {
  return (
    <button onClick={onVolver} className="sans" style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", fontSize: 13, padding: 0, marginBottom: 18 }}>
      <ChevronLeft size={16} /> Volver
    </button>
  );
}

// ---------------- VENDEDOR ----------------

function VistaVendedor({ vendedor, onSalir }) {
  const [form, setForm] = useState({ ...vacio(), vendedor });
  const [misPedidos, setMisPedidos] = useState([]);
  const [enviando, setEnviando] = useState(false);
  const [mensaje, setMensaje] = useState(null);
  const [cargandoLista, setCargandoLista] = useState(true);
  const [focoPendiente, setFocoPendiente] = useState(null);
  const [importando, setImportando] = useState(false);
  const [errorImport, setErrorImport] = useState(null);
  const [previewImport, setPreviewImport] = useState(null);
  const [filtroDesde, setFiltroDesde] = useState(hace(7));
  const [filtroHasta, setFiltroHasta] = useState("");
  const codigoRefs = useRef({});
  const cantidadRefs = useRef({});
  const archivoInputRef = useRef(null);

  useEffect(() => {
    if (focoPendiente && codigoRefs.current[focoPendiente]) {
      codigoRefs.current[focoPendiente].focus();
      setFocoPendiente(null);
    }
  }, [form.lineas, focoPendiente]);

  useEffect(() => {
    refrescarMisPedidos();
    // eslint-disable-next-line
  }, []);

  const refrescarMisPedidos = useCallback(async () => {
    setCargandoLista(true);
    const todos = await leerPedidos();
    setMisPedidos(todos);
    setCargandoLista(false);
  }, []);

  const propios = misPedidos
    .filter((p) => p.vendedor.trim().toLowerCase() === vendedor.trim().toLowerCase())
    .filter((p) => (filtroDesde ? p.fecha >= filtroDesde : true))
    .filter((p) => (filtroHasta ? p.fecha <= filtroHasta : true))
    .sort((a, b) => b.creadoEn - a.creadoEn);

  const actualizarLinea = (id, campo, valor) => {
    setForm((f) => ({ ...f, lineas: f.lineas.map((l) => (l.id === id ? { ...l, [campo]: valor } : l)) }));
  };
  const agregarLinea = (enfocar = false) => {
    const nuevoId = uid();
    setForm((f) => ({ ...f, lineas: [...f.lineas, { id: nuevoId, codigo: "", cantidad: "" }] }));
    if (enfocar) setFocoPendiente(nuevoId);
  };
  const quitarLinea = (id) => setForm((f) => ({ ...f, lineas: f.lineas.length > 1 ? f.lineas.filter((l) => l.id !== id) : f.lineas }));

  const enterEnCodigo = (linea) => (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    cantidadRefs.current[linea.id]?.focus();
  };

  const enterEnCantidad = (idx, linea) => (e) => {
    if (e.key !== "Enter") return;
    e.preventDefault();
    if (!linea.codigo.trim() || !String(linea.cantidad).trim()) return;
    if (idx === form.lineas.length - 1) {
      agregarLinea(true);
    } else {
      const siguiente = form.lineas[idx + 1];
      codigoRefs.current[siguiente.id]?.focus();
    }
  };

  const manejarArchivoPdf = async (e) => {
    const archivo = e.target.files?.[0];
    e.target.value = "";
    if (!archivo) return;
    setErrorImport(null);
    setImportando(true);
    try {
      const resultado = await extraerCotizacionPdf(archivo);
      setPreviewImport(resultado);
    } catch (err) {
      setErrorImport(err.message || "No se pudo leer el PDF.");
    }
    setImportando(false);
  };

  const quitarItemPreview = (indice) => {
    setPreviewImport((prev) => {
      if (!prev) return prev;
      const items = prev.items.filter((_, i) => i !== indice);
      return { ...prev, items };
    });
  };

  const confirmarImportacion = () => {
    if (!previewImport) return;
    setForm((f) => {
      const lineasActuales = f.lineas.filter((l) => l.codigo.trim() || String(l.cantidad).trim());
      const lineasNuevas = previewImport.items.map((it) => ({ id: uid(), codigo: it.codigo, cantidad: it.cantidad }));

      const agregarSinRepetir = (actual, nuevo) => {
        if (!nuevo) return actual;
        const partes = actual.split(",").map((s) => s.trim()).filter(Boolean);
        if (partes.some((p) => p.toLowerCase() === nuevo.toLowerCase())) return actual;
        return partes.length === 0 ? nuevo : `${actual}, ${nuevo}`;
      };

      const nuevaRef = previewImport.referenciaNum
        ? previewImport.referenciaTipo === "Pedido"
          ? previewImport.referenciaNum
          : `COT-${previewImport.referenciaNum}`
        : null;

      return {
        ...f,
        cliente: agregarSinRepetir(f.cliente, previewImport.cliente),
        pedidoRef: agregarSinRepetir(f.pedidoRef, nuevaRef),
        metodoPago: previewImport.metodo || f.metodoPago,
        lineas: [...lineasActuales, ...lineasNuevas],
      };
    });
    setPreviewImport(null);
    setMensaje({ tipo: "ok", texto: `Se agregaron ${previewImport.items.length} código${previewImport.items.length !== 1 ? "s" : ""} de la cotización. Puedes borrar o agregar líneas antes de enviar.` });
  };

  const valido = () => {
    if (!form.cliente.trim() || !form.pedidoRef.trim() || !form.fecha) return false;
    const lineasOk = form.lineas.some((l) => l.codigo.trim() && String(l.cantidad).trim());
    return lineasOk;
  };

  const enviar = async () => {
    if (!valido()) {
      setMensaje({ tipo: "error", texto: "Falta completar fecha, pedido, cliente, o al menos un código con cantidad." });
      return;
    }
    setEnviando(true);
    try {
      const lineasLimpias = form.lineas.filter((l) => l.codigo.trim() && String(l.cantidad).trim());
      const nuevo = { ...form, id: uid(), lineas: lineasLimpias.map((l) => ({ ...l, codigo: l.codigo.trim().toUpperCase(), cantidad: Number(l.cantidad), precio: null })) };
      const ok = await insertarPedido(nuevo);
      if (ok) {
        setMensaje({ tipo: "ok", texto: `Pedido ${form.pedidoRef} enviado a compras. Dale clic al botón de abajo para mandar también el correo.`, pedido: nuevo });
        setForm({ ...vacio(), vendedor });
        await refrescarMisPedidos();
        avisarNuevoPedido(nuevo);
      } else {
        setMensaje({ tipo: "error", texto: "No se pudo confirmar el envío por una conexión inestable. Tu pedido NO se guardó — vuelve a darle a \"Enviar pedido\"." });
      }
    } catch (e) {
      setMensaje({ tipo: "error", texto: "No se pudo guardar el pedido. Intenta de nuevo." });
    }
    setEnviando(false);
  };

  return (
    <div style={{ maxWidth: 760, margin: "0 auto", padding: "36px 20px 80px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
        <div className="sans" style={{ fontSize: 13, color: "#8A6E4B" }}>Conectado como <strong>{vendedor}</strong></div>
        <button onClick={onSalir} className="sans" style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", fontSize: 13, padding: 0 }}>
          <LogOut size={14} /> Cerrar sesión
        </button>
      </div>
      <h1 style={{ fontSize: 28, margin: "18px 0 4px" }}>Nuevo pedido</h1>
      <p className="sans" style={{ color: "#5C5748", fontSize: 13, marginBottom: 16 }}>Llena el detalle del pedido.</p>

      <div className="sans" style={{ display: "flex", gap: 10, alignItems: "flex-start", background: "#F0EDE4", border: "1px solid #C9C2AE", borderRadius: 4, padding: "12px 14px", marginBottom: 22, fontSize: 12.5, color: "#5C5748", lineHeight: 1.6 }}>
        <Clock size={16} style={{ flexShrink: 0, marginTop: 2, color: "#8A6E4B" }} />
        <div>
          <strong style={{ color: "#20241F" }}>Cierres de pedido:</strong> los pedidos enviados de lunes a miércoles antes de las 10:00 a.m. llegan el <strong>viernes</strong>. Los pedidos enviados desde el miércoles después de las 10:00 a.m., y durante jueves, viernes y sábado, llegan el <strong>miércoles</strong> siguiente.
        </div>
      </div>

      <div style={{ background: "#FBFAF6", border: "1px solid #C9C2AE", borderRadius: 4, padding: 22, marginBottom: 22 }}>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, marginBottom: 14 }}>
          <Campo label="Fecha">
            <input type="date" value={form.fecha} onChange={(e) => setForm({ ...form, fecha: e.target.value })} style={estiloInput} />
          </Campo>
          <Campo label="N.º de pedido">
            <input value={form.pedidoRef} onChange={(e) => setForm({ ...form, pedidoRef: e.target.value })} placeholder="Ej. PED-0231" style={estiloInput} />
          </Campo>
          <Campo label="Cliente">
            <input value={form.cliente} onChange={(e) => setForm({ ...form, cliente: e.target.value })} placeholder="Nombre del cliente" style={estiloInput} />
          </Campo>
          <Campo label="Método de pago">
            <select value={form.metodoPago} onChange={(e) => setForm({ ...form, metodoPago: e.target.value })} style={estiloInput}>
              <option value="contado">Contado</option>
              <option value="credito">Crédito</option>
            </select>
          </Campo>
        </div>

        <div className="sans" style={{ fontSize: 11, letterSpacing: 1.5, textTransform: "uppercase", color: "#8A6E4B", margin: "18px 0 2px" }}>Códigos del pedido</div>
        <div className="sans" style={{ fontSize: 12, color: "#8A8370", marginBottom: 10 }}>
          Escribe el código, Enter, escribe la cantidad, Enter — se agrega la siguiente línea sola.
        </div>

        <input ref={archivoInputRef} type="file" accept="application/pdf" onChange={manejarArchivoPdf} style={{ display: "none" }} />
        <button
          onClick={() => archivoInputRef.current?.click()}
          disabled={importando}
          className="sans"
          style={{ display: "flex", alignItems: "center", gap: 8, background: "#FBFAF6", border: "1px solid #8A6E4B", color: "#8A6E4B", borderRadius: 4, padding: "9px 14px", fontSize: 13, cursor: importando ? "default" : "pointer", marginBottom: 12, opacity: importando ? 0.6 : 1 }}
        >
          <FileUp size={15} /> {importando ? "Leyendo cotización…" : "Importar cotización (PDF)"}
        </button>

        {errorImport && (
          <div className="sans" style={{ marginBottom: 12, fontSize: 13, padding: "10px 12px", borderRadius: 4, background: "#F3DCD3", color: "#8A3B22", display: "flex", gap: 8, alignItems: "flex-start" }}>
            <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />
            {errorImport}
          </div>
        )}

        {previewImport && (
          <div style={{ border: "1px solid #8A6E4B", borderRadius: 4, background: "#FBFAF6", padding: 16, marginBottom: 16 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
              <div className="sans" style={{ fontSize: 13, color: "#20241F" }}>
                Se encontraron <strong>{previewImport.items.length}</strong> código{previewImport.items.length !== 1 ? "s" : ""}
                {previewImport.referenciaNum ? ` en ${previewImport.referenciaTipo} #${previewImport.referenciaNum}` : ""}.
                {previewImport.cliente && <> Cliente: <strong>{previewImport.cliente}</strong>.</>}
              </div>
              <button onClick={() => setPreviewImport(null)} style={{ background: "none", border: "none", color: "#8A8370", cursor: "pointer", padding: 2 }}>
                <X size={16} />
              </button>
            </div>

            {previewImport.vendedorCotizacion && previewImport.vendedorCotizacion.toLowerCase() !== vendedor.toLowerCase() && (
              <div className="sans" style={{ fontSize: 12, color: "#5C5748", background: "#F0EDE4", border: "1px solid #C9C2AE", borderRadius: 4, padding: "8px 10px", marginBottom: 10 }}>
                Nota (no impide continuar): esta cotización fue hecha a nombre de <strong>{previewImport.vendedorCotizacion}</strong>, y tú entraste como <strong>{vendedor}</strong>. Si de todos modos es tu pedido, sigue abajo con normalidad.
              </div>
            )}

            <div className="sans" style={{ fontSize: 12, color: "#5C5748", marginBottom: 6 }}>
              Revisa la lista; si alguna línea no va, quítala con la X. Luego dale clic a "Agregar":
            </div>
            <div className="mono" style={{ maxHeight: 220, overflowY: "auto", border: "1px solid #E4DFCF", borderRadius: 4 }}>
              <div className="sans" style={{ display: "grid", gridTemplateColumns: "1fr 90px 32px", gap: 10, fontSize: 11, letterSpacing: 1, textTransform: "uppercase", color: "#8A8370", padding: "6px 10px", background: "#F0EDE4" }}>
                <span>Código</span>
                <span>Cantidad</span>
                <span></span>
              </div>
              {previewImport.items.length === 0 && (
                <div className="sans" style={{ padding: "12px 10px", fontSize: 13, color: "#8A8370" }}>
                  Quitaste todas las líneas. Dale "Cancelar" o importa otro PDF.
                </div>
              )}
              {previewImport.items.map((it, i) => (
                <div key={i} style={{ display: "grid", gridTemplateColumns: "1fr 90px 32px", gap: 10, alignItems: "center", padding: "5px 10px", fontSize: 13, borderTop: "1px solid #E4DFCF" }}>
                  <span>{it.codigo}</span>
                  <span>{it.cantidad}</span>
                  <button
                    onClick={() => quitarItemPreview(i)}
                    title="Quitar esta línea"
                    style={{ background: "none", border: "none", color: "#B0553C", cursor: "pointer", padding: 4, display: "flex", justifyContent: "center" }}
                  >
                    <X size={14} />
                  </button>
                </div>
              ))}
            </div>

            <div style={{ display: "flex", gap: 10, marginTop: 14 }}>
              <button
                onClick={confirmarImportacion}
                disabled={previewImport.items.length === 0}
                className="sans"
                style={{ background: "#20241F", color: "#F0EDE4", border: "none", borderRadius: 4, padding: "9px 16px", fontSize: 13, cursor: previewImport.items.length === 0 ? "default" : "pointer", opacity: previewImport.items.length === 0 ? 0.5 : 1 }}
              >
                Agregar estos códigos al pedido
              </button>
              <button onClick={() => setPreviewImport(null)} className="sans" style={{ background: "none", border: "1px solid #C9C2AE", color: "#5C5748", borderRadius: 4, padding: "9px 16px", fontSize: 13, cursor: "pointer" }}>
                Cancelar
              </button>
            </div>
          </div>
        )}
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {form.lineas.map((l, idx) => (
            <div key={l.id} style={{ display: "grid", gridTemplateColumns: "1fr 120px 32px", gap: 8, alignItems: "center" }}>
              <input
                ref={(el) => (codigoRefs.current[l.id] = el)}
                className="mono"
                value={l.codigo}
                onChange={(e) => actualizarLinea(l.id, "codigo", e.target.value)}
                onKeyDown={enterEnCodigo(l)}
                placeholder="Código"
                style={estiloInput}
              />
              <input
                ref={(el) => (cantidadRefs.current[l.id] = el)}
                className="mono"
                type="number"
                min="0"
                value={l.cantidad}
                onChange={(e) => actualizarLinea(l.id, "cantidad", e.target.value)}
                onKeyDown={enterEnCantidad(idx, l)}
                placeholder="Cant."
                style={estiloInput}
              />
              <button onClick={() => quitarLinea(l.id)} disabled={form.lineas.length === 1} style={{ background: "none", border: "none", color: form.lineas.length === 1 ? "#C9C2AE" : "#B0553C", cursor: form.lineas.length === 1 ? "default" : "pointer", padding: 6 }}>
                <Trash2 size={16} />
              </button>
            </div>
          ))}
        </div>
        <button onClick={() => agregarLinea(true)} className="sans" style={{ marginTop: 10, display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px dashed #8A6E4B", color: "#8A6E4B", borderRadius: 4, padding: "6px 12px", fontSize: 13, cursor: "pointer" }}>
          <Plus size={14} /> Agregar código
        </button>

        {mensaje && (
          <div className="sans" style={{ marginTop: 16, fontSize: 13, padding: "10px 12px", borderRadius: 4, background: mensaje.tipo === "ok" ? "#E4E9DA" : "#F3DCD3", color: mensaje.tipo === "ok" ? "#3E5A2A" : "#8A3B22", display: "flex", flexDirection: "column", gap: 8 }}>
            <div style={{ display: "flex", gap: 8, alignItems: "flex-start" }}>
              {mensaje.tipo === "ok" ? <Check size={16} style={{ flexShrink: 0, marginTop: 1 }} /> : <AlertCircle size={16} style={{ flexShrink: 0, marginTop: 1 }} />}
              {mensaje.texto}
            </div>
            {mensaje.pedido && (
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                <a
                  href={construirMailtoPedido(mensaje.pedido)}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="sans"
                  style={{ display: "inline-flex", alignItems: "center", gap: 6, background: "#3E5A2A", color: "#F0EDE4", border: "none", borderRadius: 4, padding: "8px 14px", fontSize: 13, textDecoration: "none", cursor: "pointer" }}
                >
                  <Mail size={14} /> Enviar correo con este pedido
                </a>
                <BotonCopiar
                  pedido={mensaje.pedido}
                  estilo={{ display: "inline-flex", alignItems: "center", gap: 6, background: "none", border: "1px solid #3E5A2A", color: "#3E5A2A", borderRadius: 4, padding: "8px 14px", fontSize: 13, cursor: "pointer" }}
                />
              </div>
            )}
          </div>
        )}

        <button onClick={enviar} disabled={enviando} className="sans" style={{ marginTop: 18, width: "100%", background: "#20241F", color: "#F0EDE4", border: "none", borderRadius: 4, padding: "12px 0", fontSize: 14, letterSpacing: 0.5, cursor: enviando ? "default" : "pointer", opacity: enviando ? 0.6 : 1 }}>
          {enviando ? "Enviando…" : "Enviar pedido a compras"}
        </button>
      </div>

      <div className="sans" style={{ fontSize: 11, letterSpacing: 1.5, textTransform: "uppercase", color: "#8A6E4B", margin: "30px 0 10px" }}>Mis pedidos enviados</div>

      <div style={{ display: "flex", gap: 10, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        <label className="sans" style={{ fontSize: 12, color: "#5C5748", display: "flex", alignItems: "center", gap: 6 }}>
          Desde
          <input type="date" value={filtroDesde} onChange={(e) => setFiltroDesde(e.target.value)} style={{ ...estiloInput, width: "auto", fontSize: 13, padding: "6px 8px" }} />
        </label>
        <label className="sans" style={{ fontSize: 12, color: "#5C5748", display: "flex", alignItems: "center", gap: 6 }}>
          Hasta
          <input type="date" value={filtroHasta} onChange={(e) => setFiltroHasta(e.target.value)} style={{ ...estiloInput, width: "auto", fontSize: 13, padding: "6px 8px" }} />
        </label>
        {(filtroDesde || filtroHasta) && (
          <button onClick={() => { setFiltroDesde(""); setFiltroHasta(""); }} className="sans" style={{ background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", fontSize: 12, textDecoration: "underline", padding: 0 }}>
            Ver todos
          </button>
        )}
      </div>

      {cargandoLista ? (
        <p className="sans" style={{ fontSize: 13, color: "#8A8370" }}>Cargando…</p>
      ) : propios.length === 0 ? (
        <p className="sans" style={{ fontSize: 13, color: "#8A8370" }}>
          {filtroDesde || filtroHasta ? "No hay pedidos en ese rango de fechas." : "Aún no has enviado pedidos."}
        </p>
      ) : (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {propios.map((p) => (
            <div key={p.id} style={{ border: "1px solid #C9C2AE", borderRadius: 4, padding: "10px 14px", background: "#FBFAF6", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
              <div className="sans" style={{ fontSize: 13 }}>
                <strong className="mono">{p.pedidoRef}</strong> · {p.cliente} · {p.fecha} {horaDe(p.creadoEn)} · {p.lineas.length} código{p.lineas.length !== 1 ? "s" : ""}
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8, flexShrink: 0 }}>
                <a
                  href={construirMailtoPedido(p)}
                  target="_blank"
                  rel="noopener noreferrer"
                  title="Enviar correo de este pedido"
                  className="sans"
                  style={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "1px solid #C9C2AE", color: "#5C5748", borderRadius: 4, padding: "4px 8px", fontSize: 12, cursor: "pointer", textDecoration: "none" }}
                >
                  <Mail size={13} /> Correo
                </a>
                <BotonCopiar
                  pedido={p}
                  estilo={{ display: "flex", alignItems: "center", gap: 4, background: "none", border: "1px solid #C9C2AE", color: "#5C5748", borderRadius: 4, padding: "4px 8px", fontSize: 12, cursor: "pointer" }}
                />
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

function Campo({ label, children }) {
  return (
    <label className="sans" style={{ display: "block", fontSize: 12, color: "#5C5748" }}>
      <span style={{ display: "block", marginBottom: 4 }}>{label}</span>
      {children}
    </label>
  );
}

function BotonCopiar({ pedido, estilo }) {
  const [copiado, setCopiado] = useState(false);
  const copiar = async () => {
    const texto = `Para: ${CORREO_COMPRAS}\n\n${textoPedido(pedido)}`;
    try {
      if (navigator.clipboard.write && window.ClipboardItem) {
        const item = new ClipboardItem({
          "text/plain": new Blob([texto], { type: "text/plain" }),
          "text/html": new Blob([htmlPedido(pedido)], { type: "text/html" }),
        });
        await navigator.clipboard.write([item]);
      } else {
        await navigator.clipboard.writeText(texto);
      }
      setCopiado(true);
      setTimeout(() => setCopiado(false), 2500);
    } catch {
      try {
        await navigator.clipboard.writeText(texto);
        setCopiado(true);
        setTimeout(() => setCopiado(false), 2500);
      } catch {
        setCopiado(false);
      }
    }
  };
  return (
    <button onClick={copiar} className="sans" style={estilo}>
      {copiado ? <Check size={13} /> : <ClipboardList size={13} />} {copiado ? "Copiado" : "Copiar datos"}
    </button>
  );
}

const estiloInput = {
  width: "100%",
  boxSizing: "border-box",
  padding: "9px 10px",
  border: "1px solid #C9C2AE",
  borderRadius: 3,
  background: "#FFFFFF",
  fontSize: 14,
  color: "#20241F",
};

// ---------------- COMPRAS ----------------

const PRECIO_FIJO = 1;

function VistaCompras({ onSalir }) {
  const [pedidos, setPedidos] = useState([]);
  const [cargando, setCargando] = useState(true);
  const [filtroEstado, setFiltroEstado] = useState("pendiente");
  const [filtroVendedor, setFiltroVendedor] = useState("");
  const [filtroDesde, setFiltroDesde] = useState("");
  const [filtroHasta, setFiltroHasta] = useState("");
  const [seleccion, setSeleccion] = useState({}); // pedidoId -> bool
  const [incluirEncabezados, setIncluirEncabezados] = useState(false);
  const [guardando, setGuardando] = useState(false);
  const [aviso, setAviso] = useState(null);
  const [sonidoActivo, setSonidoActivo] = useState(false);
  const [avisoNuevos, setAvisoNuevos] = useState(0);
  const timerRef = useRef(null);
  const pendientesPrevRef = useRef(null);
  const sonidoActivoRef = useRef(false);

  const sonar = () => {
    try {
      const ctx = new (window.AudioContext || window.webkitAudioContext)();
      [0, 0.18].forEach((delay) => {
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.type = "sine";
        osc.frequency.value = 880;
        gain.gain.setValueAtTime(0.18, ctx.currentTime + delay);
        gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + delay + 0.16);
        osc.connect(gain);
        gain.connect(ctx.destination);
        osc.start(ctx.currentTime + delay);
        osc.stop(ctx.currentTime + delay + 0.16);
      });
    } catch {}
  };

  const activarSonido = () => {
    setSonidoActivo(true);
    sonidoActivoRef.current = true;
    sonar();
  };

  const cargar = useCallback(async () => {
    setCargando(true);
    const data = await leerPedidos();
    setPedidos(data);
    const pendientes = data.filter((p) => p.estado === "pendiente").length;
    if (pendientesPrevRef.current !== null && pendientes > pendientesPrevRef.current) {
      const nuevos = pendientes - pendientesPrevRef.current;
      setAvisoNuevos(nuevos);
      if (sonidoActivoRef.current) sonar();
      setTimeout(() => setAvisoNuevos(0), 6000);
    }
    pendientesPrevRef.current = pendientes;
    document.title = pendientes > 0 ? `(${pendientes}) Panel de compras` : "Panel de compras";
    setCargando(false);
  }, []);

  useEffect(() => {
    return () => {
      document.title = "Control de pedidos";
    };
  }, []);

  useEffect(() => {
    cargar();
    timerRef.current = setInterval(cargar, 20000);
    const alVolverFoco = () => {
      if (document.visibilityState === "visible") cargar();
    };
    document.addEventListener("visibilitychange", alVolverFoco);
    window.addEventListener("focus", cargar);
    return () => {
      clearInterval(timerRef.current);
      document.removeEventListener("visibilitychange", alVolverFoco);
      window.removeEventListener("focus", cargar);
    };
  }, [cargar]);

  const vendedoresUnicos = [...new Set(pedidos.map((p) => p.vendedor))].filter(Boolean).sort();

  const visibles = pedidos
    .filter((p) => (filtroEstado === "todos" ? true : p.estado === filtroEstado))
    .filter((p) => (filtroVendedor ? p.vendedor === filtroVendedor : true))
    .filter((p) => (filtroDesde ? p.fecha >= filtroDesde : true))
    .filter((p) => (filtroHasta ? p.fecha <= filtroHasta : true))
    .sort((a, b) => b.creadoEn - a.creadoEn);

  const totalSeleccionados = Object.values(seleccion).filter(Boolean).length;

  const exportar = async () => {
    const aExportar = visibles.filter((p) => seleccion[p.id]);
    if (aExportar.length === 0) {
      setAviso({ tipo: "error", texto: "Selecciona al menos un pedido para exportar." });
      return;
    }
    const filas = [];
    aExportar.forEach((p) => p.lineas.forEach((l) => {
      filas.push([l.codigo, l.cantidad, PRECIO_FIJO]);
    }));
    const filasFinales = incluirEncabezados ? [["Codigo", "Cantidad", "Precio"], ...filas] : filas;
    const ws = XLSX.utils.aoa_to_sheet(filasFinales);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Hoja1");
    XLSX.writeFile(wb, `compra_${hoy()}.xls`, { bookType: "biff8" });

    // Marcar como exportado: cada fila se actualiza por su propio id en la
    // base de datos, así que no hay riesgo de pisar un pedido nuevo que
    // otro vendedor esté guardando en ese mismo instante.
    setGuardando(true);
    const ok = await marcarPedidosExportados(aExportar);
    if (ok) {
      await cargar();
      setSeleccion({});
      setAviso({ tipo: "ok", texto: `Exportado: ${aExportar.length} pedido${aExportar.length !== 1 ? "s" : ""}, ${filas.length} línea${filas.length !== 1 ? "s" : ""}. Ya se marcaron como "Exportado".` });
    } else {
      setAviso({ tipo: "error", texto: "El Excel se descargó, pero no se pudo confirmar el marcado de \"Exportado\" por una conexión inestable. Dale a Actualizar y revisa si quedaron marcados; si no, vuelve a intentar." });
    }
    setGuardando(false);
  };

  const marcarTodos = (valor) => {
    const s = {};
    visibles.forEach((p) => (s[p.id] = valor));
    setSeleccion(s);
  };

  return (
    <div style={{ maxWidth: 980, margin: "0 auto", padding: "36px 20px 100px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 6 }}>
        <h1 style={{ fontSize: 28, margin: 0 }}>Panel de compras</h1>
        <div style={{ display: "flex", gap: 10 }}>
          <button
            onClick={activarSonido}
            className="sans"
            style={{ display: "flex", alignItems: "center", gap: 6, background: sonidoActivo ? "#20241F" : "none", color: sonidoActivo ? "#F0EDE4" : "#5C5748", border: "1px solid #C9C2AE", borderRadius: 4, padding: "6px 12px", fontSize: 12, cursor: "pointer" }}
          >
            {sonidoActivo ? "🔔 Sonido activo" : "🔕 Activar sonido"}
          </button>
          <button onClick={cargar} className="sans" style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px solid #C9C2AE", borderRadius: 4, padding: "6px 12px", fontSize: 12, color: "#5C5748", cursor: "pointer" }}>
            <RefreshCw size={13} /> Actualizar
          </button>
          <button onClick={onSalir} className="sans" style={{ display: "flex", alignItems: "center", gap: 6, background: "none", border: "1px solid #C9C2AE", borderRadius: 4, padding: "6px 12px", fontSize: 12, color: "#5C5748", cursor: "pointer" }}>
            <LogOut size={13} /> Cerrar sesión
          </button>
        </div>
      </div>

      {avisoNuevos > 0 && (
        <div className="sans stamp" style={{ marginBottom: 16, fontSize: 14, padding: "10px 14px", borderRadius: 4, background: "#F3DCD3", color: "#8A3B22", border: "1px solid #B0553C" }}>
          🔔 {avisoNuevos === 1 ? "Llegó un pedido nuevo." : `Llegaron ${avisoNuevos} pedidos nuevos.`}
        </div>
      )}

      <p className="sans" style={{ color: "#5C5748", fontSize: 13, marginBottom: 24 }}>
        Datos compartidos: todos los pedidos enviados por los vendedores aparecen aquí. El precio de exportación es fijo en $1.00 — el precio real se actualiza después directamente en el sistema.
      </p>

      <div style={{ display: "flex", gap: 10, marginBottom: 10, flexWrap: "wrap", alignItems: "center" }}>
        <select value={filtroEstado} onChange={(e) => setFiltroEstado(e.target.value)} className="sans" style={{ ...estiloInput, width: "auto", fontSize: 13 }}>
          <option value="pendiente">Pendientes</option>
          <option value="exportado">Exportados</option>
          <option value="todos">Todos</option>
        </select>
        <select value={filtroVendedor} onChange={(e) => setFiltroVendedor(e.target.value)} className="sans" style={{ ...estiloInput, width: "auto", fontSize: 13 }}>
          <option value="">Todos los vendedores</option>
          {vendedoresUnicos.map((v) => <option key={v} value={v}>{v}</option>)}
        </select>
        <label className="sans" style={{ fontSize: 12, color: "#5C5748", display: "flex", alignItems: "center", gap: 6, marginLeft: "auto" }}>
          <input type="checkbox" checked={incluirEncabezados} onChange={(e) => setIncluirEncabezados(e.target.checked)} />
          Incluir encabezados en el Excel
        </label>
      </div>

      <div style={{ display: "flex", gap: 10, marginBottom: 18, flexWrap: "wrap", alignItems: "center" }}>
        <label className="sans" style={{ fontSize: 12, color: "#5C5748", display: "flex", alignItems: "center", gap: 6 }}>
          Desde
          <input type="date" value={filtroDesde} onChange={(e) => setFiltroDesde(e.target.value)} style={{ ...estiloInput, width: "auto", fontSize: 13, padding: "6px 8px" }} />
        </label>
        <label className="sans" style={{ fontSize: 12, color: "#5C5748", display: "flex", alignItems: "center", gap: 6 }}>
          Hasta
          <input type="date" value={filtroHasta} onChange={(e) => setFiltroHasta(e.target.value)} style={{ ...estiloInput, width: "auto", fontSize: 13, padding: "6px 8px" }} />
        </label>
        {(filtroDesde || filtroHasta) && (
          <button onClick={() => { setFiltroDesde(""); setFiltroHasta(""); }} className="sans" style={{ background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", fontSize: 12, textDecoration: "underline", padding: 0 }}>
            Quitar rango de fechas
          </button>
        )}
        <span className="sans" style={{ fontSize: 12, color: "#8A8370", marginLeft: "auto" }}>
          {visibles.length} pedido{visibles.length !== 1 ? "s" : ""}
        </span>
      </div>

      {aviso && (
        <div className="sans" style={{ marginBottom: 16, fontSize: 13, padding: "10px 12px", borderRadius: 4, background: aviso.tipo === "ok" ? "#E4E9DA" : "#F3DCD3", color: aviso.tipo === "ok" ? "#3E5A2A" : "#8A3B22" }}>
          {aviso.texto}
        </div>
      )}

      {cargando ? (
        <p className="sans" style={{ fontSize: 13, color: "#8A8370" }}>Cargando pedidos…</p>
      ) : visibles.length === 0 ? (
        <div className="sans" style={{ textAlign: "center", padding: "50px 0", color: "#8A8370" }}>
          <PackageSearch size={28} style={{ marginBottom: 8 }} />
          <div>No hay pedidos que coincidan con este filtro.</div>
        </div>
      ) : (
        <>
          <div className="sans" style={{ display: "flex", gap: 14, marginBottom: 8, fontSize: 12, color: "#8A6E4B" }}>
            <button onClick={() => marcarTodos(true)} style={{ background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", padding: 0, textDecoration: "underline" }}>Seleccionar todo</button>
            <button onClick={() => marcarTodos(false)} style={{ background: "none", border: "none", color: "#8A6E4B", cursor: "pointer", padding: 0, textDecoration: "underline" }}>Quitar selección</button>
          </div>

          <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
            {visibles.map((p) => (
              <div key={p.id} style={{ border: "1px solid #C9C2AE", borderRadius: 4, background: "#FBFAF6" }}>
                <div className="sans" style={{ display: "flex", alignItems: "center", gap: 10, padding: "10px 14px", borderBottom: "1px solid #E4DFCF", fontSize: 13 }}>
                  <input type="checkbox" checked={!!seleccion[p.id]} onChange={(e) => setSeleccion({ ...seleccion, [p.id]: e.target.checked })} />
                  <strong className="mono">{p.pedidoRef}</strong>
                  <span style={{ color: "#5C5748" }}>{p.cliente}</span>
                  <span style={{ color: "#8A8370" }}>{p.fecha} {horaDe(p.creadoEn)}</span>
                  <span style={{ color: "#8A8370" }}>{p.vendedor}</span>
                  <span style={{ color: "#8A8370", textTransform: "uppercase", fontSize: 11 }}>{p.metodoPago}</span>
                  <span className="stamp" style={{ marginLeft: "auto", fontSize: 11, letterSpacing: 1, textTransform: "uppercase", padding: "3px 9px", borderRadius: 3, border: `1px solid ${p.estado === "exportado" ? "#3E5A2A" : "#8A6E4B"}`, color: p.estado === "exportado" ? "#3E5A2A" : "#8A6E4B" }}>
                    {p.estado === "exportado" ? "Exportado" : "Pendiente"}
                  </span>
                </div>
                <div style={{ padding: "10px 14px" }}>
                  <div className="sans" style={{ display: "grid", gridTemplateColumns: "1fr 90px", gap: 10, fontSize: 11, letterSpacing: 1, textTransform: "uppercase", color: "#8A8370", padding: "0 0 4px" }}>
                    <span>Código</span>
                    <span>Cantidad</span>
                  </div>
                  {p.lineas.map((l) => (
                    <div key={l.id} className="mono" style={{ display: "grid", gridTemplateColumns: "1fr 90px", gap: 10, alignItems: "center", padding: "5px 0", fontSize: 13 }}>
                      <span>{l.codigo}</span>
                      <span>{l.cantidad}</span>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>

          <div style={{ position: "sticky", bottom: 16, marginTop: 24, display: "flex", gap: 10, justifyContent: "flex-end" }}>
            <button onClick={exportar} disabled={guardando} className="sans" style={{ display: "flex", alignItems: "center", gap: 8, background: "#20241F", color: "#F0EDE4", border: "none", borderRadius: 4, padding: "10px 18px", fontSize: 13, cursor: "pointer" }}>
              <Download size={15} /> Exportar seleccionados ({totalSeleccionados}) · $1.00 c/u
            </button>
          </div>
        </>
      )}
    </div>
  );
}
