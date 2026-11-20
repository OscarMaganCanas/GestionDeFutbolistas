const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.title = "Gestión de Futbolistas";

// ── Paleta ──
const VERDE_OSCURO  = "1A472A";
const VERDE_MED     = "2D6A4F";
const DORADO        = "F0B429";
const BLANCO        = "FFFFFF";
const GRIS_CLARO    = "F5F5F5";
const GRIS_TEXTO    = "374151";
const GRIS_SUAVE    = "D1D5DB";

// ── Helpers ──
const makeShadow = () => ({ type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.18 });

function addContentSlide(numSlide, title) {
  let s = pres.addSlide();
  s.background = { color: GRIS_CLARO };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 1.0, fill: { color: VERDE_OSCURO }, line: { color: VERDE_OSCURO } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.12, h: 1.0, fill: { color: DORADO }, line: { color: DORADO } });
  s.addShape(pres.shapes.OVAL, { x: 9.2, y: 0.1, w: 0.7, h: 0.7, fill: { color: DORADO }, line: { color: DORADO } });
  s.addText(String(numSlide), { x: 9.2, y: 0.1, w: 0.7, h: 0.7, fontSize: 14, bold: true, color: VERDE_OSCURO, align: "center", valign: "middle", margin: 0 });
  s.addText(title, { x: 0.3, y: 0.1, w: 8.7, h: 0.8, fontSize: 24, bold: true, color: BLANCO, fontFace: "Calibri", valign: "middle", margin: 0 });
  return s;
}

function addCard(slide, x, y, w, h, headerText, headerColor, bodyLines) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h: 0.4, fill: { color: headerColor }, line: { color: headerColor } });
  slide.addText(headerText, { x: x + 0.12, y, w: w - 0.15, h: 0.4, fontSize: 13, bold: true, color: BLANCO, valign: "middle", margin: 0 });
  if (bodyLines && bodyLines.length > 0) {
    const items = bodyLines.map((line, i) => ({
      text: line,
      options: { bullet: true, breakLine: i < bodyLines.length - 1 }
    }));
    slide.addText(items, { x: x + 0.12, y: y + 0.45, w: w - 0.2, h: h - 0.55, fontSize: 13, color: GRIS_TEXTO, fontFace: "Calibri", valign: "top" });
  }
}

// ══════════════════════════════════════
// SLIDE 0 – PORTADA
// ══════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: VERDE_OSCURO };
  s.addShape(pres.shapes.OVAL, { x: 6.5, y: -0.8, w: 6, h: 6, fill: { color: VERDE_MED, transparency: 50 }, line: { color: VERDE_MED, transparency: 50 } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.8, w: 4, h: 4, fill: { color: VERDE_MED, transparency: 65 }, line: { color: VERDE_MED, transparency: 65 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: DORADO }, line: { color: DORADO } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 2.5, w: 9.2, h: 0.04, fill: { color: DORADO, transparency: 50 }, line: { color: DORADO, transparency: 50 } });

  s.addText("GESTION DE FUTBOLISTAS", {
    x: 0.5, y: 0.7, w: 9, h: 1.1,
    fontSize: 40, bold: true, color: BLANCO, fontFace: "Calibri", charSpacing: 3, margin: 0
  });
  s.addText("Aplicacion de escritorio Java + MySQL", {
    x: 0.5, y: 1.85, w: 9, h: 0.65,
    fontSize: 20, color: DORADO, fontFace: "Calibri", italic: true, margin: 0
  });
  s.addText("Presentacion tecnica del proyecto", {
    x: 0.5, y: 2.65, w: 9, h: 0.55,
    fontSize: 16, color: GRIS_SUAVE, fontFace: "Calibri", margin: 0
  });

  const temas = ["Problema", "Funcionalidades", "Interfaz", "Arquitectura", "Modelo datos", "Codigo", "Demo en vivo"];
  const cols = [0.5, 2.1, 3.7, 5.3, 6.9, 0.5, 2.1];
  const rows = [3.45, 3.45, 3.45, 3.45, 3.45, 4.35, 4.35];
  temas.forEach((t, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: cols[i], y: rows[i], w: 1.45, h: 0.55, fill: { color: VERDE_MED }, line: { color: DORADO, pt: 1 } });
    s.addText((i+1) + ". " + t, { x: cols[i], y: rows[i], w: 1.45, h: 0.55, fontSize: 11, color: BLANCO, align: "center", valign: "middle", bold: true, margin: 0 });
  });
}

// ══════════════════════════════════════
// SLIDE 1 – PROBLEMA A SOLUCIONAR
// ══════════════════════════════════════
{
  let s = addContentSlide(1, "1. Problema a solucionar");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.15, w: 9.3, h: 1.2, fill: { color: VERDE_OSCURO }, line: { color: VERDE_OSCURO }, shadow: makeShadow() });
  s.addText([
    { text: "Contexto: ", options: { bold: true, color: DORADO } },
    { text: "Los clubes de futbol necesitan registrar y consultar informacion de sus jugadores y equipos de forma rapida y organizada.", options: { color: BLANCO } }
  ], { x: 0.5, y: 1.2, w: 9.0, h: 1.1, fontSize: 15, fontFace: "Calibri", valign: "middle" });

  const cards = [
    { title: "Problema", lines: ["Sin herramienta dedicada, los datos se gestionan en papel o en hojas de calculo sin control.", "Dificil consultar, modificar o borrar registros."] },
    { title: "Necesidad cubierta", lines: ["Aplicacion CRUD completa para equipos y futbolistas.", "Interfaz grafica sencilla, sin conocimientos tecnicos previos."] },
    { title: "Solucion", lines: ["App Java de escritorio con base de datos MySQL.", "Permite altas, consultas, modificaciones y bajas desde pantalla."] },
  ];
  const xs = [0.35, 3.55, 6.75];
  cards.forEach((c, i) => {
    addCard(s, xs[i], 2.5, 3.0, 2.7, c.title, i === 2 ? VERDE_MED : VERDE_OSCURO, c.lines);
  });
}

// ══════════════════════════════════════
// SLIDE 2 – FUNCIONALIDADES
// ══════════════════════════════════════
{
  let s = addContentSlide(2, "2. Funcionalidades desarrolladas");

  const colsData = [
    {
      title: "Gestion de Equipos",
      color: VERDE_OSCURO,
      x: 0.35,
      lines: [
        "Anadir nuevo equipo (nombre, estadio, presupuesto)",
        "Ver todos los equipos en tabla",
        "Modificar datos de un equipo",
        "Eliminar un equipo",
      ]
    },
    {
      title: "Gestion de Futbolistas",
      color: VERDE_MED,
      x: 5.15,
      lines: [
        "Anadir jugador (nombre, posicion, dorsal, equipo)",
        "Ver todos los jugadores en tabla",
        "Modificar datos de un jugador",
        "Eliminar un jugador",
      ]
    }
  ];

  colsData.forEach(col => {
    s.addShape(pres.shapes.RECTANGLE, { x: col.x, y: 1.1, w: 4.5, h: 4.1, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: col.x, y: 1.1, w: 4.5, h: 0.5, fill: { color: col.color }, line: { color: col.color } });
    s.addText(col.title, { x: col.x + 0.12, y: 1.1, w: 4.25, h: 0.5, fontSize: 15, bold: true, color: BLANCO, valign: "middle", margin: 0 });
    const items = col.lines.map((line, i) => ({
      text: line,
      options: { bullet: true, breakLine: i < col.lines.length - 1 }
    }));
    s.addText(items, { x: col.x + 0.15, y: 1.7, w: 4.2, h: 3.3, fontSize: 14, color: GRIS_TEXTO, fontFace: "Calibri", valign: "top" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.0, w: 9.3, h: 0.45, fill: { color: DORADO }, line: { color: DORADO } });
  s.addText("Validaciones: nombre sin numeros  *  dorsal entre 1 y 99  *  campos obligatorios  *  posiciones por enum", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.45, fontSize: 13, bold: true, color: VERDE_OSCURO, valign: "middle", margin: 0
  });
}

// ══════════════════════════════════════
// SLIDE 3 – INTERFAZ GRAFICA
// ══════════════════════════════════════
{
  let s = addContentSlide(3, "3. Interfaz Grafica");

  const screens = [
    { label: "Menu Principal", desc: ["Boton: Gestionar Equipos", "Boton: Gestionar Futbolistas"], x: 0.35, y: 1.15, color: VERDE_OSCURO },
    { label: "Ventana Equipos", desc: ["Tabla de equipos", "Campos: nombre, estadio, presupuesto", "Botones: Guardar / Modificar / Eliminar / Volver"], x: 3.55, y: 1.15, color: VERDE_MED },
    { label: "Ventana Futbolistas", desc: ["Tabla de futbolistas", "Campos: nombre, dorsal", "ComboBox: posicion y equipo", "Botones: Guardar / Modificar / Eliminar / Volver"], x: 6.75, y: 1.15, color: "1B4332" },
  ];

  screens.forEach(sc => {
    addCard(s, sc.x, sc.y, 3.0, 2.8, sc.label, sc.color, sc.desc);
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 4.1, w: 9.3, h: 1.3, fill: { color: VERDE_OSCURO }, line: { color: VERDE_OSCURO }, shadow: makeShadow() });
  s.addText("Navegabilidad", { x: 0.5, y: 4.1, w: 2.5, h: 0.45, fontSize: 14, bold: true, color: DORADO, margin: 0 });

  s.addText([
    { text: "Menu Principal", options: { bold: true, color: DORADO } },
    { text: "  ->  ", options: { color: BLANCO } },
    { text: "Ventana Equipos", options: { bold: true, color: GRIS_SUAVE } },
    { text: "  /  ", options: { color: BLANCO } },
    { text: "Ventana Futbolistas", options: { bold: true, color: GRIS_SUAVE } },
    { text: "  ->  ", options: { color: BLANCO } },
    { text: "Boton Volver", options: { bold: true, color: DORADO } },
    { text: "  ->  ", options: { color: BLANCO } },
    { text: "Menu Principal", options: { bold: true, color: GRIS_SUAVE } },
  ], { x: 0.5, y: 4.55, w: 9.0, h: 0.7, fontSize: 14, valign: "middle" });
}

// ══════════════════════════════════════
// SLIDE 4 – ARQUITECTURA MVC + DAO
// ══════════════════════════════════════
{
  let s = addContentSlide(4, "4. Arquitectura utilizada: MVC + DAO");

  const capas = [
    { name: "VISTA  (View)", color: VERDE_OSCURO, y: 1.1, desc: "VentanaPrincipal · VentanaEquipos · VentanaFutbolistas — Interfaz grafica en Java Swing. Solo muestra datos y captura eventos del usuario." },
    { name: "CONTROLADOR  (Controller)", color: VERDE_MED, y: 2.3, desc: "ControladorMenuPrincipal · ControladorEquipos · ControladorFutbolistas — Escuchan los eventos de la vista, validan datos y llaman al DAO." },
    { name: "MODELO  (Model)", color: "1B4332", y: 3.5, desc: "Equipo · Futbolista — Clases POJO que representan los datos (atributos + getters/setters)." },
    { name: "DAO  (Data Access Object)", color: "374151", y: 4.7, desc: "EquipoDAO · FutbolistaDAO · ConexionDB — Encapsulan toda la logica SQL. Metodos: insertar, listar, modificar, eliminar." },
  ];

  capas.forEach(capa => {
    s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: capa.y, w: 9.3, h: 0.95, fill: { color: capa.color }, line: { color: capa.color }, shadow: makeShadow() });
    s.addText(capa.name, { x: 0.5, y: capa.y + 0.05, w: 3.2, h: 0.4, fontSize: 13, bold: true, color: DORADO, margin: 0 });
    s.addText(capa.desc, { x: 0.5, y: capa.y + 0.44, w: 9.1, h: 0.45, fontSize: 12, color: GRIS_SUAVE, fontFace: "Calibri", margin: 0 });
  });
}

// ══════════════════════════════════════
// SLIDE 5 – MODELO DE DATOS
// ══════════════════════════════════════
{
  let s = addContentSlide(5, "5. Modelo de datos");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.15, w: 4.3, h: 3.8, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.15, w: 4.3, h: 0.5, fill: { color: VERDE_OSCURO }, line: { color: VERDE_OSCURO } });
  s.addText("Tabla: equipo", { x: 0.5, y: 1.15, w: 4.1, h: 0.5, fontSize: 15, bold: true, color: BLANCO, valign: "middle", margin: 0 });

  const headerRow = [
    [{ text: "Campo", options: { bold: true, color: BLANCO, fill: { color: VERDE_MED } } },
     { text: "Tipo", options: { bold: true, color: BLANCO, fill: { color: VERDE_MED } } },
     { text: "Descripcion", options: { bold: true, color: BLANCO, fill: { color: VERDE_MED } } }]
  ];

  const equipoCols = [
    ["ideEquipo", "INT", "PK - AUTO_INCREMENT"],
    ["nombre", "VARCHAR(100)", "Nombre del club"],
    ["estadio", "VARCHAR(100)", "Nombre del estadio"],
    ["presupuesto", "DOUBLE", "Presupuesto en euros"],
  ];
  const equipoRows = equipoCols.map(r => r.map(cell => ({ text: cell, options: { color: GRIS_TEXTO } })));
  s.addTable([...headerRow, ...equipoRows], {
    x: 0.5, y: 1.75, w: 4.0, colW: [1.35, 1.15, 1.5],
    border: { pt: 0.5, color: GRIS_SUAVE },
    fill: { color: BLANCO },
    fontSize: 11, fontFace: "Calibri",
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.15, w: 4.3, h: 3.8, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.35, y: 1.15, w: 4.3, h: 0.5, fill: { color: VERDE_MED }, line: { color: VERDE_MED } });
  s.addText("Tabla: futbolista", { x: 5.5, y: 1.15, w: 4.1, h: 0.5, fontSize: 15, bold: true, color: BLANCO, valign: "middle", margin: 0 });

  const futCols = [
    ["ideJugador", "INT", "PK - AUTO_INCREMENT"],
    ["nombre", "VARCHAR(100)", "Nombre del jugador"],
    ["posicion", "VARCHAR(50)", "DELANTERO, MEDIO..."],
    ["dorsal", "INT", "Numero (1-99)"],
    ["ideEquipo", "INT", "FK -> equipo.ideEquipo"],
  ];
  const futRows = futCols.map((r, ri) => r.map((cell, ci) => ({
    text: cell,
    options: { color: (ri === 4 && ci === 0) ? DORADO : GRIS_TEXTO, bold: (ri === 4 && ci === 0) }
  })));
  s.addTable([...headerRow, ...futRows], {
    x: 5.5, y: 1.75, w: 4.0, colW: [1.35, 1.15, 1.5],
    border: { pt: 0.5, color: GRIS_SUAVE },
    fill: { color: BLANCO },
    fontSize: 11, fontFace: "Calibri",
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.1, w: 9.3, h: 0.38, fill: { color: DORADO }, line: { color: DORADO } });
  s.addText("Relacion: futbolista.ideEquipo (FK)  ->  equipo.ideEquipo (PK)  |  Un equipo puede tener muchos futbolistas (1:N)", {
    x: 0.5, y: 5.1, w: 9.0, h: 0.38, fontSize: 12, bold: true, color: VERDE_OSCURO, valign: "middle", margin: 0
  });
}

// ══════════════════════════════════════
// SLIDE 6 – ESTRUCTURA DEL CODIGO
// ══════════════════════════════════════
{
  let s = addContentSlide(6, "6. Estructura del codigo");

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 1.1, w: 4.8, h: 4.3, fill: { color: "1E293B" }, line: { color: "1E293B" }, shadow: makeShadow() });
  s.addText("Arbol de paquetes (Eclipse)", { x: 0.5, y: 1.12, w: 4.6, h: 0.45, fontSize: 14, bold: true, color: DORADO, margin: 0 });

  const tree = [
    { text: "src/", options: { bold: true, color: DORADO, breakLine: true } },
    { text: "  +-- main/  -->  Main.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  +-- modelo/", options: { color: BLANCO, breakLine: true } },
    { text: "  |      +-- Equipo.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- Futbolista.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  +-- vista/", options: { color: BLANCO, breakLine: true } },
    { text: "  |      +-- VentanaPrincipal.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- VentanaEquipos.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- VentanaFutbolistas.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- EnumeradosPosiciones/", options: { color: BLANCO, breakLine: true } },
    { text: "  |              +-- Posiciones.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  +-- controlador/", options: { color: BLANCO, breakLine: true } },
    { text: "  |      +-- ControladorMenuPrincipal.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- ControladorEquipos.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  |      +-- ControladorFutbolistas.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "  +-- dao/", options: { color: BLANCO, breakLine: true } },
    { text: "         +-- ConexionDB.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "         +-- EquipoDAO.java", options: { color: GRIS_SUAVE, breakLine: true } },
    { text: "         +-- FutbolistaDAO.java", options: { color: GRIS_SUAVE } },
  ];
  s.addText(tree, { x: 0.5, y: 1.65, w: 4.55, h: 3.65, fontSize: 10, fontFace: "Consolas", valign: "top" });

  const clases = [
    { name: "Main.java", desc: "Punto de entrada. Crea VentanaPrincipal y su controlador.", color: VERDE_MED },
    { name: "ConexionDB.java", desc: "Gestiona la conexion JDBC a MySQL (localhost:3306/gestion_futbol).", color: VERDE_OSCURO },
    { name: "Posiciones.java", desc: "Enum: DELANTERO, EXTREMOS, MEDIO, CENTRALES, LATERALES, PORTERO.", color: "374151" },
    { name: "ControladorFutbolistas", desc: "Implementa ActionListener. Valida, llama al DAO y refresca la tabla.", color: "1B4332" },
  ];
  clases.forEach((c, i) => {
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.1 + i * 1.1, w: 4.25, h: 0.95, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.4, y: 1.1 + i * 1.1, w: 0.1, h: 0.95, fill: { color: c.color }, line: { color: c.color } });
    s.addText(c.name, { x: 5.55, y: 1.1 + i * 1.1 + 0.06, w: 4.05, h: 0.35, fontSize: 13, bold: true, color: c.color, margin: 0 });
    s.addText(c.desc, { x: 5.55, y: 1.1 + i * 1.1 + 0.46, w: 4.05, h: 0.42, fontSize: 11, color: GRIS_TEXTO, fontFace: "Calibri", margin: 0 });
  });
}

// ══════════════════════════════════════
// SLIDE 7 – DEMO EN VIVO
// ══════════════════════════════════════
{
  let s = addContentSlide(7, "7. Demo en vivo");

  const pasos = [
    { num: "1", titulo: "ALTA", desc: "Rellenar los campos del formulario (nombre, posicion, dorsal y equipo) y pulsar Guardar. La fila aparece al instante en la tabla.", color: VERDE_OSCURO },
    { num: "2", titulo: "CONSULTA", desc: "La tabla muestra en tiempo real todos los registros de la BD. Al hacer clic en una fila, los datos se cargan en el formulario automaticamente.", color: VERDE_MED },
    { num: "3", titulo: "BORRADO", desc: "Seleccionar un registro en la tabla, comprobar que el ID se carga en el campo y pulsar Eliminar. La fila desaparece de la tabla y de la BD.", color: "1B4332" },
  ];

  pasos.forEach((p, i) => {
    const x = 0.35 + i * 3.2;
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 3.0, h: 3.8, fill: { color: BLANCO }, line: { color: GRIS_SUAVE, pt: 1 }, shadow: makeShadow() });
    s.addShape(pres.shapes.RECTANGLE, { x, y: 1.1, w: 3.0, h: 1.1, fill: { color: p.color }, line: { color: p.color } });
    s.addText(p.num, { x, y: 1.1, w: 1.0, h: 1.1, fontSize: 48, bold: true, color: DORADO, align: "center", valign: "middle", margin: 0 });
    s.addText(p.titulo, { x: x + 0.95, y: 1.3, w: 1.9, h: 0.7, fontSize: 22, bold: true, color: BLANCO, valign: "middle", margin: 0 });
    s.addText(p.desc, { x: x + 0.12, y: 2.3, w: 2.75, h: 2.45, fontSize: 13, color: GRIS_TEXTO, fontFace: "Calibri", valign: "top" });
  });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.35, y: 5.05, w: 9.3, h: 0.42, fill: { color: VERDE_OSCURO }, line: { color: VERDE_OSCURO } });
  s.addText("BD: gestion_futbol  |  Tablas: equipo - futbolista  |  Conector: mysql-connector-j-9.7.0.jar", {
    x: 0.5, y: 5.05, w: 9.0, h: 0.42, fontSize: 12, color: GRIS_SUAVE, valign: "middle", margin: 0
  });
}

// ══════════════════════════════════════
// SLIDE 8 – CIERRE
// ══════════════════════════════════════
{
  let s = pres.addSlide();
  s.background = { color: VERDE_OSCURO };
  s.addShape(pres.shapes.OVAL, { x: 6.0, y: -1.0, w: 6, h: 6, fill: { color: VERDE_MED, transparency: 55 }, line: { color: VERDE_MED, transparency: 55 } });
  s.addShape(pres.shapes.OVAL, { x: -1.5, y: 3.5, w: 4, h: 4, fill: { color: VERDE_MED, transparency: 65 }, line: { color: VERDE_MED, transparency: 65 } });
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 0.15, h: 5.625, fill: { color: DORADO }, line: { color: DORADO } });

  s.addText("Gracias!", {
    x: 0.5, y: 1.4, w: 9, h: 1.2,
    fontSize: 52, bold: true, color: BLANCO, fontFace: "Calibri", charSpacing: 4, margin: 0
  });
  s.addText("Preguntas?", {
    x: 0.5, y: 2.7, w: 9, h: 0.8,
    fontSize: 28, color: DORADO, fontFace: "Calibri", italic: true, margin: 0
  });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 3.65, w: 9.0, h: 0.04, fill: { color: DORADO, transparency: 40 }, line: { color: DORADO, transparency: 40 } });
  s.addText("Gestion de Futbolistas  *  Java + Swing + MySQL  *  Patron MVC + DAO", {
    x: 0.5, y: 3.8, w: 9, h: 0.6,
    fontSize: 15, color: GRIS_SUAVE, fontFace: "Calibri", margin: 0
  });
}

// ── Guardar ──
pres.writeFile({ fileName: "D:\\CLAUDE\\gestionFutbolistas\\GestionFutbolistas_Presentacion.pptx" })
  .then(() => console.log("Presentacion guardada correctamente."))
  .catch(err => console.error("Error:", err));
