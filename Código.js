const ID_HOJA = SpreadsheetApp.getActiveSpreadsheet().getId();
const HOJA_MAESTRA = "Maestra";

const MAPA_CATEGORIAS = {
  "Nuestro Talento": ["Gestión del Cuidado por la Vida.", "Desarrollo Organizacional para soportar la Mega 2030.", "Talento global, innovador y sostenible."],
  "Nuestras Marcas": ["Marcas Líderes – Share mercado.", "Desarrollo B2B Foco USA + México", "Contexto regulatorio Global."],
  "Nuestras Capacidades para la Entrega de Valor": ["Gestión de la rentabilidad.", "Gestión de capital de trabajo.", "Recuperación Volumen"],
  "Proyecto Quantum": ["Material de empaque", "Materias Primas", "Desarrollo Logístico"]
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Registro Proyectos - Ser Extraordinario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Valida el acceso de forma estricta.
 */
function validarAcceso(cedulaRecibida) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const hoja = ss.getSheetByName(HOJA_MAESTRA);

  if (!hoja) throw new Error("No se encontró la pestaña 'Maestra'.");

  const datos = hoja.getDataRange().getValues();
  const cedulaLimpia = cedulaRecibida.toString().trim();

  // Buscamos omitiendo la cabecera
  const filaUsuario = datos.slice(1).find(f => f[0].toString().trim() === cedulaLimpia);

  if (!filaUsuario) {
    throw new Error("Cédula " + cedulaLimpia + " no encontrada en la base.");
  }

  // Mapeo explícito para asegurar que nada llegue null al cliente
  // Se asume Gerencia en columna E (índice 4)
  const resultado = {
    usuario: {
      cedula: filaUsuario[0],
      nombre: filaUsuario[1] || "Usuario",
      gerencia: filaUsuario[4] || "General", 
      cargo: filaUsuario[2] || "Colaborador"
    },
    integrantes: obtenerIntegrantes(datos),
    categorias: MAPA_CATEGORIAS,
    proyectos: obtenerProyectosPorGerencia(filaUsuario[4] || "General")
  };

  return resultado;
}

function obtenerIntegrantes(datos) {
  return datos.slice(1)
    .filter(f => f[1] && f[4] && f[3])
    .map(f => `${f[1]} - ${f[4]} - ${f[3]}`);
}

/**
 * Obtiene proyectos filtrando por gerencia con normalización de texto.
 */
function obtenerProyectosPorGerencia(gerencia) {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const hoja = ss.getSheetByName("Registros");
    
    // Si la hoja no existe o no tiene datos (solo cabecera), retornar arreglo vacío
    if (!hoja || hoja.getLastRow() < 2) return [];

    const datos = hoja.getDataRange().getValues();
    const gerenciaBusqueda = (gerencia || "").toString().trim();

    return datos.slice(1)
      .filter(f => {
        // Comparación segura: índice 5 es la columna de Gerencia en Registros
        return f[5] && f[5].toString().trim() === gerenciaBusqueda;
      })
      .map(f => ({
        titulo: f[8], // Columna I
        autor: f[3],  // Columna D
        // Convertimos la fecha a String ISO para que el cliente la reciba sin errores
        fecha: f[0] instanceof Date ? f[0].toISOString() : f[0] 
      }))
      .reverse(); // Los más recientes primero
      
  } catch (e) {
    console.error("Error en obtenerProyectosPorGerencia: " + e.message);
    return []; // SIEMPRE retornar un arreglo, incluso en error
  }
}

function guardarNuevoProyecto(datos) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let hoja = ss.getSheetByName("Registros");

  if (!hoja) {
    hoja = ss.insertSheet("Registros");
    hoja.appendRow(["Fecha", "Correo", "Cédula", "Nombre", "Celular", "Gerencia", "Categoría", "Subcategoría", "Título", "Descripción", "Implementación", "Beneficio", "Moneda", "I1", "I2", "I3", "I4", "I5", "I6", "I7", "I8", "I9", "I10", "Obs"]);
  }

  const nuevaFila = [
    new Date(), 
    Session.getActiveUser().getEmail(), 
    datos.cedula, 
    datos.nombre, 
    datos.celular, 
    datos.gerencia, 
    datos.categoria, 
    datos.subcategoria, 
    datos.titulo, 
    datos.descripcion, 
    datos.fechaImpl, 
    datos.beneficio, 
    datos.moneda, 
    ...datos.integrantes, 
    datos.observaciones
  ];
  
  hoja.appendRow(nuevaFila);
  // Retorna la lista actualizada para la gerencia del usuario
  return obtenerProyectosPorGerencia(datos.gerencia);
}