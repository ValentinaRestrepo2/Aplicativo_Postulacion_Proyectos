const ID_HOJA = SpreadsheetApp.getActiveSpreadsheet().getId();
const HOJA_MAESTRA = "Maestra";

const MAPA_CATEGORIAS = {
  "Nuestro Talento": ["Gestión del Cuidado por la Vida.", "Desarrollo Organizacional para soportar la Mega 2030.", "Talento global, innovador y sostenible."],
  "Nuestras Marcas": ["Marcas Líderes – Share mercado.", "Desarrollo B2B Global", "Contexto regulatorio Global."],
  "Nuestras Capacidades para la Entrega de Valor": ["Gestión de la rentabilidad.", "Gestión de capital de trabajo.", "Recuperación Volumen"],
  "Proyecto Quantum": ["Revenue growth managment", "Go to market", "Manufactura", "Supply Chain", "Compras"]
};

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Registro Proyectos - Ser Extraordinario')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function validarAcceso(cedulaRecibida) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  const hoja = ss.getSheetByName(HOJA_MAESTRA);

  if (!hoja) throw new Error("No se encontró la pestaña 'Maestra'.");

  const datos = hoja.getDataRange().getValues();
  const cedulaLimpia = cedulaRecibida.toString().trim();

  const filaUsuario = datos.slice(1).find(f => f[0].toString().trim() === cedulaLimpia);

  if (!filaUsuario) {
    throw new Error("Cédula " + cedulaLimpia + " no encontrada en la base.");
  }

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


function obtenerProyectosPorGerencia(gerencia) {
  try {
    const ss = SpreadsheetApp.openById(ID_HOJA);
    const hoja = ss.getSheetByName("Registros");

    if (!hoja || hoja.getLastRow() < 2) return [];

    const datos = hoja.getDataRange().getValues();
    const gerenciaBusqueda = (gerencia || "").toString().trim();

    return datos.slice(1)
      .filter(f => {
        return f[5] && f[5].toString().trim() === gerenciaBusqueda;
      })
      .map(f => ({
        titulo: f[8],
        autor: f[3],
        fecha: f[0] instanceof Date ? f[0].toISOString() : f[0]
      }))
      .reverse();

  } catch (e) {
    console.error("Error en obtenerProyectosPorGerencia: " + e.message);
    return [];
  }
}

function guardarNuevoProyecto(datos) {
  const ss = SpreadsheetApp.openById(ID_HOJA);
  let hoja = ss.getSheetByName("Registros");

  if (!hoja) {
    hoja = ss.insertSheet("Registros");
    hoja.appendRow(["Fecha", "Correo", "Cédula", "Nombre", "Celular", "Gerencia", "Categoría", "Subcategoría", "Título", "Descripción", "Implementación", "Beneficio", "Moneda", "Intangible", "I1", "I2", "I3", "I4", "I5", "I6", "I7", "I8", "I9", "I10", "Obs"]);
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
    datos.intangible,
    ...datos.integrantes,
    datos.observaciones
  ];

  hoja.appendRow(nuevaFila);
  try {
    const colorAzul = "#1c3d5a";
    const colorVerde = "#4d8d8a";
    const colorOro = "#c9a43b";
    const correoUsuario = Session.getActiveUser().getEmail();
    Logger.log(correoUsuario)

    const primerNombre = datos.nombre.split(' ')[0];
    const asunto = `¡Tu proyecto "${datos.titulo}" ya está postulado!`;
    const html = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; color: #333; max-width: 600px; margin: auto; border: 1px solid #eee; border-radius: 15px; overflow: hidden;">
      <div style="background-color: ${colorAzul}; padding: 30px; text-align: center;">
        <h1 style="color: white; margin: 0; font-size: 24px;">¡${primerNombre}!</h1>
      </div>
      
      <div style="padding: 30px; line-height: 1.6;">
        <p style="font-size: 18px; color: ${colorVerde};"><b>Tu proyecto ya fue postulado a Ser Extraordinario.</b></p>
        <p>Gracias por ser parte de <b>Ser Extraordinario</b>. Personas como tú, que buscan transformar y agregar valor, son las que hacen la diferencia en nuestra compañía.</p>
        
        <div style="background-color: #f9f9f9; border-left: 5px solid ${colorOro}; padding: 20px; margin: 25px 0;">
          <h3 style="margin-top: 0; color: ${colorAzul};">Resumen de tu postulación:</h3>
          <ul style="list-style: none; padding: 0;">
            <li><b>Fecha de postulación:</b> ${new Date().toLocaleDateString()}</li>
            <li><b>Proyecto:</b> ${datos.titulo}</li>
            <li><b>Categoría:</b> ${datos.categoria}</li>
            <li><b>Subcategoría:</b> ${datos.subcategoria}</li>
            <li><b>Generación de valor (Económico):</b>$ ${datos.beneficio} ${datos.moneda}</li>
            <li><b>Generación de valor (No económico):</b> ${datos.intangible}</li>
            <li><b>Descripción:</b> ${datos.descripcion}</li>
          </ul>
        </div>
        
        <p>¿Qué sigue? El equipo encargado revisará tu propuesta y nos pondremos en contacto contigo si necesitamos más detalles.</p>
        
      </div>
      
      <div style="background-color: #f4f4f4; padding: 15px; text-align: center; font-size: 12px; color: #999;">
        Este es un mensaje automático del Sistema de Postulación de Proyectos.
      </div>
    </div>
  `;

    GmailApp.sendEmail(correoUsuario, asunto, "", {
      htmlBody: html
    });
    Logger.log("Mensaje enviado")
  } catch (e) {
    console.error("No se pudo enviar el correo: " + e.message);
  }
  return obtenerProyectosPorGerencia(datos.gerencia);
}