const express = require("express");
const { exec } = require("child_process");
const { promisify } = require("util");
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");

const ejecutarComando = promisify(exec);
const app = express();

const PUERTO = 3000;
const URL_OCUPACION_AULAS = "https://academia.unibague.edu.co/aulas.php";
const ARCHIVO_CATALOGO = path.join(
    __dirname,
    "Data",
    "Aulas_2026_catalogo_agente.xlsx"
);

const BLOQUES_PERMITIDOS = new Set([
    "BLQ1",
    "BLQ2",
    "BLQ3",
    "BLQ4",
    "BLQ5",
    "BLQ6",
    "BLQ7",
    "BING",
    "CEP",
    "BDER",
    "BLAB"
]);

const CARACTERISTICAS_ADMITIDAS = new Set([
    "sillas",
    "mesas",
    "hibrido",
    "experiencial",
    "laboratorio",
    "sala_computo",
    "sala_redes",
    "sala_satelital",
    "sala_audiencias",
    "consultorio_juridico",
    "pantalla",
    "video_beam",
     "aire_acondicionado"
]);

let procesando = false;

app.use(express.json());

function normalizarTexto(valor) {
    return String(valor ?? "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[_–—-]+/g, " ")
        .replace(/\s+/g, " ");
}

function normalizarCodigo(valor) {
    return String(valor ?? "")
        .trim()
        .toUpperCase()
        .replace(/[–—]/g, "-")
        .replace(/\s+/g, "");
}

function coincideCodigoPortal(etiquetaPortal, codigoCatalogo) {
    const esperado = normalizarCodigo(codigoCatalogo);
    const recibido = normalizarCodigo(etiquetaPortal);

    if (recibido === esperado) {
        return true;
    }

    // Algunos encabezados del portal agregan texto como el bloque o la
    // capacidad alrededor del codigo del salon. Se busca el codigo completo
    // sin permitir coincidencias parciales como 6-1 dentro de 6-12.
    const etiqueta = String(etiquetaPortal ?? "")
        .toUpperCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[–—]/g, "-");
    const patronCodigo = esperado
        .split("-")
        .map(parte => parte.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"))
        .join("\\s*-\\s*");

    return new RegExp(
        `(^|[^A-Z0-9])${patronCodigo}(?![A-Z0-9])`
    ).test(etiqueta);
}

function esSi(valor) {
    return ["si", "true", "1", "x"].includes(normalizarTexto(valor));
}

function convertirEnLista(valor) {
    if (Array.isArray(valor)) {
        return valor.flatMap(convertirEnLista);
    }

    if (valor === undefined || valor === null || valor === "") {
        return [];
    }

    const texto = String(valor).trim();

    // Algunos modelos locales envian una lista como texto JSON, por ejemplo
    // '["mesas"]'. Convertirla aqui evita tratar los corchetes y las comillas
    // como parte del nombre de la caracteristica.
    if (texto.startsWith("[") && texto.endsWith("]")) {
        try {
            const lista = JSON.parse(texto);

            if (Array.isArray(lista)) {
                return lista.flatMap(convertirEnLista);
            }
        } catch {
            // Si no es JSON valido, se procesa como texto separado por comas.
        }
    }

    return texto
        .split(/[,;]+/)
        .map(elemento => elemento.trim())
        .filter(Boolean);
}

function normalizarCaracteristica(valor) {
    const texto = normalizarTexto(valor);

    const equivalencias = {
        silla: "sillas",
        sillas: "sillas",
        pupitre: "sillas",
        pupitres: "sillas",
        mesa: "mesas",
        mesas: "mesas",
        hibrido: "hibrido",
        experiencial: "experiencial",
        experienciales: "experiencial",
        experimental: "experiencial",
        laboratorio: "laboratorio",
        laboratorios: "laboratorio",
        lab: "laboratorio",
        "sala de computo": "sala_computo",
        "sala computo": "sala_computo",
        computo: "sala_computo",
        computadores: "sala_computo",
        "sala de redes": "sala_redes",
        "sala redes": "sala_redes",
        redes: "sala_redes",
        "sala satelital": "sala_satelital",
        satelital: "sala_satelital",
        "sala de audiencias": "sala_audiencias",
        audiencias: "sala_audiencias",
        "consultorio juridico": "consultorio_juridico",
        consultorio: "consultorio_juridico",
        pantalla: "pantalla",
        "video beam": "video_beam",
        videobeam: "video_beam",
        proyector: "video_beam",
        "aire acondicionado": "aire_acondicionado",
        aire: "aire_acondicionado",
        acondicionado: "aire_acondicionado",
        climatizado: "aire_acondicionado",
        aire_acondicionado: "aire_acondicionado"
    };

    return equivalencias[texto] || null;
}

function leerCaracteristicas(body) {
    const obligatoriasCrudas = [
        ...convertirEnLista(body.caracteristicas),
        ...convertirEnLista(body.tipologias),
        ...convertirEnLista(body.tipoSalon)
    ];
    const alternativasCrudas = convertirEnLista(
        body.cualquieraDe ?? body.cualquiera_de
    );

    const desconocidas = [];

    function normalizarLista(lista) {
        const resultado = [];

        for (const elemento of lista) {
            // Los modelos pequenos pueden repetir por error la capacidad
            // dentro de caracteristicas (por ejemplo, "mesas,30"). La
            // capacidad ya se valida en su propio campo, por lo que se ignora.
            if (/^\d+\s*(personas?)?$/i.test(String(elemento).trim())) {
                continue;
            }

            const caracteristica = normalizarCaracteristica(elemento);

            if (!caracteristica || !CARACTERISTICAS_ADMITIDAS.has(caracteristica)) {
                desconocidas.push(elemento);
                continue;
            }

            if (!resultado.includes(caracteristica)) {
                resultado.push(caracteristica);
            }
        }

        return resultado;
    }

    return {
        obligatorias: normalizarLista(obligatoriasCrudas),
        alternativas: normalizarLista(alternativasCrudas),
        desconocidas
    };
}

function normalizarFecha(valor) {
    const fecha = String(valor ?? "").replace(/\D/g, "");

    if (!/^\d{8}$/.test(fecha)) {
        throw new Error("La fecha debe tener el formato AAAAMMDD o AAAA-MM-DD");
    }

    const anio = Number(fecha.slice(0, 4));
    const mes = Number(fecha.slice(4, 6));
    const dia = Number(fecha.slice(6, 8));
    const comprobacion = new Date(Date.UTC(anio, mes - 1, dia));

    if (
        comprobacion.getUTCFullYear() !== anio ||
        comprobacion.getUTCMonth() + 1 !== mes ||
        comprobacion.getUTCDate() !== dia
    ) {
        throw new Error("La fecha indicada no es valida");
    }

    return fecha;
}

function leerFechas(body) {
    const esValorFechaUtil = valor => {
        if (valor === undefined || valor === null) {
            return false;
        }

        if (Array.isArray(valor)) {
            return valor.some(esValorFechaUtil);
        }

        const texto = String(valor).trim().toLowerCase();
        return Boolean(texto) && texto !== "null" && texto !== "undefined";
    };
    const valorFechas = [
        body.fechas,
        body.fechasConsulta,
        body.fechas_consulta,
        body.fecha
    ].find(esValorFechaUtil);
    let valores = [];

    if (Array.isArray(valorFechas)) {
        valores = valorFechas;
    } else if (valorFechas !== undefined && valorFechas !== null) {
        const texto = String(valorFechas).trim();

        if (texto.startsWith("[") && texto.endsWith("]")) {
            try {
                const lista = JSON.parse(texto);
                valores = Array.isArray(lista) ? lista : [texto];
            } catch {
                valores = texto.split(/[,;\n]+/);
            }
        } else if (texto) {
            valores = texto.split(/[,;\n]+/);
        }
    }

    const fechas = [...new Set(
        valores
            .map(valor => String(valor ?? "").trim())
            .filter(Boolean)
            .map(normalizarFecha)
    )];

    if (fechas.length === 0) {
        throw new Error("Debe indicar al menos una fecha");
    }

    if (fechas.length > 5) {
        throw new Error("Se permiten maximo 5 fechas por consulta");
    }

    return fechas;
}

const MESES = {
    enero: 1, febrero: 2, marzo: 3, abril: 4, mayo: 5, junio: 6,
    julio: 7, agosto: 8, septiembre: 9, setiembre: 9, octubre: 10,
    noviembre: 11, diciembre: 12
};

const DIAS_SEMANA = {
    domingo: 0, domingos: 0, lunes: 1, martes: 2, miercoles: 3,
    jueves: 4, viernes: 5, sabado: 6, sabados: 6
};

function fechaAaaaMmDd(anio, mes, dia) {
    return `${String(anio).padStart(4, "0")}` +
        `${String(mes).padStart(2, "0")}` +
        `${String(dia).padStart(2, "0")}`;
}

function fechaColombiaActual() {
    const partes = new Intl.DateTimeFormat("en-CA", {
        timeZone: "America/Bogota", year: "numeric", month: "2-digit",
        day: "2-digit"
    }).formatToParts(new Date());
    const valores = Object.fromEntries(
        partes.filter(parte => parte.type !== "literal")
            .map(parte => [parte.type, parte.value])
    );
    return fechaAaaaMmDd(valores.year, valores.month, valores.day);
}

function sumarDiasFecha(fecha, cantidad) {
    const anio = Number(fecha.slice(0, 4));
    const mes = Number(fecha.slice(4, 6));
    const dia = Number(fecha.slice(6, 8));
    const resultado = new Date(Date.UTC(anio, mes - 1, dia + cantidad));
    return fechaAaaaMmDd(
        resultado.getUTCFullYear(), resultado.getUTCMonth() + 1,
        resultado.getUTCDate()
    );
}

function extraerFechasExplicitas(texto) {
    const resultados = [];
    const patron = /\b(\d{4})[-\/]([01]?\d)[-\/]([0-3]?\d)\b|\b([0-3]?\d)[-\/]([01]?\d)[-\/](\d{4})\b/g;
    for (const coincidencia of texto.matchAll(patron)) {
        const anio = coincidencia[1] || coincidencia[6];
        const mes = coincidencia[2] || coincidencia[5];
        const dia = coincidencia[3] || coincidencia[4];
        resultados.push(normalizarFecha(fechaAaaaMmDd(anio, mes, dia)));
    }
    return resultados;
}

function fechasDeDiasDelMes(diasSemana, mes, anio) {
    const resultado = [];
    const ultimoDia = new Date(Date.UTC(anio, mes, 0)).getUTCDate();
    for (let dia = 1; dia <= ultimoDia; dia += 1) {
        const fecha = new Date(Date.UTC(anio, mes - 1, dia));
        if (diasSemana.includes(fecha.getUTCDay())) {
            resultado.push(fechaAaaaMmDd(anio, mes, dia));
        }
    }
    return resultado;
}

function extraerFechasDelMensaje(valor) {
    const texto = normalizarTexto(valor);
    const textoFechas = String(valor ?? "")
        .trim()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ");
    if (!texto) return [];

    const hoy = fechaColombiaActual();
    if (/\bpasado manana\b/.test(texto)) return [sumarDiasFecha(hoy, 2)];
    if (/\bmanana\b/.test(texto)) return [sumarDiasFecha(hoy, 1)];
    if (/\bhoy\b/.test(texto)) return [hoy];

    const recurrencia = texto.match(
        /\b(?:todos?|cada)\s+los?\s+(.+?)\s+de\s+([a-z]+)\s+(?:de\s+)?(\d{4})\b/
    );
    if (recurrencia) {
        const dias = Object.entries(DIAS_SEMANA)
            .filter(([nombre]) => new RegExp(`\\b${nombre}\\b`).test(recurrencia[1]))
            .map(([, numero]) => numero);
        const mes = MESES[recurrencia[2]];
        const anio = Number(recurrencia[3]);
        if (dias.length > 0 && mes) {
            return [...new Set(fechasDeDiasDelMes(dias, mes, anio))];
        }
    }

    const explicitas = extraerFechasExplicitas(textoFechas);
    const contieneRango = (
        /\b(?:del|desde)\b/.test(texto) && /\b(?:al|hasta)\b/.test(texto)
    ) || (/\bentre\b/.test(texto) && /\by\b/.test(texto));
    if (contieneRango && explicitas.length >= 2) {
        const resultado = [];
        let cursor = explicitas[0];
        const final = explicitas[1];
        const diasSolicitados = [...new Set(
            Object.entries(DIAS_SEMANA)
                .filter(([nombre]) => new RegExp(`\\b${nombre}\\b`).test(texto))
                .map(([, numero]) => numero)
        )];
        if (cursor > final) {
            throw new Error("La fecha final debe ser posterior a la fecha inicial");
        }
        let diasRecorridos = 0;
        while (cursor <= final && diasRecorridos <= 366) {
            const fechaCursor = new Date(Date.UTC(
                Number(cursor.slice(0, 4)),
                Number(cursor.slice(4, 6)) - 1,
                Number(cursor.slice(6, 8))
            ));
            if (
                diasSolicitados.length === 0 ||
                diasSolicitados.includes(fechaCursor.getUTCDay())
            ) {
                resultado.push(cursor);
            }
            if (resultado.length > 5) {
                throw new Error("Se permiten maximo 5 fechas por consulta");
            }
            cursor = sumarDiasFecha(cursor, 1);
            diasRecorridos += 1;
        }
        return resultado;
    }
    return [...new Set(explicitas)];
}

function extraerCapacidadDelMensaje(valor) {
    const coincidencia = normalizarTexto(valor)
        .match(/\b(?:para\s+)?(\d+)\s+personas?\b/);
    return coincidencia ? Number(coincidencia[1]) : null;
}

function extraerCaracteristicasDelMensaje(valor) {
    const texto = normalizarTexto(valor);
    if (!texto) return [];
    const patrones = [
        ["aire_acondicionado", /\b(?:aire acondicionado|salon con aire|aire|climatizado|acondicionado)\b/],
        ["sala_computo", /\b(?:sala de computo|sala computo|computadores)\b/],
        ["video_beam", /\b(?:video beam|videobeam|proyector)\b/],
        ["experiencial", /\b(?:experiencial|experienciales|experimental)\b/],
        ["laboratorio", /\b(?:laboratorio|laboratorios|lab)\b/],
        ["hibrido", /\bhibrido\b/], ["pantalla", /\bpantalla\b/],
        ["mesas", /\b(?:mesa|mesas)\b/],
        ["sillas", /\b(?:silla|sillas|pupitre|pupitres)\b/]
    ];
    return patrones.filter(([, patron]) => patron.test(texto))
        .map(([caracteristica]) => caracteristica);
}

function normalizarHora(valor, nombreCampo) {
    const hora = String(valor ?? "")
        .trim()
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, "");
    const coincidencia = hora.match(
        /^(\d{1,2})(?::(\d{2}))?\s*(am|pm)?$/
    );

    if (!coincidencia) {
        throw new Error(
            `${nombreCampo} debe tener el formato HH:MM o hh:mm am/pm`
        );
    }

    let horas = Number(coincidencia[1]);
    const minutos = Number(coincidencia[2] ?? 0);
    const periodo = coincidencia[3] || "";

    if (
        (periodo && (horas < 1 || horas > 12)) ||
        (!periodo && horas > 23) ||
        minutos > 59 ||
        minutos % 30 !== 0
    ) {
        throw new Error(
            `${nombreCampo} debe ser una hora valida y usar minutos 00 o 30`
        );
    }

    if (periodo === "am" && horas === 12) {
        horas = 0;
    } else if (periodo === "pm" && horas !== 12) {
        horas += 12;
    }

    return `${String(horas).padStart(2, "0")}:${String(minutos).padStart(2, "0")}`;
}

function minutosDesdeMedianoche(hora) {
    const [horas, minutos] = hora.split(":").map(Number);
    return horas * 60 + minutos;
}

function contieneRangoHorarioExplicito(valor) {
    const texto = normalizarTexto(valor).replace(/\./g, "");
    const horasConMinutos = texto.match(
        /\b(?:[01]?\d|2[0-3]):[0-5]\d\b/g
    ) || [];

    // Dos horas en formato HH:MM son suficientes para identificar el rango.
    // Tambien se aceptan expresiones como "de 8 a 10" o "8 am hasta 10 am".
    if (horasConMinutos.length >= 2) {
        return true;
    }

    return /\b(?:de\s+|desde\s+)?(?:[01]?\d|2[0-3])\s*(?:am|pm)?\s+(?:a|hasta)\s+(?:[01]?\d|2[0-3])\s*(?:am|pm)?\b/.test(
        texto
    );
}

function extraerHorarioDelMensaje(valor) {
    const texto = String(valor ?? "")
        .toLowerCase()
        .replace(/\./g, "")
        .replace(/\s+/g, " ")
        .trim();

    // Prioriza horas con minutos. Esto evita confundir los componentes de
    // una fecha (por ejemplo 2026-09-24) con horas del rango solicitado.
    const coincidencias = [...texto.matchAll(
        /\b([01]?\d|2[0-3]):([0-5]\d)\s*(am|pm)?\b/g
    )];

    if (coincidencias.length >= 2) {
        return {
            inicio: normalizarHora(
                `${coincidencias[0][1]}:${coincidencias[0][2]}` +
                `${coincidencias[0][3] || ""}`,
                "horaInicio"
            ),
            fin: normalizarHora(
                `${coincidencias[1][1]}:${coincidencias[1][2]}` +
                `${coincidencias[1][3] || ""}`,
                "horaFin"
            )
        };
    }

    // Tambien admite rangos sin minutos: "de 8 a 10" o
    // "desde las 8 am hasta las 10 am".
    const rango = texto.match(
        /\b(?:de|desde)?\s*(?:las?\s+)?(\d{1,2})\s*(am|pm)?\s*(?:a|hasta)\s*(?:las?\s+)?(\d{1,2})\s*(am|pm)?\b/
    );

    if (!rango) {
        return null;
    }

    return {
        inicio: normalizarHora(`${rango[1]}${rango[2] || ""}`, "horaInicio"),
        fin: normalizarHora(`${rango[3]}${rango[4] || ""}`, "horaFin")
    };
}

function horasEsperadas(horaInicio, horaFin) {
    const resultado = [];

    for (
        let minuto = minutosDesdeMedianoche(horaInicio);
        minuto < minutosDesdeMedianoche(horaFin);
        minuto += 30
    ) {
        const horas = Math.floor(minuto / 60);
        const minutos = minuto % 60;
        resultado.push(
            `${String(horas).padStart(2, "0")}:${String(minutos).padStart(2, "0")}`
        );
    }

    return resultado;
}

function cargarCatalogo() {
    if (!fs.existsSync(ARCHIVO_CATALOGO)) {
        throw new Error(
            `No se encontro el catalogo de aulas en ${ARCHIVO_CATALOGO}`
        );
    }

    const libro = XLSX.readFile(ARCHIVO_CATALOGO);
    const hoja = libro.Sheets.Catalogo_agente;

    if (!hoja) {
        throw new Error(
            "El Excel no contiene la hoja requerida: Catalogo_agente"
        );
    }

    return XLSX.utils.sheet_to_json(hoja, {
        defval: "",
        raw: false
    });
}

function escaparArgumento(valor) {
    const texto = String(valor);

    if (texto.includes('"')) {
        throw new Error("Se intento ejecutar un argumento no permitido");
    }

    return `"${texto}"`;
}

async function ejecutarAgentBrowser(argumentos, opciones = {}) {
    const comando = ["agent-browser", ...argumentos.map(escaparArgumento)].join(" ");

    return ejecutarComando(comando, {
        maxBuffer: 20 * 1024 * 1024,
        timeout: 120000,
        ...opciones
    });
}

function parsearSalidaEval(salida) {
    let valor = String(salida ?? "").trim();

    if (!valor) {
        throw new Error("Agent-browser no devolvio datos de disponibilidad");
    }

    for (let intento = 0; intento < 2; intento += 1) {
        if (typeof valor !== "string") {
            break;
        }

        try {
            valor = JSON.parse(valor);
        } catch {
            break;
        }
    }

    if (!valor || typeof valor !== "object") {
        throw new Error("No fue posible interpretar la respuesta del portal");
    }

    return valor;
}

function esperar(milisegundos) {
    return new Promise(resolve => setTimeout(resolve, milisegundos));
}

async function consultarDisponibilidadPortal(fecha, horaInicio, horaFin) {
    const sesion = `salones-api-${process.pid}-${Date.now()}`;

    try {
        await ejecutarAgentBrowser([
            "--session",
            sesion,
            "open",
            URL_OCUPACION_AULAS
        ], { timeout: 15000 });

        const prepararFormulario =
            "(()=>{const campos=[...document.querySelectorAll('input[type=text]')];" +
            "if(!campos.length){throw new Error('No se encontro el campo de fecha')}" +
            `campos[0].value='${fecha}';` +
            "campos[0].dispatchEvent(new Event('input',{bubbles:true}));" +
            "campos[0].dispatchEvent(new Event('change',{bubbles:true}));" +
            "const casilla=document.querySelector('input[type=checkbox]');" +
            "if(!casilla){throw new Error('No se encontro Ver solo el dia')}" +
            "casilla.checked=true;" +
            "casilla.dispatchEvent(new Event('change',{bubbles:true}));" +
            "return {fecha:campos[0].value,soloDia:casilla.checked}})()";

        await ejecutarAgentBrowser([
            "--session",
            sesion,
            "eval",
            prepararFormulario
        ], { timeout: 8000 });

        await ejecutarAgentBrowser([
            "--session",
            sesion,
            "find",
            "role",
            "button",
            "click",
            "--name",
            "Consultar"
        ], { timeout: 10000 });

        const extraerDisponibles =
            "(()=>{const tablas=[...document.querySelectorAll('table')];" +
            "const tabla=tablas.find(t=>t.rows.length>1&&t.rows[0].cells.length>2);" +
            "if(!tabla){return {error:'El portal no mostro la tabla de ocupacion'," +
            "url:location.href,titulo:document.title," +
            "tablas:tablas.map(t=>({filas:t.rows.length,columnas:t.rows[0]?.cells.length||0}))," +
            "texto:(document.body?.innerText||'').slice(0,500)}}" +
            // Reconstruir la cuadricula logica respeta rowspan y colspan. Sin
            // esto, una clase que ocupa varias franjas desplaza los indices de
            // las columnas y puede hacer que se revise otro salon por error.
            "const matriz=[];[...tabla.rows].forEach((fila,r)=>{" +
            "matriz[r]??=[];let col=0;[...fila.cells].forEach(celda=>{" +
            "while(matriz[r][col]!==undefined){col++}" +
            "const rs=Math.max(1,celda.rowSpan||1);" +
            "const cs=Math.max(1,celda.colSpan||1);" +
            "for(let dr=0;dr<rs;dr++){matriz[r+dr]??=[];" +
            "for(let dc=0;dc<cs;dc++){matriz[r+dr][col+dc]=celda}}" +
            "col+=cs})});" +
            "const cabeceras=matriz[0]||[];" +
            "const filas=matriz.slice(1).filter(f=>{" +
            "const hora=f[1]?.innerText.trim()||'';" +
            `return hora>='${horaInicio}'&&hora<'${horaFin}'});` +
            "const salones=cabeceras.slice(2).filter((c,i)=>" +
            "filas.length>0&&filas.every(f=>!f[i+2]?.innerText.trim()))" +
            ".map(c=>c.innerText.trim()).filter(Boolean);" +
            "return {horas:filas.map(f=>f[1].innerText.trim()),salones}})()";

        let disponibilidad;

        const maximoIntentos = 4;

        for (let intento = 1; intento <= maximoIntentos; intento += 1) {
            // La primera espera permite que termine la navegacion iniciada por
            // el boton Consultar. Los siguientes intentos son mas cortos.
            await esperar(intento === 1 ? 4000 : 2000);

            try {
                const respuesta = await ejecutarAgentBrowser([
                    "--session",
                    sesion,
                    "eval",
                    extraerDisponibles
                ], { timeout: 6000 });

                disponibilidad = parsearSalidaEval(respuesta.stdout);

                if (!disponibilidad.error) {
                    console.log(
                        `Tabla de ocupacion leida en el intento ${intento}.`
                    );
                    return disponibilidad;
                }

                console.log(
                    `Esperando la tabla de ocupacion. Intento ${intento} ` +
                    `de ${maximoIntentos}...`
                );
            } catch (error) {
                console.log(
                    `La lectura del portal no respondio en el intento ` +
                    `${intento} de ${maximoIntentos}. Reintentando...`
                );

                if (intento === maximoIntentos) {
                    throw new Error(
                        "Agent-browser no respondio al intentar leer la " +
                        "tabla de ocupacion"
                    );
                }
            }
        }

        const diagnostico = disponibilidad
            ? ` URL: ${disponibilidad.url || 'desconocida'}. ` +
              `Tablas: ${JSON.stringify(disponibilidad.tablas || [])}. ` +
              `Contenido: ${disponibilidad.texto || 'sin contenido'}`
            : "";

        throw new Error(
            "El portal no mostro la tabla de ocupacion despues de esperar." +
            diagnostico
        );
    } finally {
        // No se espera el cierre de la sesion para responder a n8n. Esto evita
        // sumar demoras del navegador a una consulta que ya termino.
        ejecutarAgentBrowser(
            ["--session", sesion, "close"],
            { timeout: 5000 }
        ).catch(() => {
            // El cierre de la sesion no debe ocultar el resultado de la consulta.
        });
    }
}

function filtrarCatalogo(
    catalogo,
    salonesLibres,
    capacidadMinima,
    obligatorias,
    alternativas,
    codigoSalon = ""
) {
    const etiquetasDisponibles = salonesLibres.map(String);
    const caracteristicasEspecializadas = new Set([
        "laboratorio",
        "sala_computo",
        "sala_redes",
        "sala_satelital"
    ]);
    const necesitaLaboratorio = [...obligatorias, ...alternativas]
        .some(caracteristica => caracteristicasEspecializadas.has(caracteristica));

    const candidatos = catalogo.filter(aula => {
        const codigoBloque = normalizarCodigo(aula.codigo_bloque);
        const codigoPortal = normalizarCodigo(aula.codigo_portal);
        const capacidad = Number(aula.capacidad);

        if (codigoSalon && codigoPortal !== codigoSalon) {
            return false;
        }

        if (!esSi(aula.incluido_agente)) {
            return false;
        }

        if (!BLOQUES_PERMITIDOS.has(codigoBloque)) {
            return false;
        }

        if (
            !codigoSalon &&
            esSi(aula.solo_si_laboratorio) &&
            !necesitaLaboratorio
        ) {
            return false;
        }

        if (capacidadMinima > 0 && (!capacidad || capacidad < capacidadMinima)) {
            return false;
        }

        if (!obligatorias.every(caracteristica => esSi(aula[caracteristica]))) {
            return false;
        }

        if (
            alternativas.length > 0 &&
            !alternativas.some(caracteristica => esSi(aula[caracteristica]))
        ) {
            return false;
        }

        return true;
    });

    const salones = candidatos
        .filter(aula => etiquetasDisponibles.some(etiqueta =>
            coincideCodigoPortal(etiqueta, aula.codigo_portal)
        ))
        .map(aula => {
            const caracteristicas = {
               sillas: esSi(aula.sillas),
    mesas: esSi(aula.mesas),
    hibrido: esSi(aula.hibrido),
    experiencial: esSi(aula.experiencial),
    laboratorio: esSi(aula.laboratorio),
    sala_computo: esSi(aula.sala_computo),
    sala_redes: esSi(aula.sala_redes),
    sala_satelital: esSi(aula.sala_satelital),
    sala_audiencias: esSi(aula.sala_audiencias),
    pantalla: esSi(aula.pantalla),
    video_beam: esSi(aula.video_beam),
    aire_acondicionado: esSi(aula.aire_acondicionado)
 };

            return {
                codigo: String(aula.codigo_portal).trim(),
                bloque: aula.bloque,
                codigo_bloque: aula.codigo_bloque,
                capacidad: Number(aula.capacidad) || null,
                tipo_espacio: aula.tipo_espacio || null,
                mobiliario_base: aula.mobiliario_base || null,
                // Entregar solo los valores verdaderos evita que modelos
                // pequenos confundan una caracteristica falsa con una
                // disponible al redactar la respuesta final.
                caracteristicas: Object.entries(caracteristicas)
                    .filter(([, disponible]) => disponible)
                    .map(([nombre]) => nombre),
                requiere_revision: esSi(aula.requiere_revision),
                observaciones: aula.observaciones || null
            };
        })
        .sort((a, b) => {
            if ((a.capacidad ?? Infinity) !== (b.capacidad ?? Infinity)) {
                return (a.capacidad ?? Infinity) - (b.capacidad ?? Infinity);
            }

            return a.codigo.localeCompare(b.codigo, "es", { numeric: true });
        });

    return {
        candidatosCatalogo: candidatos.length,
        salones
    };
}

function buscarAlternativasSalon(
    catalogo,
    salonesLibres,
    aulaSolicitada,
    codigoSalon,
    limite = 5
) {
    const capacidadReferencia = Number(aulaSolicitada.capacidad) || 0;
    const caracteristicasReferencia = [
        "sillas",
        "mesas",
        "laboratorio",
        "sala_computo",
        "sala_redes",
        "sala_satelital"
    ].filter(caracteristica => esSi(aulaSolicitada[caracteristica]));

    // Para un aula convencional se conserva el mobiliario (sillas o mesas).
    // Para un laboratorio, ademas se exige su tipologia especializada. Los
    // laboratorios siguen excluidos cuando el aula solicitada no lo es.
    const resultado = filtrarCatalogo(
        catalogo,
        salonesLibres,
        capacidadReferencia,
        caracteristicasReferencia,
        [],
        ""
    );

    return {
        capacidad_referencia: capacidadReferencia,
        caracteristicas_referencia: caracteristicasReferencia,
        candidatosCatalogo: resultado.candidatosCatalogo,
        salones: resultado.salones
            .filter(salon => normalizarCodigo(salon.codigo) !== codigoSalon)
            .slice(0, limite)
    };
}

function compactarSalonParaAgente(salon) {
    return {
        codigo: salon.codigo,
        bloque: salon.bloque,
        capacidad: salon.capacidad,
        caracteristicas: salon.caracteristicas
    };
}

function compactarResultadoParaAgente(resultado) {
    const respuesta = {
        fecha: resultado.fecha,
        message: resultado.message,
        // Limitar la salida reduce de forma importante los tokens enviados al
        // segundo llamado de Ollama sin alterar la busqueda realizada.
        salones: (resultado.salones || [])
            .slice(0, 5)
            .map(compactarSalonParaAgente)
    };

    if (typeof resultado.disponible === "boolean") {
        respuesta.disponible = resultado.disponible;
    }

    if (Array.isArray(resultado.alternativas)) {
        respuesta.alternativas = resultado.alternativas
            .slice(0, 5)
            .map(compactarSalonParaAgente);
    }

    return respuesta;
}

function formatearSalonParaUsuario(salon) {
    const capacidad = salon.capacidad ?? "sin información";
    const caracteristicas = Array.isArray(salon.caracteristicas)
        ? salon.caracteristicas.join(", ")
        : salon.caracteristicas || "sin información";

    return [
        `- Salón: ${salon.codigo || "sin información"}`,
        `  - Bloque: ${salon.bloque || "sin información"}`,
        `  - Capacidad: ${capacidad} personas`,
        `  - Características: ${caracteristicas}`
    ].join("\n");
}

function construirRespuestaUsuario(resultados, codigoSalon) {
    return resultados.map(resultado => {
        const lineas = [resultado.message + "."];

        if (codigoSalon) {
            if (
                resultado.disponible === false &&
                Array.isArray(resultado.alternativas) &&
                resultado.alternativas.length > 0
            ) {
                lineas.push("", "Salones alternativos:", "");
                lineas.push(
                    resultado.alternativas
                        .slice(0, 5)
                        .map(formatearSalonParaUsuario)
                        .join("\n")
                );
            }
        } else if (
            Array.isArray(resultado.salones) &&
            resultado.salones.length > 0
        ) {
            lineas.push("");
            lineas.push(
                resultado.salones
                    .slice(0, 5)
                    .map(formatearSalonParaUsuario)
                    .join("\n")
            );
        }

        return lineas.join("\n");
    }).join("\n\n");
}

function mensajeValidacionParaUsuario(mensaje) {
    if (mensaje === "horaFin debe ser posterior a horaInicio") {
        return "La hora de finalización debe ser posterior a la hora de " +
            "inicio. Por favor, indique un horario válido.";
    }

    if (mensaje === "Se permiten maximo 5 fechas por consulta") {
        return "La consulta contiene más de cinco fechas. Por favor, " +
            "divida el periodo en grupos de máximo cinco fechas.";
    }

    return mensaje;
}

app.get("/", (req, res) => {
    res.json({
        status: "ok",
        message: "Servidor de scraping y consulta de salones funcionando",
        processing: procesando,
        endpoints: ["/scrape", "/consultar-salones"]
    });
});

// Endpoint existente para el proyecto de empleos. Se conserva su funcionamiento.
app.post("/scrape", async (req, res) => {
    const { url, instruction } = req.body;

    console.log("Solicitud recibida desde n8n");
    console.log("URL:", url);
    console.log("Instruccion:", instruction);

    if (!url) {
        return res.status(400).json({
            status: "error",
            message: "No se recibio ninguna URL"
        });
    }

    // Evita ejecutar dos solicitudes simultaneamente.
    if (procesando) {
        return res.status(429).json({
            status: "busy",
            message: "Ya existe una solicitud de scraping en proceso"
        });
    }

    procesando = true;

    try {
        const direccion = new URL(url);

        if (!["http:", "https:"].includes(direccion.protocol)) {
            throw new Error("Solo se permiten direcciones HTTP o HTTPS");
        }

        console.log("Leyendo la pagina con Agent-browser...");

        const comando =
            `agent-browser read "${direccion.href}" --timeout 90000`;

        const respuesta = await ejecutarComando(comando, {
            maxBuffer: 20 * 1024 * 1024,
            timeout: 120000
        });

        const contenidoCompleto = respuesta.stdout.trim();

        if (!contenidoCompleto) {
            throw new Error("Agent-browser no devolvio contenido");
        }

        const lineas = contenidoCompleto
            .split(/\r?\n/)
            .filter(linea => linea.trim() !== "");

        const titulo = lineas.length > 0
            ? lineas[0].replace(/^#+\s*/, "").trim()
            : "";

        const limite = 30000;
        const contenido = contenidoCompleto.slice(0, limite);
        const contenidoRecortado = contenidoCompleto.length > limite;

        console.log("Lectura terminada correctamente");
        console.log("Caracteres obtenidos:", contenidoCompleto.length);

        res.json({
            status: "ok",
            message: "Scraping realizado correctamente",
            url: direccion.href,
            instruction: instruction || "Sin instruccion",
            extraction_mode: "agent-browser-read",
            result: {
                title: titulo,
                content: contenido,
                truncated: contenidoRecortado,
                total_characters: contenidoCompleto.length
            }
        });
    } catch (error) {
        console.error("Error durante el scraping:", error.message);

        const esTiempoAgotado =
            error.killed ||
            error.signal === "SIGTERM" ||
            error.code === "ETIMEDOUT";

        res.status(500).json({
            status: "error",
            message: esTiempoAgotado
                ? "La pagina tardo demasiado en responder"
                : "No fue posible realizar el scraping",
            error: error.message
        });
    } finally {
        procesando = false;
    }
});

app.post("/consultar-salones", async (req, res) => {
    if (procesando) {
        return res.status(429).json({
            status: "busy",
            message: "Ya existe una consulta en proceso. Intente nuevamente."
        });
    }

    let consulta;

    try {
        const mensajeUsuario = String(
            req.body.mensajeUsuario ?? req.body.mensaje_usuario ?? ""
        ).trim();

        // Los modelos locales pequenos pueden inventar horarios al construir
        // los argumentos de la herramienta. Cuando n8n envia el mensaje
        // original, se exige que el usuario haya escrito dos horas o un rango.
        // La respuesta usa HTTP 200 para que el agente pueda convertirla en
        // una pregunta y continuar la misma conversacion.
        const horarioPresenteEnBody = Boolean(
            req.body.horaInicio ?? req.body.hora_inicio
        ) && Boolean(req.body.horaFin ?? req.body.hora_fin);

        if (
            mensajeUsuario &&
            !contieneRangoHorarioExplicito(mensajeUsuario) &&
            !horarioPresenteEnBody
        ) {
            return res.json({
                status: "needs_input",
                message:
                    "Falta el horario. Pregunte al usuario cual es la hora " +
                    "de inicio y la hora de finalizacion. No consulte el portal.",
                respuesta_usuario:
                    "Por favor, indique la hora de inicio y la hora de " +
                    "finalización.",
                missing_fields: ["hora_inicio", "hora_fin"],
                salones: []
            });
        }

        const codigoSalonCrudo = String(
            req.body.codigoSalon ??
            req.body.codigo_salon ??
            req.body.salon ??
            ""
        ).trim();
        const codigoSalon = codigoSalonCrudo
            ? normalizarCodigo(codigoSalonCrudo)
            : "";

        if (
            codigoSalon &&
            !/^[A-Z0-9]+(?:-[A-Z0-9]+)*$/.test(codigoSalon)
        ) {
            throw new Error("codigoSalon no tiene un formato valido");
        }

        const fechasMensaje = extraerFechasDelMensaje(mensajeUsuario);
        if (fechasMensaje.length > 5) {
            throw new Error("Se permiten maximo 5 fechas por consulta");
        }
        const fechas = fechasMensaje.length > 0
            ? fechasMensaje
            : leerFechas(req.body);
        const horarioMensaje = extraerHorarioDelMensaje(mensajeUsuario);
        const horaInicio = normalizarHora(
            horarioMensaje?.inicio ??
                req.body.horaInicio ??
                req.body.hora_inicio,
            "horaInicio"
        );
        const horaFin = normalizarHora(
            horarioMensaje?.fin ??
                req.body.horaFin ??
                req.body.hora_fin,
            "horaFin"
        );
        // En una consulta por codigo exacto no se exige capacidad ni
        // tipologia. Esto evita que el modelo invente condiciones que el
        // usuario no solicito y descarte el salon indicado.
        const capacidadMensaje = extraerCapacidadDelMensaje(mensajeUsuario);
        const capacidadMinima = codigoSalon
            ? 0
            : Number(
                  capacidadMensaje ?? req.body.capacidad ??
                  req.body.capacidadMinima ?? 0
              );
        const caracteristicasMensaje = extraerCaracteristicasDelMensaje(
            mensajeUsuario
        );
        const bodyCaracteristicas = caracteristicasMensaje.length > 0
            ? {
                  ...req.body,
                  caracteristicas: caracteristicasMensaje.join(","),
                  tipologias: undefined,
                  tipoSalon: undefined
              }
            : req.body;
        const caracteristicas = codigoSalon
            ? { obligatorias: [], alternativas: [], desconocidas: [] }
            : leerCaracteristicas(bodyCaracteristicas);

        if (minutosDesdeMedianoche(horaInicio) >= minutosDesdeMedianoche(horaFin)) {
            throw new Error("horaFin debe ser posterior a horaInicio");
        }

        if (
            !codigoSalon &&
            (
                !Number.isFinite(capacidadMinima) ||
                capacidadMinima <= 0 ||
                !Number.isInteger(capacidadMinima)
            )
        ) {
            throw new Error("La capacidad debe ser un numero entero mayor que 0");
        }

        if (caracteristicas.desconocidas.length > 0) {
            throw new Error(
                "Caracteristicas no reconocidas: " +
                caracteristicas.desconocidas.join(", ")
            );
        }

        if (
            !codigoSalon &&
            caracteristicas.obligatorias.length === 0 &&
            caracteristicas.alternativas.length === 0
        ) {
            throw new Error("Caracteristicas: debe indicar al menos una tipologia");
        }

        consulta = {
            fecha: fechas.length === 1 ? fechas[0] : null,
            fechas,
            hora_inicio: horaInicio,
            hora_fin: horaFin,
            codigo_salon: codigoSalon || null,
            capacidad_minima: capacidadMinima,
            caracteristicas_obligatorias: caracteristicas.obligatorias,
            cualquiera_de: caracteristicas.alternativas
        };

        procesando = true;
        console.log("Consultando disponibilidad de salones:", consulta);

        const catalogo = cargarCatalogo();
        const aulaSolicitada = codigoSalon
            ? catalogo.find(aula =>
                normalizarCodigo(aula.codigo_portal) === codigoSalon
            )
            : null;

        if (codigoSalon && !aulaSolicitada) {
            return res.json({
                status: "not_found",
                message: `El salon ${codigoSalon} no aparece en el catalogo`,
                respuesta_usuario:
                    `El salón ${codigoSalon} no aparece en el catálogo.`,
                consulta,
                salones: []
            });
        }

        const esperadas = horasEsperadas(horaInicio, horaFin);
        const resultadosPorFecha = [];

        // Las fechas se consultan una por una para evitar abrir varias
        // sesiones simultaneas contra el portal institucional.
        for (const fecha of fechas) {
            const disponibilidad = await consultarDisponibilidadPortal(
                fecha,
                horaInicio,
                horaFin
            );
            const encontradas = new Set(disponibilidad.horas || []);
            const faltantes = esperadas.filter(hora => !encontradas.has(hora));

            if (faltantes.length > 0) {
                throw new Error(
                    `El portal no devolvio todas las franjas para ${fecha}: ` +
                    faltantes.join(", ")
                );
            }

            const resultado = filtrarCatalogo(
                catalogo,
                disponibilidad.salones || [],
                capacidadMinima,
                caracteristicas.obligatorias,
                caracteristicas.alternativas,
                codigoSalon
            );
            const busquedaAlternativas =
                codigoSalon && resultado.salones.length === 0
                    ? buscarAlternativasSalon(
                          catalogo,
                          disponibilidad.salones || [],
                          aulaSolicitada,
                          codigoSalon
                      )
                    : null;
            const alternativas = busquedaAlternativas?.salones || [];
            const fechaPresentacion =
                `${fecha.slice(0, 4)}-${fecha.slice(4, 6)}-${fecha.slice(6, 8)}`;
            const mensajeFecha = codigoSalon
                ? resultado.salones.length > 0
                    ? `El salon ${codigoSalon} esta disponible el ` +
                      `${fechaPresentacion}, de ${horaInicio} a ${horaFin}`
                    : `El salon ${codigoSalon} no esta disponible el ` +
                      `${fechaPresentacion}, de ${horaInicio} a ${horaFin}. ` +
                      (alternativas.length > 0
                          ? `Se encontraron ${alternativas.length} salones ` +
                            "alternativos con capacidad igual o superior"
                          : "No se encontraron salones alternativos con " +
                            "capacidad igual o superior y mobiliario compatible")
                : resultado.salones.length > 0
                    ? resultado.salones.length === 1
                        ? `Se encontró 1 salón disponible el ${fechaPresentacion}`
                        : `Se encontraron ${resultado.salones.length} salones ` +
                          `disponibles el ${fechaPresentacion}`
                    : `No se encontraron salones que cumplan todas las ` +
                      `condiciones el ${fechaPresentacion}`;

            resultadosPorFecha.push({
                fecha: fechaPresentacion,
                disponible: codigoSalon
                    ? resultado.salones.length > 0
                    : undefined,
                message: mensajeFecha,
                resumen: {
                    franjas_revisadas: esperadas,
                    salones_libres_en_portal: disponibilidad.salones.length,
                    candidatos_segun_catalogo: resultado.candidatosCatalogo,
                    salones_disponibles: resultado.salones.length,
                    alternativas_disponibles: codigoSalon
                        ? alternativas.length
                        : undefined
                },
                salones: resultado.salones,
                alternativas: codigoSalon ? alternativas : undefined,
                criterios_alternativas: busquedaAlternativas
                    ? {
                          capacidad_minima:
                              busquedaAlternativas.capacidad_referencia,
                          caracteristicas:
                              busquedaAlternativas.caracteristicas_referencia
                      }
                    : undefined
            });
        }

        const mensajeResultado = fechas.length === 1
            ? resultadosPorFecha[0].message
            : codigoSalon
                ? `Se revisaron ${fechas.length} fechas para el salon ` +
                  `${codigoSalon}`
                : `Se revisaron ${fechas.length} fechas para encontrar ` +
                  "salones disponibles";
        const resultadosRespuesta = resultadosPorFecha.map(
            compactarResultadoParaAgente
        );

        res.json({
            status: "ok",
            message: mensajeResultado,
            respuesta_usuario: construirRespuestaUsuario(
                resultadosRespuesta,
                codigoSalon
            ),
            consulta,
            resultados_por_fecha: resultadosRespuesta
        });
    } catch (error) {
        console.error("Error consultando salones:", error.message);

        const errorDeEntrada = [
            "fecha",
            "horaInicio",
            "horaFin",
            "capacidad",
            "Caracteristicas",
            "codigoSalon",
            "Debe indicar",
            "Se permiten",
            "La fecha final"
        ].some(texto => error.message.includes(texto));

        if (errorDeEntrada) {
            // Para validaciones se devuelve un unico campo listo para mostrar.
            // Esto evita que modelos locales pequenos mezclen el error con
            // mensajes de disponibilidad o interpreten una lista inexistente.
            return res.json({
                respuesta_usuario: mensajeValidacionParaUsuario(error.message)
            });
        }

        res.status(500).json({
            status: "error",
            message: "No fue posible consultar los salones",
            respuesta_usuario:
                "No fue posible consultar la disponibilidad de salones: " +
                error.message,
            consulta: consulta || null,
            error: error.message
        });
    } finally {
        procesando = false;
    }
});

app.listen(PUERTO, () => {
    console.log(`Servidor ejecutandose en http://localhost:${PUERTO}`);
    console.log(`Catalogo de aulas: ${ARCHIVO_CATALOGO}`);
});
