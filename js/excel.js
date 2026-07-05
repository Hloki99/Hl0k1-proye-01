// js/excel.js

// Estados temporales para los archivos leídos
var registrosLeidosDelExcel = [];
var datosExcelA = null;
var datosExcelB = null;

// Inicializar detectores de eventos cuando el DOM esté listo
document.addEventListener("DOMContentLoaded", function() {
    
    // 1. Lector Excel Principal
    const inputExcel = document.getElementById('input-excel');
    if (inputExcel) {
        inputExcel.addEventListener('change', function(e) {
            const archivo = e.target.files[0];
            if (!archivo) return;
            const lector = new FileReader();
            lector.onload = function(evento) {
                const datos = new Uint8Array(evento.target.result);
                const libro = XLSX.read(datos, { type: 'array' });
                registrosLeidosDelExcel = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]]);
                if(registrosLeidosDelExcel.length > 0) {
                    document.getElementById('info-archivo').innerHTML = `
                        <p class="text-green-600 font-bold">📄 ${archivo.name}</p>
                        <p class="text-gray-700">📊 Filas leídas: ${registrosLeidosDelExcel.length}</p>
                    `;
                    mostrarNotificacion("Archivo cargado al lector temporal.", "success");
                }
            };
            lector.readAsArrayBuffer(archivo);
        });
    }

    // 2. Lector para Comparador Independiente - Archivo A
    const excelCompA = document.getElementById('excel-comp-a');
    if (excelCompA) {
        excelCompA.addEventListener('change', function(e) {
            const archivo = e.target.files[0];
            if (!archivo) return;
            document.getElementById('label-comp-a').innerText = archivo.name;
            const lector = new FileReader();
            lector.onload = function(ev) {
                const libro = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
                datosExcelA = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]], { raw: false });
            };
            lector.readAsArrayBuffer(archivo);
        });
    }

    // 3. Lector para Comparador Independiente - Archivo B
    const excelCompB = document.getElementById('excel-comp-b');
    if (excelCompB) {
        excelCompB.addEventListener('change', function(e) {
            const archivo = e.target.files[0];
            if (!archivo) return;
            document.getElementById('label-comp-b').innerText = archivo.name;
            const lector = new FileReader();
            lector.onload = function(ev) {
                const libro = XLSX.read(new Uint8Array(ev.target.result), { type: 'array' });
                datosExcelB = XLSX.utils.sheet_to_json(libro.Sheets[libro.SheetNames[0]], { raw: false });
            };
            lector.readAsArrayBuffer(archivo);
        });
    }
});

// Procesar bloque masivo a IndexedDB
function procesarYGuardarTodo() {
    if (registrosLeidosDelExcel.length === 0) return alert("Pana, primero selecciona un archivo Excel Principal.");
    if (!baseDatosLocal) return alert("La base de datos local no está lista todavía.");

    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    const almacen = transaccion.objectStore("registros");
    let guardados = 0;

    registrosLeidosDelExcel.forEach(celda => {
        const llaves = Object.keys(celda);
        let cedula = String(celda.Cedula || celda.Cédula || celda.CEDULA || celda.id || celda[llaves[0]] || "").trim();
        let nombre = String(celda.Nombre || celda.NOMBRE || celda.nombre || celda[llaves[1]] || "").trim();
        let status = String(celda.Status || celda.STATUS || celda.Estado || celda[llaves[2]] || "").trim();
        let serial = String(celda.Serial || celda.Seriar || celda.SERIAL || celda[llaves[3]] || "").trim();
        let statusFactura = String(celda['Status Factura'] || celda['Status de la factura'] || celda[llaves[4]] || "").trim();
        let fecha = String(celda.Fecha || celda.FECHA || celda[llaves[5]] || "").trim();
        let pagoOriginal = celda.Pago || celda.PAGO || celda[llaves[6]] || 0;

        if (!cedula || cedula === "undefined") return;
        let pagoElegido = (pagoOriginal == "SÍ" || pagoOriginal == "SI" || pagoOriginal == 1 || pagoOriginal == true);

        almacen.put({ cedula, nombre, status, serial, statusFactura, fecha, pago: pagoElegido });
        guardados++;
    });

    transaccion.oncomplete = function() {
        mostrarNotificacion(`Se cargaron ${guardados} registros en bloque correctamente.`, "success");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

// Comparación en memoria volatíl (En el aire)
function ejecutarComparacionIndependiente() {
    if (!datosExcelA || !datosExcelB) return alert("Pana, debes cargar ambos archivos para poder cruzarlos.");
    const contenedor = document.getElementById('resultados-comparacion');
    contenedor.innerHTML = `<p class="animate-pulse text-purple-600 font-bold">Cruzando datos en memoria...</p>`;

    const mapaA = new Map();
    datosExcelA.forEach(f => {
        const k = Object.keys(f);
        let c = String(f.Cedula || f.Cédula || f.CEDULA || f[k[0]] || "").trim();
        if (c) mapaA.set(c, { nombre: String(f.Nombre||f[k[1]]||""), status: String(f.Status||f[k[2]]||""), serial: String(f.Serial||f[k[3]]||"") });
    });

    let html = ""; let err = 0;
    datosExcelB.forEach((f, i) => {
        const k = Object.keys(f); const fila = i + 2;
        let c = String(f.Cedula || f.Cédula || f.CEDULA || f[k[0]] || "").trim();
        if (!c) return;
        if (!mapaA.has(c)) {
            err++; html += `<div class="p-2 bg-amber-50 border border-amber-200 text-amber-800 rounded-lg">⚠️ Fila B-${fila}: Cédula ${c} no existe en Archivo A.</div>`;
            return;
        }
        const dA = mapaA.get(c); let dif = [];
        if (dA.nombre !== String(f.Nombre||f[k[1]]||"").trim()) dif.push("Nombre");
        if (dA.status !== String(f.Status||f[k[2]]||"").trim()) dif.push("Status");
        if (dA.serial !== String(f.Serial||f[k[3]]||"").trim()) dif.push("Serial");
        if (dif.length > 0) {
            err++; html += `<div class="p-2 bg-rose-50 border border-rose-200 text-rose-700 rounded-lg">❌ Fila B-${fila} (${c}): Distinto en [${dif.join(', ')}]</div>`;
        }
    });
    contenedor.innerHTML = err === 0 ? `<div class="p-3 bg-green-100 text-green-800 font-bold rounded-lg text-center">✔️ Coincidencia perfecta en el aire.</div>` : `<div class="space-y-2 max-h-96 overflow-y-auto">${html}</div>`;
}

// Exportar Base de datos local hacia Excel físico
function exportarBaseDatos() {
    if (!baseDatosLocal) return alert("Base de datos no inicializada.");
    const transaccion = baseDatosLocal.transaction(["registros"], "readonly");
    let datosExport = [];
    transaccion.objectStore("registros").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            const r = cursor.value;
            datosExport.push({ "Cédula": r.cedula, "Nombre": r.nombre, "Status": r.status, "Serial": r.serial, "Status Factura": r.statusFactura, "Fecha": r.fecha, "Pago": r.pago ? "SÍ" : "NO" });
            cursor.continue();
        } else {
            if (datosExport.length === 0) return alert("La base de datos está vacía.");
            const hoja = XLSX.utils.json_to_sheet(datosExport);
            const libro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(libro, hoja, "Facturación");
            XLSX.writeFile(libro, "BaseDatos_Facturas.xlsx");
            cerrarTodosLosDrawers();
        }
    };
}