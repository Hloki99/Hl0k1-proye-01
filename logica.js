// js/logica.js

// ==========================================================================
// 1. CONFIGURACIÓN Y APERTURA DE INDEXEDDB
// ==========================================================================
let baseDatosLocal;
const peticionDB = indexedDB.open("ProcesadorExcelMasivoDB", 2);

peticionDB.onupgradeneeded = function(e) {
    let db = e.target.result;
    if (!db.objectStoreNames.contains("registros")) {
        db.createObjectStore("registros", { keyPath: "cedula" });
    }
    if (!db.objectStoreNames.contains("registros_eliminados")) {
        db.createObjectStore("registros_eliminados", { keyPath: "cedula" });
    }
};

peticionDB.onsuccess = function(e) {
    baseDatosLocal = e.target.result;
    console.log("IndexedDB inicializada y conectada con éxito.");
    
    // Renderiza la tabla inmediatamente al conectar la BD
    renderizarTablaDesdeDB();
    
    // Coloca la fecha de hoy por defecto en el formulario manual si existe
    const campoFecha = document.getElementById('manual-fecha');
    if (campoFecha) campoFecha.value = new Date().toISOString().split('T')[0];
};

peticionDB.onerror = function(e) {
    console.error("Error crítico al abrir IndexedDB:", e.target.error);
};

function limpiarTodaLaDB() {
    if(!confirm("¿Seguro que quieres vaciar la tabla por completo?")) return;
    baseDatosLocal.transaction(["registros"], "readwrite").objectStore("registros").clear().onsuccess = function() {
        alert("Almacenamiento local vaciado con éxito.");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

// ==========================================================================
// 2. CONTROLADORES DE EVENTOS (EXCEL E INTERFAZ)
// ==========================================================================
let registrosLeidosDelExcel = [];
let datosExcelA = null;
let datosExcelB = null;
let modoEliminacionActivo = false;

document.addEventListener('DOMContentLoaded', () => {
    // Escuchador para la carga del Excel Principal
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
                        <p class="text-emerald-600 font-bold">📄 ${archivo.name}</p>
                        <p class="text-slate-700 text-xs">📊 Filas detectadas: ${registrosLeidosDelExcel.length}</p>
                    `;
                }
            };
            lector.readAsArrayBuffer(archivo);
        });
    }

    // Escuchadores para la herramienta de comparación (Excel A y B)
    const compA = document.getElementById('excel-comp-a');
    if (compA) {
        compA.addEventListener('change', function(e) {
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

    const compB = document.getElementById('excel-comp-b');
    if (compB) {
        compB.addEventListener('change', function(e) {
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

// ==========================================================================
// 3. LÓGICA DE PROCESAMIENTO Y CRUCE DE DATOS
// ==========================================================================
function procesarYGuardarTodo() {
    if (registrosLeidosDelExcel.length === 0) return alert("Primero selecciona un archivo Excel Principal.");
    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    const almacen = transaccion.objectStore("registros");
    let guardados = 0;

    registrosLeidosDelExcel.forEach(celda => {
        const llaves = Object.keys(celda);
        let cedula = String(celda.Cedula || celda.Cédula || celda.CEDULA || celda.id || celda[llaves[0]] || "").trim();
        let nombre = String(celda.Nombre || celda.NOMBRE || celda.nombre || celda[llaves[1]] || "").trim();
        let status = String(celda.Status || celda.STATUS || celda.Estado || celda[llaves[2]] || "").trim();
        let serial = String(celda.Serial || celda.Seriar || celda.SERIAL || celda[llaves[3]] || "").trim();
        let statusFactura = String(celda['Status Factura'] || celda[llaves[4]] || "").trim();
        let fecha = String(celda.Fecha || celda[llaves[5]] || "").trim();
        let pagoOriginal = celda.Pago || celda[llaves[6]] || 0;

        if (!cedula || cedula === "undefined") return;
        let pagoElegido = (pagoOriginal == "SÍ" || pagoOriginal == "SI" || pagoOriginal == 1 || pagoOriginal == true);

        almacen.put({ cedula, nombre, status, serial, statusFactura, fecha, pago: pagoElegido });
        guardados++;
    });

    transaccion.oncomplete = function() {
        alert(`Éxito: Se procesaron y guardaron ${guardados} registros.`);
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

function ejecutarComparacionIndependiente() {
    if (!datosExcelA || !datosExcelB) return alert("Debes cargar ambos archivos Excel para poder cruzarlos.");
    const contenedor = document.getElementById('resultados-comparacion');
    contenedor.innerHTML = `<p class="animate-pulse text-indigo-600 font-bold text-xs">Cruzando registros en memoria...</p>`;

    const mapaA = new Map();
    datosExcelA.forEach(f => {
        const k = Object.keys(f);
        let c = String(f.Cedula || f.Cédula || f.CEDULA || f[k[0]] || "").trim();
        if (c) mapaA.set(c, { nombre: String(f.Nombre||""), status: String(f.Status||""), serial: String(f.Serial||"") });
    });

    let html = ""; let err = 0;
    datosExcelB.forEach((f, i) => {
        const k = Object.keys(f); const fila = i + 2;
        let c = String(f.Cedula || f[k[0]] || "").trim();
        if (!c) return;
        if (!mapaA.has(c)) {
            err++; html += `<div class="p-2 bg-amber-50 text-amber-800 rounded-lg text-xs border border-amber-100">⚠️ Fila B-${fila}: Cédula ${c} no existe en el Excel A.</div>`;
            return;
        }
        const dA = mapaA.get(c); let dif = [];
        if (dA.nombre !== String(f.Nombre||"").trim()) dif.push("Nombre");
        if (dA.status !== String(f.Status||"").trim()) dif.push("Status");
        if (dA.serial !== String(f.Serial||"").trim()) dif.push("Serial");
        if (dif.length > 0) {
            err++; html += `<div class="p-2 bg-rose-50 text-rose-700 rounded-lg text-xs border border-rose-100">❌ Fila B-${fila} (${c}): Discrepancia en ${dif.join(', ')}</div>`;
        }
    });
    contenedor.innerHTML = err === 0 ? `<div class="p-3 bg-emerald-100 text-emerald-800 font-bold rounded-lg text-center text-xs">✔️ Coincidencia al 100%: Los archivos son idénticos.</div>` : html;
}

function exportarBaseDatos() {
    if (!baseDatosLocal) return alert("La base de datos aún no está lista.");
    const transaccion = baseDatosLocal.transaction(["registros"], "readonly");
    let datosExport = [];
    transaccion.objectStore("registros").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            const r = cursor.value;
            datosExport.push({ "Cédula": r.cedula, "Nombre": r.nombre, "Status": r.status, "Serial": r.serial, "Status Factura": r.statusFactura, "Fecha": r.fecha, "Pago": r.pago ? "SÍ" : "NO" });
            cursor.continue();
        } else {
            if (datosExport.length === 0) return alert("No hay datos en la tabla para poder exportar.");
            const hoja = XLSX.utils.json_to_sheet(datosExport);
            const libro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(libro, hoja, "Facturación");
            XLSX.writeFile(libro, "BaseDatos_Facturas.xlsx");
            cerrarTodosLosDrawers();
        }
    };
}

// ==========================================================================
// 4. GESTIÓN VISUAL DEL PANEL (INTERFACE DRAWERS)
// ==========================================================================
function alternarVisibilidadBorrado() {
    modoEliminacionActivo = document.getElementById('switch-modo-borrar').checked;
    const botonesBorrar = document.querySelectorAll('.btn-borrar-dinamico');
    botonesBorrar.forEach(boton => {
        if (modoEliminacionActivo) boton.classList.remove('hidden');
        else boton.classList.add('hidden');
    });
}

function toggleDrawer(idDrawer) {
    const panel = document.getElementById(idDrawer);
    const overlay = document.getElementById('drawer-overlay');
    const estaCerrado = panel.classList.contains('translate-x-full') || panel.classList.contains('-translate-x-full');
    
    cerrarTodosLosDrawers();

    if (estaCerrado) {
        if (idDrawer === 'drawer-importar') panel.classList.remove('translate-x-full');
        else panel.classList.remove('-translate-x-full');
        if (overlay) {
            overlay.classList.remove('hidden');
            setTimeout(() => overlay.classList.add('opacity-100'), 10);
        }
    }
}

function cerrarTodosLosDrawers() {
    const dImp = document.getElementById('drawer-importar');
    const dComp = document.getElementById('drawer-comparar');
    const dUsu = document.getElementById('drawer-usuarios');
    if (dImp) dImp.classList.add('translate-x-full');
    if (dComp) dComp.classList.add('-translate-x-full');
    if (dUsu) dUsu.classList.add('-translate-x-full');
    
    const overlay = document.getElementById('drawer-overlay');
    if (overlay) {
        overlay.classList.remove('opacity-100');
        setTimeout(() => overlay.classList.add('hidden'), 30);
    }
}

function cambiarPestanaInterna(tipo) {
    const btnForm = document.getElementById('tab-btn-form');
    const btnTrash = document.getElementById('tab-btn-trash');
    const conForm = document.getElementById('contenido-tab-form');
    const conTrash = document.getElementById('contenido-tab-trash');

    if (tipo === 'form') {
        if(btnForm) btnForm.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold text-xs";
        if(btnTrash) btnTrash.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer text-xs";
        if(conForm) conForm.classList.remove('hidden');
        if(conTrash) conTrash.classList.add('hidden');
    } else {
        if(btnTrash) btnTrash.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold text-xs";
        if(btnForm) btnForm.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer text-xs";
        if(conTrash) conTrash.classList.remove('hidden');
        if(conForm) conForm.classList.add('hidden');
        renderizarPapelera();
    }
}

// ==========================================================================
// 5. RENDERIZADO DINÁMICO DE TABLAS Y FILAS
// ==========================================================================
function renderizarTablaDesdeDB() {
    const tbody = document.getElementById('tabla-registros');
    if (!tbody || !baseDatosLocal) return;
    tbody.innerHTML = "";
    const transaccion = baseDatosLocal.transaction(["registros"], "readonly");
    let total = 0;

    transaccion.objectStore("registros").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            total++; const item = cursor.value;
            const tr = document.createElement('tr');
            tr.className = "hover:bg-slate-50 transition text-gray-700 text-xs border-b";
            const claseOcultarBorrado = modoEliminacionActivo ? "" : "hidden";

            tr.innerHTML = `
                <td class="p-3 font-bold text-indigo-950 bg-indigo-50/40">${item.cedula}</td>
                <td class="p-2"><input type="text" id="nom-${item.cedula}" value="${item.nombre}" class="w-full p-1 border rounded text-xs bg-white"></td>
                <td class="p-2"><input type="text" id="sta-${item.cedula}" value="${item.status}" class="w-full p-1 border rounded text-xs bg-white"></td>
                <td class="p-2"><input type="text" id="ser-${item.cedula}" value="${item.serial}" class="w-full p-1 border rounded text-xs bg-white"></td>
                <td class="p-2"><input type="text" id="stf-${item.cedula}" value="${item.statusFactura}" class="w-full p-1 border rounded text-xs bg-white"></td>
                <td class="p-2"><input type="text" id="fec-${item.cedula}" value="${item.fecha}" class="w-full p-1 border rounded text-xs bg-white"></td>
                <td class="p-2 text-center"><input type="checkbox" id="pag-${item.cedula}" ${item.pago ? 'checked' : ''} class="w-4 h-4 cursor-pointer"></td>
                <td class="p-2 text-center flex items-center justify-center gap-1">
                    <button onclick="guardarEdicionFila('${item.cedula}')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2 py-1 rounded text-[11px] cursor-pointer" title="Guardar Cambios">💾</button>
                    <button onclick="moverAPapelera('${item.cedula}')" class="btn-borrar-dinamico ${claseOcultarBorrado} bg-rose-600 hover:bg-rose-700 text-white px-2 py-1 rounded text-[11px] cursor-pointer" title="Mover a Papelera">🗑️</button>
                </td>
            `;
            tbody.appendChild(tr);
            cursor.continue();
        } else if (total === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 italic">No hay registros cargados actualmente.</td></tr>`;
        }
    };
}

function guardarEdicionFila(cedulaKey) {
    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    transaccion.objectStore("registros").put({
        cedula: cedulaKey,
        nombre: document.getElementById(`nom-${cedulaKey}`).value,
        status: document.getElementById(`sta-${cedulaKey}`).value,
        serial: document.getElementById(`ser-${cedulaKey}`).value,
        statusFactura: document.getElementById(`stf-${cedulaKey}`).value,
        fecha: document.getElementById(`fec-${cedulaKey}`).value,
        pago: document.getElementById(`pag-${cedulaKey}`).checked
    });
    transaccion.oncomplete = function() { console.log(`Cédula ${cedulaKey} actualizada.`); };
}

function guardarPersonaManual() {
    const cedula = document.getElementById('manual-cedula').value.trim();
    const nombre = document.getElementById('manual-nombre').value.trim();
    const status = document.getElementById('manual-status').value.trim();
    const serial = document.getElementById('manual-serial').value.trim();
    const statusFactura = document.getElementById('manual-stf').value.trim();
    const fecha = document.getElementById('manual-fecha').value;
    const pago = document.getElementById('manual-pago').checked;

    if (!cedula || !nombre) return alert("Los campos Cédula y Nombre son completamente obligatorios.");
    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    transaccion.objectStore("registros").put({ cedula, nombre, status, serial, statusFactura, fecha, pago });
    transaccion.oncomplete = function() {
        alert(`Guardado con éxito.`);
        renderizarTablaDesdeDB();
        document.getElementById('manual-cedula').value = "";
        document.getElementById('manual-nombre').value = "";
        cerrarTodosLosDrawers();
    };
}

function moverAPapelera(cedulaKey) {
    const txLectura = baseDatosLocal.transaction(["registros"], "readonly");
    txLectura.objectStore("registros").get(cedulaKey).onsuccess = function(e) {
        const reg = e.target.result; if (!reg) return;
        const txEscritura = baseDatosLocal.transaction(["registros", "registros_eliminados"], "readwrite");
        txEscritura.objectStore("registros_eliminados").put(reg);
        txEscritura.objectStore("registros").delete(cedulaKey);
        txEscritura.oncomplete = function() { renderizarTablaDesdeDB(); };
    };
}

function restaurarDesdePapelera(cedulaKey) {
    const txLectura = baseDatosLocal.transaction(["registros_eliminados"], "readonly");
    txLectura.objectStore("registros_eliminados").get(cedulaKey).onsuccess = function(e) {
        const reg = e.target.result; if (!reg) return;
        const txRestaurar = baseDatosLocal.transaction(["registros", "registros_eliminados"], "readwrite");
        txRestaurar.objectStore("registros").put(reg);
        txRestaurar.objectStore("registros_eliminados").delete(cedulaKey);
        txRestaurar.oncomplete = function() { renderizarTablaDesdeDB(); renderizarPapelera(); };
    };
}

function renderizarPapelera() {
    const contenedor = document.getElementById('lista-papelera'); if(!contenedor) return;
    contenedor.innerHTML = "";
    const tx = baseDatosLocal.transaction(["registros_eliminados"], "readonly");
    let h = ""; let t = 0;
    tx.objectStore("registros_eliminados").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            t++; const item = cursor.value;
            h += `<div class="p-2 bg-slate-50 border rounded-lg flex justify-between items-center text-xs mb-2"><div><p class="font-bold">V-${item.cedula}</p><p class="text-gray-500">${item.nombre}</p></div><button onclick="restaurarDesdePapelera('${item.cedula}')" class="bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] px-2 py-1 rounded cursor-pointer">Restaurar</button></div>`;
            cursor.continue();
        } else { contenedor.innerHTML = t === 0 ? `<p class="italic text-center py-4 text-gray-400 text-xs">Papelera vacía.</p>` : h; }
    };
}

function vaciarPapeleraPorCompleto() {
    if (!confirm("¿Deseas vaciar la papelera permanentemente?")) return;
    baseDatosLocal.transaction(["registros_eliminados"], "readwrite").objectStore("registros_eliminados").clear().onsuccess = function() {
        renderizarPapelera();
    };
}