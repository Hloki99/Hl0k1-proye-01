// js/logica.js

// ==========================================================================
// VARIABLES GLOBALES (Exactamente como las tienes en tu diseño)
// ==========================================================================
let registrosLeidosDelExcel = [];
let datosExcelA = null;
let datosExcelB = null;
let baseDatosLocal;
let modoEliminacionActivo = false;

// ==========================================================================
// FUNCIÓN DE NOTIFICACIONES (Apunta a tu div "zona-alerta")
// ==========================================================================
function mostrarNotificacion(mensaje, tipo = "info") {
    const zonaAlerta = document.getElementById('zona-alerta');
    if (zonaAlerta) {
        const alerta = document.createElement('div');
        alerta.className = `p-3 rounded-xl text-xs font-bold text-white shadow-md mb-4 text-center transition-all transform duration-300 ${
            tipo === 'success' ? 'bg-emerald-600' : tipo === 'error' ? 'bg-rose-600' : 'bg-indigo-600'
        }`;
        alerta.innerText = mensaje;
        zonaAlerta.appendChild(alerta);
        
        setTimeout(() => {
            alerta.remove();
        }, 3500);
    } else {
        alert(mensaje);
    }
}

// ==========================================================================
// INTERFAZ, SWITCHES Y CONTROL DE DRAWERS
// ==========================================================================
function alternarVisibilidadBorrado() {
    modoEliminacionActivo = document.getElementById('switch-modo-borrar').checked;
    const botonesBorrar = document.querySelectorAll('.btn-borrar-dinamico');
    botonesBorrar.forEach(boton => {
        if (modoEliminacionActivo) {
            boton.classList.remove('hidden');
        } else {
            boton.classList.add('hidden');
        }
    });
}

function toggleDrawer(idDrawer) {
    const panel = document.getElementById(idDrawer);
    const overlay = document.getElementById('drawer-overlay');
    const estaCerrado = panel.classList.contains('translate-x-full') || panel.classList.contains('-translate-x-full');
    
    cerrarTodosLosDrawers();

    if (estaCerrado) {
        if (idDrawer === 'drawer-importar') {
            panel.classList.remove('translate-x-full');
        } else {
            panel.classList.remove('-translate-x-full');
        }
        overlay.classList.remove('hidden');
        setTimeout(() => overlay.classList.add('opacity-100'), 10);
    }
}

function cerrarTodosLosDrawers() {
    document.getElementById('drawer-importar').classList.add('translate-x-full');
    document.getElementById('drawer-comparar').classList.add('-translate-x-full');
    document.getElementById('drawer-usuarios').classList.add('-translate-x-full');
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
        btnForm.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold";
        btnTrash.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer";
        conForm.classList.remove('hidden');
        conTrash.classList.add('hidden');
    } else {
        btnTrash.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold";
        btnForm.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer";
        conTrash.classList.remove('hidden');
        conForm.classList.add('hidden');
        renderizarPapelera();
    }
}

// ==========================================================================
// INITIALIZACIÓN DE INDEXEDDB
// ==========================================================================
const peticionDB = indexedDB.open("ProcesadorExcelMasivoDB", 2);

peticionDB.onupgradeneeded = function(e) {
    let db = e.target.result;
    if (!db.objectStoreNames.contains("registros")) db.createObjectStore("registros", { keyPath: "cedula" });
    if (!db.objectStoreNames.contains("registros_eliminados")) db.createObjectStore("registros_eliminados", { keyPath: "cedula" });
};

peticionDB.onsuccess = function(e) {
    baseDatosLocal = e.target.result;
    renderizarTablaDesdeDB();
    const inputFecha = document.getElementById('manual-fecha');
    if (inputFecha) inputFecha.value = new Date().toISOString().split('T')[0];
};

// ==========================================================================
// CONTROLADORES DE ARCHIVOS EXCEL (Conectados por ID al DOM)
// ==========================================================================
document.addEventListener('DOMContentLoaded', () => {
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
// OPERACIONES PRINCIPALES: EXCEL Y BASE DE DATOS
// ==========================================================================
function procesarYGuardarTodo() {
    if (registrosLeidosDelExcel.length === 0) return alert("Hermano, primero selecciona un archivo Excel Principal.");
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
        mostrarNotificacion(`Se cargaron ${guardados} registros en bloque.`, "success");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

function ejecutarComparacionIndependiente() {
    if (!datosExcelA || !datosExcelB) return alert("Hermano, debes cargar ambos archivos.");
    const contenedor = document.getElementById('resultados-comparacion');
    contenedor.innerHTML = `<p class="animate-pulse text-purple-600 font-bold">Cruzando en memoria...</p>`;

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
            err++; html += `<div class="p-2 bg-amber-50 text-amber-800 rounded-lg">⚠️ Fila B-${fila}: Cédula ${c} no existe en A.</div>`;
            return;
        }
        const dA = mapaA.get(c); let dif = [];
        if (dA.nombre !== String(f.Nombre||"").trim()) dif.push("Nombre");
        if (dA.status !== String(f.Status||"").trim()) dif.push("Status");
        if (dA.serial !== String(f.Serial||"").trim()) dif.push("Serial");
        if (dif.length > 0) {
            err++; html += `<div class="p-2 bg-rose-50 text-rose-700 rounded-lg">❌ Fila B-${fila} (${c}): Distinto en ${dif.join(', ')}</div>`;
        }
    });
    contenedor.innerHTML = err === 0 ? `<div class="p-3 bg-green-100 text-green-800 font-bold rounded-lg text-center">✔️ Coincidencia perfecta.</div>` : html;
}

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
            if (datosExport.length === 0) return alert("Base de datos vacía.");
            const hoja = XLSX.utils.json_to_sheet(datosExport);
            const libro = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(libro, hoja, "Facturación");
            XLSX.writeFile(libro, "BaseDatos_Facturas.xlsx");
            cerrarTodosLosDrawers();
        }
    };
}

function limpiarTodaLaDB() {
    if(!confirm("Hermano, ¿seguro que quieres vaciar la tabla por completo?")) return;
    baseDatosLocal.transaction(["registros"], "readwrite").objectStore("registros").clear().onsuccess = function() {
        mostrarNotificacion("Almacenamiento vaciado.", "success");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

// ==========================================================================
// GESTIÓN MANUAL Y RENDERIZADO (Originales del script)
// ==========================================================================
function guardarPersonaManual() {
    const cedula = document.getElementById('manual-cedula').value.trim();
    const nombre = document.getElementById('manual-nombre').value.trim();
    const status = document.getElementById('manual-status').value.trim();
    const serial = document.getElementById('manual-serial').value.trim();
    const statusFactura = document.getElementById('manual-stf').value.trim();
    const fecha = document.getElementById('manual-fecha').value;
    const pago = document.getElementById('manual-pago').checked;

    if (!cedula || !nombre) return alert("Hermano, Cédula y Nombre son obligatorios.");

    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    transaccion.objectStore("registros").put({ cedula, nombre, status, serial, statusFactura, fecha, pago });

    transaccion.oncomplete = function() {
        mostrarNotificacion(`Persona [${cedula}] guardada.`, "success");
        renderizarTablaDesdeDB();
        document.getElementById('manual-cedula').value = "";
        document.getElementById('manual-nombre').value = "";
        cerrarTodosLosDrawers();
    };
}

function moverAPapelera(cedulaKey) {
    const txLectura = baseDatosLocal.transaction(["registros"], "readonly");
    txLectura.objectStore("registros").get(cedulaKey).onsuccess = function(e) {
        const registroOriginal = e.target.result;
        if (!registroOriginal) return;

        const txEscritura = baseDatosLocal.transaction(["registros", "registros_eliminados"], "readwrite");
        txEscritura.objectStore("registros_eliminados").put(registroOriginal);
        txEscritura.objectStore("registros").delete(cedulaKey);

        txEscritura.oncomplete = function() {
            mostrarNotificacion(`Registro ${cedulaKey} movido a la papelera.`, "error");
            renderizarTablaDesdeDB();
        };
    };
}

function restaurarDesdePapelera(cedulaKey) {
    const txLectura = baseDatosLocal.transaction(["registros_eliminados"], "readonly");
    txLectura.objectStore("registros_eliminados").get(cedulaKey).onsuccess = function(e) {
        const registro = e.target.result;
        if (!registro) return;

        const txRestaurar = baseDatosLocal.transaction(["registros", "registros_eliminados"], "readwrite");
        txRestaurar.objectStore("registros").put(registro);
        txRestaurar.objectStore("registros_eliminados").delete(cedulaKey);

        txRestaurar.oncomplete = function() {
            mostrarNotificacion(`Registro ${cedulaKey} restaurado.`, "success");
            renderizarTablaDesdeDB();
            renderizarPapelera();
        };
    };
}

function renderizarPapelera() {
    const contenedor = document.getElementById('lista-papelera');
    contenedor.innerHTML = "";
    const tx = baseDatosLocal.transaction(["registros_eliminados"], "readonly");
    let listadoHtml = ""; let totalBorrados = 0;

    tx.objectStore("registros_eliminados").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            totalBorrados++; const item = cursor.value;
            listadoHtml += `
                <div class="p-2 bg-slate-50 border border-slate-200 rounded-lg flex justify-between items-center gap-2">
                    <div class="truncate flex-1">
                        <p class="font-bold text-slate-800 text-xs">V-${item.cedula}</p>
                        <p class="text-gray-500 text-[11px] truncate">${item.nombre}</p>
                    </div>
                    <button onclick="restaurarDesdePapelera('${item.cedula}')" class="bg-emerald-600 hover:bg-emerald-700 text-white text-[10px] font-bold px-2 py-1 rounded cursor-pointer">Restaurar</button>
                </div>`;
            cursor.continue();
        } else {
            contenedor.innerHTML = totalBorrados === 0 ? `<p class="italic text-gray-400 text-center py-4">La papelera está vacía, hermano.</p>` : listadoHtml;
        }
    };
}

function vaciarPapeleraPorCompleto() {
    if (!confirm("Hermano, ¿seguro que quieres purgar la papelera definitivamente?")) return;
    baseDatosLocal.transaction(["registros_eliminados"], "readwrite").objectStore("registros_eliminados").clear().onsuccess = function() {
        mostrarNotificacion("Papelera purgada.", "success");
        renderizarPapelera();
    };
}

function renderizarTablaDesdeDB() {
    const tbody = document.getElementById('tabla-registros');
    tbody.innerHTML = "";
    const transaccion = baseDatosLocal.transaction(["registros"], "readonly");
    let total = 0;

    transaccion.objectStore("registros").openCursor().onsuccess = function(e) {
        const cursor = e.target.result;
        if (cursor) {
            total++; const item = cursor.value;
            const tr = document.createElement('tr');
            tr.className = "hover:bg-slate-50 transition text-gray-700";
            
            const claseOcultarBorrado = modoEliminacionActivo ? "" : "hidden";

            tr.innerHTML = `
                <td class="p-3 font-bold text-indigo-950 bg-indigo-50/40">${item.cedula}</td>
                <td class="p-2"><input type="text" id="nom-${item.cedula}" value="${item.nombre}" class="w-full p-1 border border-gray-200 rounded text-xs bg-transparent focus:bg-white focus:border-indigo-500"></td>
                <td class="p-2"><input type="text" id="sta-${item.cedula}" value="${item.status}" class="w-full p-1 border border-gray-200 rounded text-xs bg-transparent focus:bg-white focus:border-indigo-500"></td>
                <td class="p-2"><input type="text" id="ser-${item.cedula}" value="${item.serial}" class="w-full p-1 border border-gray-200 rounded text-xs bg-transparent focus:bg-white focus:border-indigo-500"></td>
                <td class="p-2"><input type="text" id="stf-${item.cedula}" value="${item.statusFactura}" class="w-full p-1 border border-gray-200 rounded text-xs bg-transparent focus:bg-white focus:border-indigo-500"></td>
                <td class="p-2"><input type="text" id="fec-${item.cedula}" value="${item.fecha}" class="w-full p-1 border border-gray-200 rounded text-xs bg-transparent focus:bg-white focus:border-indigo-500"></td>
                <td class="p-2 text-center"><input type="checkbox" id="pag-${item.cedula}" ${item.pago ? 'checked' : ''} class="w-5 h-5 text-indigo-600 border-gray-300 rounded cursor-pointer mt-1"></td>
                <td class="p-2 text-center flex items-center justify-center gap-1 min-w-[70px]">
                    <button onclick="guardarEdicionFila('${item.cedula}')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-xs font-semibold transition shadow cursor-pointer" title="Guardar cambios">💾</button>
                    <button onclick="moverAPapelera('${item.cedula}')" class="btn-borrar-dinamico ${claseOcultarBorrado} bg-rose-600 hover:bg-rose-700 text-white px-2 py-1 rounded text-xs font-semibold transition shadow cursor-pointer animate-fade-in" title="Eliminar y enviar a papelera">🗑️</button>
                </td>
            `;
            tbody.appendChild(tr);
            cursor.continue();
        } else if (total === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 italic">No hay registros cargados.</td></tr>`;
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
    transaccion.oncomplete = function() {
        mostrarNotificacion(`Cambios guardados en Cédula ${cedulaKey}`, "success");
    };
}
