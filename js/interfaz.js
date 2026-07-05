// js/interfaz.js

var modoEliminacionActivo = false;

// 1. ALTERNAR MODO ELIMINACIÓN
function alternarVisibilidadBorrado() {
    const switchElement = document.getElementById('switch-modo-borrar');
    modoEliminacionActivo = switchElement ? switchElement.checked : false;
    
    const botonesBorrar = document.querySelectorAll('.btn-borrar-dinamico');
    botonesBorrar.forEach(boton => {
        if (modoEliminacionActivo) {
            boton.classList.remove('hidden');
        } else {
            boton.classList.add('hidden');
        }
    });
}

// 2. GESTIÓN DE DRAWERS DESPLEGABLES
function toggleDrawer(idDrawer) {
    const panel = document.getElementById(idDrawer);
    const overlay = document.getElementById('drawer-overlay');
    if (!panel || !overlay) return;

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
    overlay.classList.remove('opacity-100');
    setTimeout(() => overlay.classList.add('hidden'), 30);
}

// Pestañas del submenú personal/papelera
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

// 3. RENDERIZAR TABLA PRINCIPAL DESDE INDEXEDDB
function renderizarTablaDesdeDB() {
    const tbody = document.getElementById('tabla-registros');
    if (!tbody) return;
    tbody.innerHTML = "";

    if (!baseDatosLocal) {
        tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 italic">Conectando con el almacenamiento local...</td></tr>`;
        return;
    }

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
                    <button onclick="moverAPapelera('${item.cedula}')" class="btn-borrar-dinamico ${claseOcultarBorrado} bg-rose-600 hover:bg-rose-700 text-white px-2 py-1 rounded text-xs font-semibold transition shadow cursor-pointer" title="Eliminar y enviar a papelera">🗑️</button>
                </td>
            `;
            tbody.appendChild(tr);
            cursor.continue();
        } else if (total === 0) {
            tbody.innerHTML = `<tr><td colspan="8" class="p-8 text-center text-gray-400 italic">No hay registros cargados.</td></tr>`;
        }
    };
}

// 4. OPERACIONES DE ESCRITURA MANUAL Y ACTUALIZACIÓN
function guardarPersonaManual() {
    const cedula = document.getElementById('manual-cedula').value.trim();
    const nombre = document.getElementById('manual-nombre').value.trim();
    const status = document.getElementById('manual-status').value.trim();
    const serial = document.getElementById('manual-serial').value.trim();
    const statusFactura = document.getElementById('manual-stf').value.trim();
    const fecha = document.getElementById('manual-fecha').value;
    const pago = document.getElementById('manual-pago').checked;

    if (!cedula || !nombre) return alert("Pana, Cédula y Nombre son completamente obligatorios.");

    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    transaccion.objectStore("registros").put({ cedula, nombre, status, serial, statusFactura, fecha, pago });

    transaccion.oncomplete = function() {
        mostrarNotificacion(`Persona [${cedula}] registrada exitosamente.`, "success");
        renderizarTablaDesdeDB();
        document.getElementById('manual-cedula').value = "";
        document.getElementById('manual-nombre').value = "";
        cerrarTodosLosDrawers();
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
        mostrarNotificacion(`Cambios aplicados a la cédula ${cedulaKey}.`, "success");
    };
}

// 5. FLUJO DE BORRADO SEGURO (PAPELERA RECICLABLE)
function moverAPapelera(cedulaKey) {
    const txLectura = baseDatosLocal.transaction(["registros"], "readonly");
    txLectura.objectStore("registros").get(cedulaKey).onsuccess = function(e) {
        const registroOriginal = e.target.result;
        if (!registroOriginal) return;

        const txEscritura = baseDatosLocal.transaction(["registros", "registros_eliminados"], "readwrite");
        txEscritura.objectStore("registros_eliminados").put(registroOriginal);
        txEscritura.objectStore("registros").delete(cedulaKey);

        txEscritura.oncomplete = function() {
            mostrarNotificacion(`Registro ${cedulaKey} enviado a la papelera.`, "error");
            renderizarTablaDesdeDB();
        };
    };
}

function renderizarPapelera() {
    const contenedor = document.getElementById('lista-papelera');
    if (!contenedor) return;
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
            contenedor.innerHTML = totalBorrados === 0 ? `<p class="italic text-gray-400 text-center py-4">La papelera está vacía.</p>` : listadoHtml;
        }
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

function vaciarPapeleraPorCompleto() {
    if (!confirm("¿Pana, seguro que quieres purgar la papelera definitivamente? Esto no se puede deshacer.")) return;
    baseDatosLocal.transaction(["registros_eliminados"], "readwrite").objectStore("registros_eliminados").clear().onsuccess = function() {
        mostrarNotificacion("Papelera purgada definitivamente.", "success");
        renderizarPapelera();
    };
}

function limpiarTodaLaDB() {
    if(!confirm("¿Seguro que quieres vaciar la tabla por completo? Se perderán todos los datos actuales.")) return;
    baseDatosLocal.transaction(["registros"], "readwrite").objectStore("registros").clear().onsuccess = function() {
        mostrarNotificacion("Almacenamiento vaciado.", "success");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}

// 6. NOTIFICACIONES ALERTA
function mostrarNotificacion(mensaje, tipo) {
    const contenedor = document.getElementById('zona-alerta');
    if (!contenedor) return;
    const estilosColor = tipo === "success" ? "bg-emerald-100 text-emerald-800 border-emerald-200" : "bg-rose-100 text-rose-800 border-rose-200";
    contenedor.innerHTML = `<div class="${estilosColor} border p-3 rounded-xl mb-4 text-xs font-bold shadow-sm text-center">${mensaje}</div>`;
    setTimeout(() => { contenedor.innerHTML = ""; }, 4000);
}