// js/interfaz.js
let modoEliminacionActivo = false;

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
        if(btnForm) btnForm.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold";
        if(btnTrash) btnTrash.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer";
        if(conForm) conForm.classList.remove('hidden');
        if(conTrash) conTrash.classList.add('hidden');
    } else {
        if(btnTrash) btnTrash.className = "flex-1 py-2 text-center border-b-2 border-slate-800 text-slate-800 cursor-pointer font-bold";
        if(btnForm) btnForm.className = "flex-1 py-2 text-center border-b-2 border-transparent text-gray-400 hover:text-gray-600 cursor-pointer";
        if(conTrash) conTrash.classList.remove('hidden');
        if(conForm) conForm.classList.add('hidden');
        renderizarPapelera();
    }
}

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
            tr.className = "hover:bg-slate-50 transition text-gray-700";
            const claseOcultarBorrado = modoEliminacionActivo ? "" : "hidden";

            tr.innerHTML = `
                <td class="p-3 font-bold text-indigo-950 bg-indigo-50/40">${item.cedula}</td>
                <td class="p-2"><input type="text" id="nom-${item.cedula}" value="${item.nombre}" class="w-full p-1 border rounded text-xs bg-transparent"></td>
                <td class="p-2"><input type="text" id="sta-${item.cedula}" value="${item.status}" class="w-full p-1 border rounded text-xs bg-transparent"></td>
                <td class="p-2"><input type="text" id="ser-${item.cedula}" value="${item.serial}" class="w-full p-1 border rounded text-xs bg-transparent"></td>
                <td class="p-2"><input type="text" id="stf-${item.cedula}" value="${item.statusFactura}" class="w-full p-1 border rounded text-xs bg-transparent"></td>
                <td class="p-2"><input type="text" id="fec-${item.cedula}" value="${item.fecha}" class="w-full p-1 border rounded text-xs bg-transparent"></td>
                <td class="p-2 text-center"><input type="checkbox" id="pag-${item.cedula}" ${item.pago ? 'checked' : ''} class="w-5 h-5 cursor-pointer"></td>
                <td class="p-2 text-center flex items-center justify-center gap-1">
                    <button onclick="guardarEdicionFila('${item.cedula}')" class="bg-indigo-600 text-white px-2.5 py-1 rounded text-xs cursor-pointer">💾</button>
                    <button onclick="moverAPapelera('${item.cedula}')" class="btn-borrar-dinamico ${claseOcultarBorrado} bg-rose-600 text-white px-2 py-1 rounded text-xs cursor-pointer">🗑️</button>
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
    transaccion.oncomplete = function() { alert(`Cédula ${cedulaKey} actualizada.`); };
}

function guardarPersonaManual() {
    const cedula = document.getElementById('manual-cedula').value.trim();
    const nombre = document.getElementById('manual-nombre').value.trim();
    const status = document.getElementById('manual-status').value.trim();
    const serial = document.getElementById('manual-serial').value.trim();
    const statusFactura = document.getElementById('manual-stf').value.trim();
    const fecha = document.getElementById('manual-fecha').value;
    const pago = document.getElementById('manual-pago').checked;

    if (!cedula || !nombre) return alert("Cédula y Nombre requeridos.");
    const transaccion = baseDatosLocal.transaction(["registros"], "readwrite");
    transaccion.objectStore("registros").put({ cedula, nombre, status, serial, statusFactura, fecha, pago });
    transaccion.oncomplete = function() {
        alert(`Guardado de forma manual.`);
        renderizarTablaDesdeDB();
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
            h += `<div class="p-2 bg-slate-50 border rounded-lg flex justify-between items-center"><div><p class="font-bold">V-${item.cedula}</p><p class="text-gray-500">${item.nombre}</p></div><button onclick="restaurarDesdePapelera('${item.cedula}')" class="bg-emerald-600 text-white text-[10px] px-2 py-1 rounded">Restaurar</button></div>`;
            cursor.continue();
        } else { contenedor.innerHTML = t === 0 ? `<p class="italic text-center py-4 text-gray-400">Papelera vacía.</p>` : h; }
    };
}

function vaciarPapeleraPorCompleto() {
    if (!confirm("¿Deseas vaciar la papelera?")) return;
    baseDatosLocal.transaction(["registros_eliminados"], "readwrite").objectStore("registros_eliminados").clear().onsuccess = function() {
        renderizarPapelera();
    };
}