// js/logica.js

// Configuración Global del Motor de Base de Datos Local
const NOMBRE_BD = 'SistemaFacturasDB';
const VERSION_BD = 1;
const ALMACEN_OBJETOS = 'facturas';
let bdLocal = null;

// Inicialización automática al cargar el documento
document.addEventListener('DOMContentLoaded', () => {
    conectarBaseDatos().then(() => {
        renderizarFacturas();
    }).catch(err => console.error("Error crítico de inicialización:", err));

    // Registro de Controladores de Eventos del Sistema
    document.getElementById('form-factura').addEventListener('submit', procesarGuardadoFactura);
    document.getElementById('input-buscar').addEventListener('input', filtrarFacturasEnPantalla);
    
    // Disparadores para la gestión con Excel
    const inputOcultoExcel = document.getElementById('input-archivo-excel');
    document.getElementById('btn-importar-excel').addEventListener('click', () => inputOcultoExcel.click());
    inputOcultoExcel.addEventListener('change', procesarImportacionExcel);
    document.getElementById('btn-exportar-excel').addEventListener('click', procesarExportacionExcel);
});

// Abre o crea la base de datos IndexedDB nativa del navegador
function conectarBaseDatos() {
    return new Promise((resolver, rechazar) => {
        const peticion = indexedDB.open(NOMBRE_BD, VERSION_BD);

        peticion.onupgradeneeded = (evento) => {
            const db = evento.target.result;
            if (!db.objectStoreNames.contains(ALMACEN_OBJETOS)) {
                db.createObjectStore(ALMACEN_OBJETOS, { keyPath: 'id', autoIncrement: true });
            }
        };

        peticion.onsuccess = (evento) => {
            bdLocal = evento.target.result;
            resolver(bdLocal);
        };

        peticion.onerror = (evento) => {
            rechazar(evento.target.error);
        };
    });
}

// Obtiene todas las facturas guardadas de manera asíncrona
function obtenerTodasLasFacturas() {
    return new Promise((resolver, rechazar) => {
        if (!bdLocal) return resolver([]);
        const transaccion = bdLocal.transaction([ALMACEN_OBJETOS], 'readonly');
        const almacen = transaccion.objectStore(ALMACEN_OBJETOS);
        const peticion = almacen.getAll();

        peticion.onsuccess = () => resolver(peticion.result);
        peticion.onerror = () => rechazar(peticion.error);
    });
}

// Guarda un registro nuevo en la base de datos local
function guardarFacturaEnBD(factura) {
    return new Promise((resolver, rechazar) => {
        const transaccion = bdLocal.transaction([ALMACEN_OBJETOS], 'readwrite');
        const almacen = transaccion.objectStore(ALMACEN_OBJETOS);
        const peticion = almacen.add(factura);

        peticion.onsuccess = () => resolver();
        peticion.onerror = () => rechazar(peticion.error);
    });
}

// Elimina una factura específica utilizando su ID único
function eliminarFacturaDeBD(id) {
    return new Promise((resolver, rechazar) => {
        const transaccion = bdLocal.transaction([ALMACEN_OBJETOS], 'readwrite');
        const almacen = transaccion.objectStore(ALMACEN_OBJETOS);
        const peticion = almacen.delete(Number(id));

        peticion.onsuccess = () => resolver();
        peticion.onerror = () => rechazar(peticion.error);
    });
}

// Procesa el formulario para añadir una nueva factura
async function procesarGuardadoFactura(evento) {
    evento.preventDefault();
    
    const nuevaFactura = {
        cliente: document.getElementById('input-cliente').value.trim(),
        fecha: document.getElementById('input-fecha').value,
        monto: parseFloat(document.getElementById('input-monto').value),
        estado: document.getElementById('select-estado').value
    };

    try {
        await guardarFacturaEnBD(nuevaFactura);
        document.getElementById('form-factura').reset();
        renderizarFacturas();
    } catch (err) {
        alert("Ocurrió un error al guardar los datos localmente.");
        console.error(err);
    }
}

// Dibuja las filas de la tabla leyendo IndexedDB en tiempo real
async function renderizarFacturas(filtroBusqueda = "") {
    const cuerpoTabla = document.getElementById('tabla-cuerpo-facturas');
    const panelVacio = document.getElementById('mensaje-vacio');
    cuerpoTabla.innerHTML = "";

    try {
        const registros = await obtenerTodasLasFacturas();
        
        // Aplica el criterio de búsqueda si el usuario ha escrito algo
        const registrosFiltrados = registros.filter(f => 
            f.cliente.toLowerCase().includes(filtroBusqueda.toLowerCase())
        );

        if (registrosFiltrados.length === 0) {
            panelVacio.classList.remove('hidden');
            return;
        }
        panelVacio.classList.add('hidden');

        // Generación dinámica de la estructura de filas con estilos estilizados de Tailwind
        registrosFiltrados.forEach(factura => {
            const fila = document.createElement('tr');
            fila.className = "hover:bg-slate-50 transition-colors animar-fila";
            
            // Define el distintivo visual del estado de la factura
            let estiloBadge = "bg-slate-100 text-slate-700";
            if (factura.estado.includes("Pagada")) estiloBadge = "bg-emerald-50 text-emerald-700 border border-emerald-200";
            if (factura.estado.includes("Pendiente")) estiloBadge = "bg-amber-50 text-amber-700 border border-amber-200";
            if (factura.estado.includes("Anulada")) estiloBadge = "bg-rose-50 text-rose-700 border border-rose-200";

            fila.innerHTML = `
                <td class="p-3 font-medium text-slate-900">${factura.cliente}</td>
                <td class="p-3 text-center text-slate-500">${factura.fecha}</td>
                <td class="p-3 text-right font-semibold text-slate-900">$${factura.monto.toFixed(2)}</td>
                <td class="p-3 text-center">
                    <span class="px-2.5 py-1 rounded-full text-xs font-medium inline-block ${estiloBadge}">
                        ${factura.estado}
                    </span>
                </td>
                <td class="p-3 text-center">
                    <button data-id="${factura.id}" class="btn-borrar text-rose-600 hover:text-rose-900 bg-rose-50 hover:bg-rose-100 p-1.5 rounded-md transition-all cursor-pointer text-xs" title="Eliminar registro">
                        🗑️ Borrar
                    </button>
                </td>
            `;
            cuerpoTabla.appendChild(fila);
        });

        // Vincula la acción de borrado a los botones recién inyectados
        document.querySelectorAll('.btn-borrar').forEach(boton => {
            boton.addEventListener('click', async (e) => {
                const idRegistro = e.target.getAttribute('data-id');
                if (confirm("¿Estás seguro de que deseas eliminar de forma permanente esta factura?")) {
                    await eliminarFacturaDeBD(idRegistro);
                    renderizarFacturas(document.getElementById('input-buscar').value);
                }
            });
        });

    } catch (err) {
        console.error("Fallo al renderizar la tabla de contenidos:", err);
    }
}

// Filtra la tabla de forma fluida mientras el usuario escribe
function filtrarFacturasEnPantalla(evento) {
    renderizarFacturas(evento.target.value);
}

// EXPORTACIÓN A EXCEL: Convierte los datos de IndexedDB a un documento .xlsx real
async function procesarExportacionExcel() {
    try {
        const datos = await obtenerTodasLasFacturas();
        if (datos.length === 0) {
            alert("No hay registros disponibles para exportar.");
            return;
        }

        // Mapea y limpia las claves internas para el reporte impreso
        const datosLimpios = datos.map(f => ({
            'Cliente': f.cliente,
            'Fecha de Emisión': f.fecha,
            'Monto Total ($)': f.monto,
            'Estado Actual': f.estado
        }));

        const libroTrabajo = XLSX.utils.book_new();
        const hojaTrabajo = XLSX.utils.json_to_sheet(datosLimpios);
        XLSX.utils.book_append_sheet(libroTrabajo, hojaTrabajo, 'Facturas Registradas');
        
        // Genera la descarga automática del archivo en el dispositivo
        XLSX.writeFile(libroTrabajo, 'Reporte_Facturas_Local.xlsx');
    } catch (err) {
        alert("Ocurrió un error al procesar el archivo Excel de salida.");
        console.error(err);
    }
}

// IMPORTACIÓN DE EXCEL: Lee un archivo cargado y guarda los registros válidos de forma masiva
function procesarImportacionExcel(evento) {
    const archivo = evento.target.files[0];
    if (!archivo) return;

    const lectorArchivos = new FileReader();
    lectorArchivos.onload = async (e) => {
        try {
            const bytesDatos = new Uint8Array(e.target.result);
            const libroTrabajo = XLSX.read(bytesDatos, { type: 'array' });
            const primerNombreHoja = libroTrabajo.SheetNames[0];
            const hojaTrabajo = libroTrabajo.Sheets[primerNombreHoja];
            const registrosJson = XLSX.utils.sheet_to_json(hojaTrabajo);

            let guardadosConExito = 0;
            for (const item of registrosJson) {
                // Validación estricta de las columnas requeridas
                const cliente = item['Cliente'] || item['cliente'];
                const fecha = item['Fecha de Emisión'] || item['fecha'];
                const monto = item['Monto Total ($)'] || item['monto'];
                const estadoBruto = item['Estado Actual'] || item['estado'] || 'Pendiente';

                if (cliente && fecha && monto !== undefined) {
                    // Normaliza el estado agregándole su emoji visual si no lo trae
                    let estadoNormalizado = estadoBruto;
                    if (estadoBruto === 'Pagada' || estadoBruto === '✅ Pagada') estadoNormalizado = '✅ Pagada';
                    if (estadoBruto === 'Pendiente' || estadoBruto === '⏳ Pendiente') estadoNormalizado = '⏳ Pendiente';
                    if (estadoBruto === 'Anulada' || estadoBruto === '❌ Anulada') estadoNormalizado = '❌ Anulada';

                    await guardarFacturaEnBD({
                        cliente: String(cliente).trim(),
                        fecha: String(fecha),
                        monto: parseFloat(monto),
                        estado: estadoNormalizado
                    });
                    guardadosConExito++;
                }
            }

            alert(`Importación completada con éxito. Se añadieron ${guardadosConExito} registros.`);
            document.getElementById('input-archivo-excel').value = ""; // Limpia el input de carga
            renderizarFacturas();
        } catch (err) {
            alert("No se pudo procesar el archivo Excel. Asegúrate de que posea el formato correcto.");
            console.error(err);
        }
    };
    lectorArchivos.readAsArrayBuffer(archivo);
}