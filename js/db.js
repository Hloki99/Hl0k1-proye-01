// js/db.js
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
    console.log("IndexedDB inicializada con éxito.");
    if (typeof renderizarTablaDesdeDB === "function") {
        renderizarTablaDesdeDB();
    }
    const campoFecha = document.getElementById('manual-fecha');
    if (campoFecha) campoFecha.value = new Date().toISOString().split('T')[0];
};

peticionDB.onerror = function(e) {
    console.error("Error al abrir IndexedDB:", e.target.error);
};

function limpiarTodaLaDB() {
    if(!confirm("Hermano, ¿seguro que quieres vaciar la tabla por completo?")) return;
    baseDatosLocal.transaction(["registros"], "readwrite").objectStore("registros").clear().onsuccess = function() {
        alert("Almacenamiento vaciado.");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}