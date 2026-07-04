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
    mostrarNotificacion("Estructura de datos lista y blindada.", "success");
    renderizarTablaDesdeDB();
    document.getElementById('manual-fecha').value = new Date().toISOString().split('T')[0];
};

function limpiarTodaLaDB() {
    if(!confirm("Hermano, ¿seguro que quieres vaciar la tabla por completo?")) return;
    baseDatosLocal.transaction(["registros"], "readwrite").objectStore("registros").clear().onsuccess = function() {
        mostrarNotificacion("Almacenamiento vaciado.", "success");
        renderizarTablaDesdeDB();
        cerrarTodosLosDrawers();
    };
}