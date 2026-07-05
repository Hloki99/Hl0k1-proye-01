// js/db.js

// Inicialización del Estado Global de la base de datos
var baseDatosLocal; 

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
    // Si la interfaz ya cargó sus funciones de renderizado, pintamos la tabla
    if (typeof renderizarTablaDesdeDB === "function") {
        renderizarTablaDesdeDB();
    }
    // Setea fecha actual por defecto en el formulario manual
    const inputFecha = document.getElementById('manual-fecha');
    if (inputFecha) {
        inputFecha.value = new Date().toISOString().split('T')[0];
    }
};

peticionDB.onerror = function(e) {
    console.error("Error abriendo IndexedDB:", e.target.error);
};