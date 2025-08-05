let pdfTextContent = '';
let currentFontSize = 14;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    initializeEventListeners();
    setupDragAndDrop();
  }
});

// Inicializar todos los event listeners
function initializeEventListeners() {
    document.getElementById("pdf-file").addEventListener("change", handlePdfUpload);
    document.getElementById("select-file-btn").addEventListener("click", () => {
        document.getElementById("pdf-file").click();
    });
    document.getElementById("search-button").addEventListener("click", searchAndInsertArticle);
    document.getElementById("insert-selection-button").addEventListener("click", insertSelectedText);
    document.getElementById("zoom-in").addEventListener("click", () => adjustFontSize(2));
    document.getElementById("zoom-out").addEventListener("click", () => adjustFontSize(-2));
    
    // Mejorar la experiencia de búsqueda con Enter
    document.getElementById("search-term").addEventListener("keypress", (e) => {
        if (e.key === 'Enter') {
            searchAndInsertArticle();
        }
    });
}

// Configurar drag and drop
function setupDragAndDrop() {
    const uploadArea = document.getElementById('upload-area');
    
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, preventDefaults, false);
    });
    
    ['dragenter', 'dragover'].forEach(eventName => {
        uploadArea.addEventListener(eventName, highlight, false);
    });
    
    ['dragleave', 'drop'].forEach(eventName => {
        uploadArea.addEventListener(eventName, unhighlight, false);
    });
    
    uploadArea.addEventListener('drop', handleDrop, false);
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function highlight(e) {
    document.getElementById('upload-area').classList.add('drag-over');
}

function unhighlight(e) {
    document.getElementById('upload-area').classList.remove('drag-over');
}

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;
    
    if (files.length > 0) {
        const file = files[0];
        if (file.type === "application/pdf") {
            // Simular la selección del archivo
            const fileInput = document.getElementById("pdf-file");
            const dataTransfer = new DataTransfer();
            dataTransfer.items.add(file);
            fileInput.files = dataTransfer.files;
            
            // Disparar el evento de cambio
            handlePdfUpload({ target: { files: [file] } });
        } else {
            showMessage('Por favor, selecciona un archivo PDF válido.', 'error');
        }
    }
}

// Función mejorada para manejar la carga del PDF
function handlePdfUpload(event) {
    const file = event.target.files[0];
    if (!file || file.type !== "application/pdf") {
        showMessage('Por favor, selecciona un archivo PDF válido.', 'error');
        return;
    }

    // Validar tamaño del archivo (10MB máximo)
    const maxSize = 10 * 1024 * 1024; // 10MB en bytes
    if (file.size > maxSize) {
        showMessage('El archivo es demasiado grande. Máximo 10MB permitido.', 'error');
        return;
    }

    const fileReader = new FileReader();
    const loadingSpinner = document.getElementById('loading-spinner');
    const mainControls = document.getElementById('main-controls');
    const uploadStep = document.getElementById('upload-step');

    // Mostrar estado de carga
    loadingSpinner.style.display = 'block';
    mainControls.style.display = 'none';
    uploadStep.style.display = 'none';
    pdfTextContent = '';

    fileReader.onload = function() {
        const typedarray = new Uint8Array(this.result);
        
        pdfjsLib.getDocument(typedarray).promise.then(pdf => {
            let textPromises = [];
            const totalPages = pdf.numPages;
            
            // Actualizar información del documento
            updateDocumentInfo(file.name, totalPages);
            
            for (let i = 1; i <= pdf.numPages; i++) {
                textPromises.push(
                    pdf.getPage(i).then(page => page.getTextContent())
                );
            }
            return Promise.all(textPromises);
        }).then(textContents => {
            // Procesar y almacenar el texto
            pdfTextContent = textContents.map((content, index) => {
                const pageText = content.items.map(item => item.str).join(' ');
                return `=== PÁGINA ${index + 1} ===\n${pageText}`;
            }).join('\n\n');
            
            // Renderizar vista previa
            renderPdfAsText(textContents);

            // Ocultar loading y mostrar controles
            loadingSpinner.style.display = 'none';
            mainControls.style.display = 'block';
            
            showMessage('Documento cargado y procesado correctamente.', 'success');
        }).catch(error => {
            console.error("Error al procesar el PDF:", error);
            loadingSpinner.style.display = 'none';
            uploadStep.style.display = 'block';
            showMessage('Error al procesar el PDF. Verifica que el archivo no esté dañado.', 'error');
        });
    };
    
    fileReader.readAsArrayBuffer(file);
}

// Actualizar información del documento
function updateDocumentInfo(fileName, totalPages) {
    document.getElementById('document-name').textContent = fileName;
    document.getElementById('document-stats').textContent = 
        `${totalPages} página${totalPages !== 1 ? 's' : ''} • Procesado correctamente`;
}

// Renderizar PDF como texto mejorado
function renderPdfAsText(textContents) {
    const viewer = document.getElementById('pdf-viewer');
    viewer.innerHTML = '';
    
    textContents.forEach((content, pageIndex) => {
        // Crear separador de página
        if (pageIndex > 0) {
            const separator = document.createElement('div');
            separator.style.cssText = `
                border-top: 2px solid var(--border-color);
                margin: 20px 0;
                padding-top: 15px;
                font-weight: 600;
                color: var(--neutral-secondary);
                font-size: 12px;
            `;
            separator.textContent = `--- PÁGINA ${pageIndex + 1} ---`;
            viewer.appendChild(separator);
        }
        
        const pageDiv = document.createElement('div');
        pageDiv.style.cssText = `
            margin-bottom: 20px;
            padding: 15px;
            background: ${pageIndex % 2 === 0 ? '#fafafa' : '#ffffff'};
            border-radius: 4px;
            border-left: 3px solid var(--primary-color);
            user-select: text;
            cursor: text;
        `;
        
        const pageText = content.items.map(item => item.str).join(' ');
        pageDiv.textContent = pageText;
        
        // Resaltar artículos automáticamente
        highlightArticles(pageDiv);
        
        viewer.appendChild(pageDiv);
    });
    
    // Aplicar tamaño de fuente actual
    viewer.style.fontSize = currentFontSize + 'px';
}

// Función para resaltar artículos en el texto
function highlightArticles(element) {
    const articleRegex = /((?:Artículo|Articulo|Art\.)\s+\d+[^\n]*)/gi;
    const text = element.textContent;
    const matches = text.match(articleRegex);
    
    if (matches) {
        let highlightedText = text;
        matches.forEach(match => {
            highlightedText = highlightedText.replace(
                match, 
                `<mark style="background: #fff3cd; padding: 2px 4px; border-radius: 3px; font-weight: 600;">${match}</mark>`
            );
        });
        element.innerHTML = highlightedText;
    }
}

// Ajustar tamaño de fuente
function adjustFontSize(change) {
    currentFontSize = Math.max(10, Math.min(24, currentFontSize + change));
    const viewer = document.getElementById('pdf-viewer');
    viewer.style.fontSize = currentFontSize + 'px';
    
    // Feedback visual
    showMessage(`Tamaño de fuente: ${currentFontSize}px`, 'info', 1500);
}

// Búsqueda mejorada de artículos
function searchAndInsertArticle() {
    const searchTerm = document.getElementById("search-term").value.trim();
    const resultElement = document.getElementById("search-result");

    if (!pdfTextContent) {
        showMessage("Primero debes cargar un PDF.", 'error');
        return;
    }
    if (!searchTerm) {
        showMessage("Por favor, ingresa un número de artículo.", 'error');
        return;
    }

    // Limpiar resultado anterior
    resultElement.textContent = '';
    
    // Mostrar estado de búsqueda
    showMessage("Buscando artículo...", 'info');

    // Expresión regular mejorada y más flexible
    const patterns = [
        `(?:Artículo|Articulo|Art\\.)\\s+${searchTerm}(?:\\s|\\.|:)[\\s\\S]*?(?=(?:Artículo|Articulo|Art\\.)\\s+\\d|CAPÍTULO|TÍTULO|===|$)`,
        `Art\\s+${searchTerm}[\\s\\S]*?(?=Art\\s+\\d|CAPÍTULO|TÍTULO|===|$)`,
        `${searchTerm}\\.[\\s\\S]*?(?=\\d+\\.|CAPÍTULO|TÍTULO|===|$)`
    ];

    let foundText = null;
    let matchPattern = null;

    // Probar cada patrón
    for (let pattern of patterns) {
        const regex = new RegExp(pattern, 'i');
        const match = pdfTextContent.match(regex);
        
        if (match && match[0]) {
            foundText = match[0].trim();
            matchPattern = pattern;
            break;
        }
    }

    if (foundText) {
        // Limpiar el texto encontrado
        foundText = cleanExtractedText(foundText);
        
        showMessage(`✅ Artículo ${searchTerm} encontrado e insertado correctamente.`, 'success');
        insertTextIntoWord(foundText);
        
        // Resaltar en la vista previa
        highlightSearchResult(searchTerm);
    } else {
        showMessage(`❌ Artículo "${searchTerm}" no encontrado. Verifica el número e intenta de nuevo.`, 'error');
    }
}

// Limpiar texto extraído
function cleanExtractedText(text) {
    return text
        .replace(/===[^===]*===/g, '') // Remover separadores de página
        .replace(/\s+/g, ' ') // Normalizar espacios
        .replace(/^\s+|\s+$/g, '') // Trim
        .replace(/([.!?])\s*([A-ZÁÉÍÓÚÑ])/g, '$1\n\n$2'); // Mejorar formato de párrafos
}

// Resaltar resultado de búsqueda en la vista previa
function highlightSearchResult(searchTerm) {
    const viewer = document.getElementById('pdf-viewer');
    const regex = new RegExp(`((?:Artículo|Articulo|Art\\.)\\s+${searchTerm}[^\\n]*)`, 'gi');
    
    const content = viewer.innerHTML;
    const highlightedContent = content.replace(regex, 
        '<span style="background: #ffeb3b; padding: 3px 6px; border-radius: 4px; font-weight: bold; animation: pulse 2s;">$1</span>'
    );
    
    viewer.innerHTML = highlightedContent;
    
    // Scroll al resultado
    const highlighted = viewer.querySelector('span[style*="ffeb3b"]');
    if (highlighted) {
        highlighted.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

// Insertar texto seleccionado mejorado
function insertSelectedText() {
    const selectedText = window.getSelection().toString().trim();
    
    if (selectedText) {
        if (selectedText.length < 10) {
            showMessage("⚠️ El texto seleccionado es muy corto. ¿Estás seguro de que es correcto?", 'warning');
        }
        
        const cleanedText = cleanExtractedText(selectedText);
        insertTextIntoWord(cleanedText);
        showMessage(`✅ Texto seleccionado insertado correctamente (${selectedText.length} caracteres).`, 'success');
        
        // Limpiar selección
        window.getSelection().removeAllRanges();
    } else {
        showMessage("❌ Por favor, selecciona primero un fragmento de texto de la vista previa.", 'error');
    }
}

// Función genérica mejorada para insertar texto en Word
function insertTextIntoWord(textToInsert) {
    Word.run(async (context) => {
        try {
            const selection = context.document.getSelection();
            
            // Insertar texto con formato mejorado
            const insertedText = selection.insertText(textToInsert + '\n\n', Word.InsertLocation.end);
            
            // Aplicar formato básico
            insertedText.font.name = "Calibri";
            insertedText.font.size = 11;
            insertedText.paragraphFormat.spaceAfter = 6;
            
            await context.sync();
            
            showMessage("✅ Texto insertado correctamente en el documento.", 'success');
        } catch (error) {
            console.error("Error al insertar texto:", error);
            showMessage("❌ Error al insertar texto. Verifica que Word esté disponible.", 'error');
        }
    }).catch(function (error) {
        console.error("Error en Word.run:", error);
        showMessage("❌ Error de conexión con Word.", 'error');
    });
}

// Sistema mejorado de mensajes
function showMessage(message, type = 'info', duration = 4000) {
    const resultElement = document.getElementById("search-result");
    
    // Limpiar clases anteriores
    resultElement.className = 'result-message';
    
    // Añadir clase según el tipo
    resultElement.classList.add(type);
    resultElement.textContent = message;
    
    // Auto-ocultar mensajes después del tiempo especificado
    if (duration > 0) {
        setTimeout(() => {
            resultElement.textContent = '';
            resultElement.className = 'result-message';
        }, duration);
    }
}

// Agregar animación de pulso para elementos destacados
const style = document.createElement('style');
style.textContent = `
    @keyframes pulse {
        0% { transform: scale(1); }
        50% { transform: scale(1.05); }
        100% { transform: scale(1); }
    }
`;
document.head.appendChild(style);

// Mejorar accesibilidad con atajos de teclado
document.addEventListener('keydown', (e) => {
    // Ctrl/Cmd + Enter para buscar
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        if (document.getElementById("search-term") === document.activeElement) {
            searchAndInsertArticle();
        }
    }
    
    // Ctrl/Cmd + I para insertar selección
    if ((e.ctrlKey || e.metaKey) && e.key === 'i') {
        e.preventDefault();
        insertSelectedText();
    }
    
    // Ctrl/Cmd + Plus/Minus para zoom
    if ((e.ctrlKey || e.metaKey) && e.key === '+') {
        e.preventDefault();
        adjustFontSize(2);
    }
    
    if ((e.ctrlKey || e.metaKey) && e.key === '-') {
        e.preventDefault();
        adjustFontSize(-2);
    }
});

// Inicialización adicional cuando el DOM esté listo
document.addEventListener('DOMContentLoaded', () => {
    // Configurar tooltips para mejor UX
    const buttons = document.querySelectorAll('button');
    buttons.forEach(button => {
        if (!button.title && button.textContent) {
            const text = button.textContent.trim();
            if (text.includes('Buscar')) {
                button.title = 'Buscar artículo e insertar en el documento (Ctrl+Enter)';
            } else if (text.includes('Insertar')) {
                button.title = 'Insertar texto seleccionado (Ctrl+I)';
            }
        }
    });
    
    // Añadir información sobre atajos de teclado
    const footer = document.querySelector('.footer-text');
    if (footer) {
        footer.innerHTML += ' • <span title="Ctrl+Enter: Buscar, Ctrl+I: Insertar, Ctrl++/-: Zoom">Atajos disponibles</span>';
    }
});