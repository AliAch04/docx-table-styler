/**
 * DOCX Table Unifier - Web Version
 * Pure JavaScript implementation for browser-based DOCX processing
 */

// Theme definitions
const THEMES = {
    '1': {
        name: 'Professional Blue',
        description: 'Clean, professional with blue headers',
        preferredStyles: ['Normal Table', 'Light Shading Accent 1', 'Medium Shading 1 Accent 1'],
        colors: {
            headerBg: 'D9E1F2',
            borderStyle: 'single',
            fontSize: 10,
            headerBold: true
        }
    },
    '2': {
        name: 'Minimalist Light',
        description: 'Minimal, clean with light borders',
        preferredStyles: ['Plain Table 3', 'Light List Accent 1', 'Light Grid Accent 1'],
        colors: {
            headerBg: 'F2F2F2',
            borderStyle: 'single',
            fontSize: 10,
            headerBold: true
        }
    },
    '3': {
        name: 'Modern Grid',
        description: 'Clear grid with alternating rows',
        preferredStyles: ['Table Grid', 'Medium Grid 1 Accent 1', 'Light Grid Accent 1'],
        colors: {
            headerBg: 'E6F0FA',
            borderStyle: 'single',
            fontSize: 10,
            headerBold: true,
            alternatingRows: true
        }
    },
    '4': {
        name: 'Academic Report',
        description: 'Traditional academic table style',
        preferredStyles: ['Light Shading Accent 1', 'Medium Shading 1 Accent 1', 'Normal Table'],
        colors: {
            headerBg: '4472C4',
            headerText: 'FFFFFF',
            borderStyle: 'single',
            fontSize: 11,
            headerBold: true
        }
    },
    '5': {
        name: 'Corporate Dark',
        description: 'Bold, corporate look',
        preferredStyles: ['Medium Shading 1 Accent 2', 'Dark List Accent 1', 'Normal Table'],
        colors: {
            headerBg: '44546A',
            headerText: 'FFFFFF',
            borderStyle: 'single',
            fontSize: 10,
            headerBold: true
        }
    },
    '6': {
        name: 'Simple Borders Only',
        description: 'Just add borders, keep existing formatting',
        preferredStyles: [],
        colors: {
            borderStyle: 'single'
        }
    }
};

class DocxTableUnifier {
    constructor() {
        this.file = null;
        this.zip = null;
        this.documentXml = null;
        this.stylesXml = null;
        this.availableStyles = [];
        this.tables = [];
        this.styleMapping = {};
        this.selectedTheme = null;
        this.customStyle = null;
    }

    // Load and parse DOCX file
    async loadFile(file) {
        this.file = file;
        this.zip = await JSZip.loadAsync(file);
        
        // Extract main document and styles
        this.documentXml = await this.getXmlContent('word/document.xml');
        this.stylesXml = await this.getXmlContent('word/styles.xml');
        
        // Parse available table styles
        this.parseAvailableStyles();
        
        // Find all tables
        this.findTables();
        
        return {
            tableCount: this.tables.length,
            styles: this.availableStyles
        };
    }

    async getXmlContent(path) {
        const file = this.zip.file(path);
        if (!file) return null;
        const content = await file.async('string');
        return new DOMParser().parseFromString(content, 'application/xml');
    }

    parseAvailableStyles() {
        if (!this.stylesXml) return;
        
        const styles = this.stylesXml.getElementsByTagName('w:style');
        this.availableStyles = [];
        
        for (let style of styles) {
            const type = style.getAttribute('w:type');
            const styleId = style.getAttribute('w:styleId');
            const nameElem = style.getElementsByTagName('w:name')[0];
            const name = nameElem ? nameElem.getAttribute('w:val') : styleId;
            
            if (type === 'table') {
                this.availableStyles.push({
                    id: styleId,
                    name: name,
                    type: this.categorizeStyle(name)
                });
            }
        }
    }

    categorizeStyle(name) {
        if (name.includes('Normal')) return '📊 Basic';
        if (name.includes('Plain')) return '✨ Minimal';
        if (name.includes('Light')) return '🌟 Light';
        if (name.includes('Medium')) return '⭐ Medium';
        if (name.includes('Dark')) return '🌙 Dark';
        if (name.includes('Grid')) return '🔲 Grid';
        if (name.includes('List')) return '📋 List';
        return '📌 Other';
    }

    findTables() {
        if (!this.documentXml) return;
        
        const tables = this.documentXml.getElementsByTagName('w:tbl');
        this.tables = [];
        
        for (let table of tables) {
            this.tables.push(table);
        }
    }

    findBestStyleMatch(preferredStyles) {
        if (!preferredStyles || preferredStyles.length === 0) return null;
        
        const availableNames = this.availableStyles.map(s => s.name.toLowerCase());
        
        // Try exact matches first
        for (let preferred of preferredStyles) {
            const exactMatch = this.availableStyles.find(s => 
                s.name.toLowerCase() === preferred.toLowerCase()
            );
            if (exactMatch) return exactMatch;
        }
        
        // Try partial matches
        for (let preferred of preferredStyles) {
            const prefLower = preferred.toLowerCase();
            const partialMatch = this.availableStyles.find(s => 
                s.name.toLowerCase().includes(prefLower) || 
                prefLower.includes(s.name.toLowerCase())
            );
            if (partialMatch) return partialMatch;
        }
        
        // Return first available if any
        return this.availableStyles.length > 0 ? this.availableStyles[0] : null;
    }

    generateStyleMapping() {
        this.styleMapping = {};
        
        for (let [key, theme] of Object.entries(THEMES)) {
            const match = this.findBestStyleMatch(theme.preferredStyles);
            this.styleMapping[key] = match ? match.name : null;
        }
    }

    async applyTheme(themeKey, customStyle = null) {
        this.selectedTheme = THEMES[themeKey];
        this.customStyle = customStyle;
        
        const styleToUse = customStyle || this.styleMapping[themeKey];
        
        // Process each table
        for (let i = 0; i < this.tables.length; i++) {
            const table = this.tables[i];
            
            // Apply style if available
            if (styleToUse && this.availableStyles.some(s => s.name === styleToUse)) {
                this.applyTableStyle(table, styleToUse);
            }
            
            // Apply manual formatting
            this.applyManualFormatting(table, this.selectedTheme.colors);
            
            // Update progress
            this.updateProgress(i + 1, this.tables.length);
        }
        
        // Generate modified DOCX
        return await this.generateDocx();
    }

    applyTableStyle(table, styleName) {
        // Find or create style reference
        const tblPr = table.getElementsByTagName('w:tblPr')[0];
        if (tblPr) {
            // Remove existing style
            const existingStyle = tblPr.getElementsByTagName('w:tblStyle')[0];
            if (existingStyle) {
                tblPr.removeChild(existingStyle);
            }
            
            // Add new style
            const styleElem = this.documentXml.createElement('w:tblStyle');
            styleElem.setAttribute('w:val', styleName);
            tblPr.appendChild(styleElem);
        }
    }

    applyManualFormatting(table, colors) {
        const rows = table.getElementsByTagName('w:tr');
        
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            const cells = row.getElementsByTagName('w:tc');
            
            for (let cell of cells) {
                let tcPr = cell.getElementsByTagName('w:tcPr')[0];
                if (!tcPr) {
                    tcPr = this.documentXml.createElement('w:tcPr');
                    cell.insertBefore(tcPr, cell.firstChild);
                }
                
                // Add borders
                if (colors.borderStyle) {
                    this.addBorders(tcPr, colors.borderStyle);
                }
                
                // Header row formatting
                if (i === 0) {
                    if (colors.headerBg) {
                        this.addShading(tcPr, colors.headerBg);
                    }
                    
                    if (colors.headerBold) {
                        this.makeTextBold(cell);
                    }
                    
                    if (colors.headerText) {
                        this.setTextColor(cell, colors.headerText);
                    }
                }
            }
        }
        
        // Alternating rows
        if (colors.alternatingRows && rows.length > 1) {
            for (let i = 1; i < rows.length; i += 2) {
                const cells = rows[i].getElementsByTagName('w:tc');
                for (let cell of cells) {
                    let tcPr = cell.getElementsByTagName('w:tcPr')[0];
                    if (!tcPr) {
                        tcPr = this.documentXml.createElement('w:tcPr');
                        cell.insertBefore(tcPr, cell.firstChild);
                    }
                    this.addShading(tcPr, 'F5F5F5');
                }
            }
        }
    }

    addBorders(tcPr, style) {
        const borders = this.documentXml.createElement('w:tcBorders');
        
        ['top', 'left', 'bottom', 'right'].forEach(side => {
            const border = this.documentXml.createElement(`w:${side}`);
            border.setAttribute('w:val', style);
            border.setAttribute('w:sz', '4');
            border.setAttribute('w:space', '0');
            border.setAttribute('w:color', 'auto');
            borders.appendChild(border);
        });
        
        tcPr.appendChild(borders);
    }

    addShading(tcPr, color) {
        const shading = this.documentXml.createElement('w:shd');
        shading.setAttribute('w:val', 'clear');
        shading.setAttribute('w:color', 'auto');
        shading.setAttribute('w:fill', color);
        tcPr.appendChild(shading);
    }

    makeTextBold(cell) {
        const texts = cell.getElementsByTagName('w:t');
        for (let text of texts) {
            const parent = text.parentNode;
            let rPr = parent.getElementsByTagName('w:rPr')[0];
            if (!rPr) {
                rPr = this.documentXml.createElement('w:rPr');
                parent.insertBefore(rPr, parent.firstChild);
            }
            
            const bold = this.documentXml.createElement('w:b');
            rPr.appendChild(bold);
        }
    }

    setTextColor(cell, color) {
        const texts = cell.getElementsByTagName('w:t');
        for (let text of texts) {
            const parent = text.parentNode;
            let rPr = parent.getElementsByTagName('w:rPr')[0];
            if (!rPr) {
                rPr = this.documentXml.createElement('w:rPr');
                parent.insertBefore(rPr, parent.firstChild);
            }
            
            const colorElem = this.documentXml.createElement('w:color');
            colorElem.setAttribute('w:val', color);
            rPr.appendChild(colorElem);
        }
    }

    async generateDocx() {
        // Update document in zip
        const serializer = new XMLSerializer();
        const docString = serializer.serializeToString(this.documentXml);
        this.zip.file('word/document.xml', docString);
        
        // Generate blob
        const blob = await this.zip.generateAsync({ type: 'blob' });
        return blob;
    }

    updateProgress(current, total) {
        const percent = (current / total) * 100;
        document.getElementById('progressFill').style.width = `${percent}%`;
        
        this.addLog(`Table ${current}/${total} processed`, 'success');
    }

    addLog(message, type = 'info') {
        const logSection = document.getElementById('logSection');
        const entry = document.createElement('div');
        entry.className = `log-entry log-${type}`;
        entry.textContent = `[${new Date().toLocaleTimeString()}] ${message}`;
        logSection.appendChild(entry);
        logSection.scrollTop = logSection.scrollHeight;
    }
}

// UI Controller
class UIController {
    constructor() {
        this.unifier = new DocxTableUnifier();
        this.initEventListeners();
    }

    initEventListeners() {
        // File input
        document.getElementById('fileInput').addEventListener('change', (e) => {
            this.handleFileSelect(e.target.files[0]);
        });

        // Drag and drop
        const dropZone = document.getElementById('dropZone');
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.classList.add('dragover');
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.classList.remove('dragover');
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file && file.name.endsWith('.docx')) {
                this.handleFileSelect(file);
            } else {
                alert('Please drop a valid DOCX file');
            }
        });

        // Theme selection
        document.addEventListener('click', (e) => {
            const themeCard = e.target.closest('.theme-card');
            if (themeCard) {
                const themeKey = themeCard.dataset.theme;
                this.selectTheme(themeKey);
            }
        });

        // Custom style input
        document.getElementById('customStyleInput').addEventListener('input', (e) => {
            this.unifier.customStyle = e.target.value;
        });

        // Process button
        document.getElementById('processBtn').addEventListener('click', () => {
            this.processDocument();
        });
    }

    async handleFileSelect(file) {
        if (!file) return;

        // Show file info
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('fileSize').textContent = 
            (file.size / 1024).toFixed(2) + ' KB';
        document.getElementById('fileInfo').classList.add('active');

        // Load and parse file
        const result = await this.unifier.loadFile(file);
        
        // Generate style mapping
        this.unifier.generateStyleMapping();
        
        // Update UI
        this.displayAvailableStyles();
        this.displayThemes();
        
        // Show sections
        document.getElementById('stylesSection').style.display = 'block';
        document.getElementById('themesSection').style.display = 'block';
        document.getElementById('processBtn').disabled = false;
    }

    displayAvailableStyles() {
        const grid = document.getElementById('stylesGrid');
        grid.innerHTML = '';
        
        this.unifier.availableStyles.forEach(style => {
            const card = document.createElement('div');
            card.className = 'style-card';
            card.innerHTML = `
                <div class="style-name">${style.name}</div>
                <div class="style-type">${style.type}</div>
                <div class="style-badge badge-available">✓ Available</div>
            `;
            grid.appendChild(card);
        });
    }

    displayThemes() {
        const grid = document.getElementById('themesGrid');
        grid.innerHTML = '';
        
        for (let [key, theme] of Object.entries(THEMES)) {
            const matchedStyle = this.unifier.styleMapping[key];
            const matchQuality = matchedStyle ? 'exact' : 'none';
            
            const card = document.createElement('div');
            card.className = 'theme-card';
            card.dataset.theme = key;
            
            card.innerHTML = `
                <div class="theme-header">
                    <span class="theme-number">${key}</span>
                    <span class="theme-name">${theme.name}</span>
                </div>
                <div class="theme-description">${theme.description}</div>
                <div class="theme-style">
                    <div class="style-match">
                        <span class="match-indicator match-${matchQuality}"></span>
                        <span>
                            ${matchedStyle ? 
                                `Will use: <strong>${matchedStyle}</strong>` : 
                                'No style match - will use manual formatting'}
                        </span>
                    </div>
                </div>
            `;
            
            grid.appendChild(card);
        }
    }

    selectTheme(themeKey) {
        // Remove selected class from all cards
        document.querySelectorAll('.theme-card').forEach(c => {
            c.classList.remove('selected');
        });
        
        // Add selected class to chosen card
        document.querySelector(`.theme-card[data-theme="${themeKey}"]`).classList.add('selected');
        
        // Show/hide custom style section
        const customSection = document.getElementById('customStyleSection');
        if (themeKey === '7') {
            customSection.classList.add('active');
        } else {
            customSection.classList.remove('active');
            this.unifier.selectedTheme = themeKey;
        }
    }

    async processDocument() {
        // Show progress section
        document.getElementById('progressSection').classList.add('active');
        document.getElementById('logSection').innerHTML = '';
        document.getElementById('processBtn').disabled = true;
        
        // Get selected theme
        const selectedCard = document.querySelector('.theme-card.selected');
        if (!selectedCard) {
            alert('Please select a theme first');
            return;
        }
        
        const themeKey = selectedCard.dataset.theme;
        const customStyle = document.getElementById('customStyleInput').value;
        
        // Process document
        try {
            const blob = await this.unifier.applyTheme(themeKey, customStyle);
            
            // Create download link
            const url = URL.createObjectURL(blob);
            const downloadBtn = document.getElementById('downloadBtn');
            downloadBtn.href = url;
            downloadBtn.download = 'styled_' + this.unifier.file.name;
            
            // Show download section
            document.getElementById('downloadSection').classList.add('active');
            
            this.unifier.addLog('✅ Processing complete!', 'success');
        } catch (error) {
            this.unifier.addLog(`❌ Error: ${error.message}`, 'error');
            console.error(error);
        }
    }
}

// Initialize app
document.addEventListener('DOMContentLoaded', () => {
    new UIController();
});