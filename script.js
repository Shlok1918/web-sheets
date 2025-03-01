class SpreadsheetApp {
    constructor() {
        this.rows = 30;
        this.columns = 20;
        this.selectedCell = { row: 0, col: 0 };
        this.cellData = {};
        this.formulaInput = document.getElementById('formula-input');
        this.cellPosition = document.getElementById('cell-position');
        this.gridContainer = document.getElementById('grid-container');
        this.columnHeaders = document.getElementById('column-headers');
        this.contextMenu = null;
        this.editMode = false;
        this.dragging = false;
        this.resizing = null;
        this.lastCalledCell = null;

        this.dragging = null;
        this.dragHighlight = null;
        this.handleDragMove = this.handleDragMove.bind(this);
        this.handleDragEnd = this.handleDragEnd.bind(this);

        this.chart = null;
        this.chartData = null;

        this.setupChartDrag();

        this.initGrid();
        this.initEventListeners();


    }

    initGrid() {

        const cornerElement = document.createElement('div');
        this.columnHeaders.appendChild(cornerElement);
        // Generate column headers
        for (let i = 0; i < this.columns; i++) {
            const columnHeader = document.createElement('div');
            columnHeader.className = 'column-header';
            columnHeader.textContent = this.getColumnName(i);

            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'column-resize-handle';
            resizeHandle.dataset.column = i;
            columnHeader.appendChild(resizeHandle);

            this.columnHeaders.appendChild(columnHeader);
        }

        // Generate rows and cells
        for (let i = 0; i < this.rows; i++) {
            const rowElement = document.createElement('div');
            rowElement.className = 'row';

            const rowHeader = document.createElement('div');
            rowHeader.className = 'row-header';
            rowHeader.textContent = i + 1;

            const resizeHandle = document.createElement('div');
            resizeHandle.className = 'row-resize-handle';
            resizeHandle.dataset.row = i;
            rowHeader.appendChild(resizeHandle);

            rowElement.appendChild(rowHeader);

            for (let j = 0; j < this.columns; j++) {
                const cell = document.createElement('div');
                cell.className = 'cell';
                cell.contentEditable = true;
                cell.dataset.row = i;
                cell.dataset.col = j;
                rowElement.appendChild(cell);
            }

            this.gridContainer.appendChild(rowElement);
        }

        // Set initial selection
        this.selectCell(0, 0);
    }

    initEventListeners() {
        // Cell selection and keyboard navigation
        document.addEventListener('click', (e) => {
            const cell = e.target.closest('.cell');
            if (cell) {
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                this.selectCell(row, col);
            } else if (this.contextMenu && !e.target.closest('.context-menu')) {
                this.removeContextMenu();
            }
        });

        // Right-click context menu
        document.addEventListener('contextmenu', (e) => {
            const cell = e.target.closest('.cell');
            if (cell) {
                e.preventDefault();
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                this.selectCell(row, col);
                this.showContextMenu(e.clientX, e.clientY);
            }
        });

        // Keyboard navigation
        document.addEventListener('keydown', (e) => {
            if (this.editMode && e.key !== 'Enter' && e.key !== 'Escape' && e.key !== 'Tab') {
                return;
            }

            const { row, col } = this.selectedCell;
            let newRow = row;
            let newCol = col;

            switch (e.key) {
                case 'ArrowUp':
                    e.preventDefault();
                    newRow = Math.max(0, row - 1);
                    break;
                case 'ArrowDown':
                    e.preventDefault();
                    newRow = Math.min(this.rows - 1, row + 1);
                    break;
                case 'ArrowLeft':
                    e.preventDefault();
                    newCol = Math.max(0, col - 1);
                    break;
                case 'ArrowRight':
                    e.preventDefault();
                    newCol = Math.min(this.columns - 1, col + 1);
                    break;
                case 'Tab':
                    e.preventDefault();
                    newCol = Math.min(this.columns - 1, col + 1);
                    if (newCol === col) {
                        newRow = Math.min(this.rows - 1, row + 1);
                        newCol = 0;
                    }
                    break;
                case 'Enter':
                    e.preventDefault();
                    if (this.editMode) {
                        this.exitEditMode();
                    } else {
                        newRow = Math.min(this.rows - 1, row + 1);
                    }
                    break;
                case 'Escape':
                    if (this.editMode) {
                        e.preventDefault();
                        this.exitEditMode(true); // Cancel edit
                    }
                    break;
                default:
                    if (!this.editMode && e.key.length === 1) {
                        this.enterEditMode();
                        return;
                    }
            }

            if (newRow !== row || newCol !== col) {
                this.selectCell(newRow, newCol);
            }
        });

        // Formula input
        this.formulaInput.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                this.updateCellValue(this.selectedCell.row, this.selectedCell.col, this.formulaInput.value);
                this.saveCell(this.selectedCell.row, this.selectedCell.col);
                this.selectCell(Math.min(this.rows - 1, this.selectedCell.row + 1), this.selectedCell.col);
            }
        });

        // Toolbar buttons
        document.getElementById('reset-btn').addEventListener('click', () => this.resetSheet());
        document.getElementById('bold-btn').addEventListener('click', () => this.toggleBold());
        document.getElementById('italic-btn').addEventListener('click', () => this.toggleItalic());
        document.getElementById('font-family').addEventListener('change', (e) => this.setFontFamily(e.target.value));
        document.getElementById('font-size').addEventListener('change', (e) => this.setFontSize(e.target.value));
        document.getElementById('font-color').addEventListener('input', (e) => this.setFontColor(e.target.value));
        document.getElementById('bg-color').addEventListener('input', (e) => this.setBackgroundColor(e.target.value));
        document.getElementById('add-row-btn').addEventListener('click', () => this.addRow());
        document.getElementById('add-column-btn').addEventListener('click', () => this.addColumn());

        // Cell content editing
        document.addEventListener('dblclick', (e) => {
            const cell = e.target.closest('.cell');
            if (cell) {
                this.enterEditMode();
            }
        });

        // Resize handling
        document.addEventListener('mousedown', (e) => {
            const colResizer = e.target.closest('.column-resize-handle');
            if (colResizer) {
                e.preventDefault();
                const column = parseInt(colResizer.dataset.column);
                this.startResize('column', column, e.clientX);
            }

            const rowResizer = e.target.closest('.row-resize-handle');
            if (rowResizer) {
                e.preventDefault();
                const row = parseInt(rowResizer.dataset.row);
                this.startResize('row', row, e.clientY);
            }
        });

        document.addEventListener('mousemove', (e) => {
            if (this.resizing) {
                e.preventDefault();
                this.updateResize(e.clientX, e.clientY);
            }
        });

        document.addEventListener('mouseup', () => {
            if (this.resizing) {
                this.stopResize();
            }
        });


        document.getElementById('export-csv').addEventListener('click', () => this.exportCSV());
        document.getElementById('import-csv-btn').addEventListener('click', () => document.getElementById('import-csv').click());
        document.getElementById('import-csv').addEventListener('change', (e) => this.importCSV(e));


        // Drag handle events
        document.addEventListener('mousedown', (e) => {
            const dragHandle = e.target.closest('.drag-handle');
            if (dragHandle) {
                e.preventDefault();
                const cell = dragHandle.parentElement;
                const row = parseInt(cell.dataset.row);
                const col = parseInt(cell.dataset.col);
                this.startDrag(row, col, e.clientX, e.clientY);
            }
        });


        document.getElementById('create-chart').addEventListener('click', () => this.showChartDialog());
        document.getElementById('close-chart').addEventListener('click', () => this.hideChartDialog());
        document.getElementById('generate-chart').addEventListener('click', () => this.createChart());

    }


    showChartDialog() {
        document.getElementById('chart-modal').style.display = 'block';
    }
    
    hideChartDialog() {
        document.getElementById('chart-modal').style.display = 'none';
        if (this.chart) {
            this.chart.destroy();
            this.chart = null;
        }
    }
    
    parseDataRange(range) {
        const [start, end] = range.split(':');
        const startCell = this.parseCellReference(start);
        const endCell = this.parseCellReference(end);
        
        const data = [];
        const labels = [];
        
        // Check if first row contains labels
        let hasLabels = false;
        const firstRow = startCell.row;
        for (let col = startCell.col; col <= endCell.col; col++) {
            const cell = this.cellData[`${firstRow},${col}`];
            if (cell && typeof cell.value !== 'number') {
                hasLabels = true;
                break;
            }
        }
    
        for (let row = startCell.row + (hasLabels ? 1 : 0); row <= endCell.row; row++) {
            const rowData = [];
            for (let col = startCell.col; col <= endCell.col; col++) {
                const cellId = `${row},${col}`;
                const cell = this.cellData[cellId];
                rowData.push(cell?.value || 0);
            }
            data.push(rowData);
        }
    
        if (hasLabels) {
            for (let col = startCell.col; col <= endCell.col; col++) {
                const cellId = `${startCell.row},${col}`;
                const cell = this.cellData[cellId];
                labels.push(cell?.value?.toString() || '');
            }
        }
    
        return { data, labels, hasLabels };
    }
    
    createChart() {
        const range = document.getElementById('data-range').value;
        const chartType = document.getElementById('chart-type').value;
        
        try {
            const { data, labels, hasLabels } = this.parseDataRange(range);
            
            // Create new chart window
            const chartWindow = document.createElement('div');
            chartWindow.className = 'chart-window';
            chartWindow.style.left = `${Math.random() * 200 + 50}px`;
            chartWindow.style.top = `${Math.random() * 200 + 50}px`;
            
            // Chart header
            const header = document.createElement('div');
            header.className = 'chart-header';
            header.innerHTML = `
                <h3>${chartType.charAt(0).toUpperCase() + chartType.slice(1)} Chart</h3>
                <button class="close-chart">&times;</button>
            `;
            
            // Chart content
            const canvas = document.createElement('canvas');
            canvas.className = 'chart-container';
            
            chartWindow.appendChild(header);
            chartWindow.appendChild(canvas);
            document.body.appendChild(chartWindow);
            
            // Add close handler
            header.querySelector('.close-chart').addEventListener('click', () => {
                chartWindow.remove();
            });
            
            // Initialize chart
            const ctx = canvas.getContext('2d');
            const config = this.getChartConfig(chartType, data, labels, hasLabels); // Now using the method
            new Chart(ctx, config);
            
            this.hideChartDialog();
        } catch (error) {
            alert('Invalid data range or format: ' + error.message);
        }
    }

    // Add this method to your SpreadsheetApp class
getChartConfig(chartType, data, labels, hasLabels) {
    const config = {
        type: chartType,
        options: {
            responsive: false,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'top',
                },
                title: {
                    display: true,
                    text: `${chartType.charAt(0).toUpperCase() + chartType.slice(1)} Chart`
                }
            }
        }
    };

    if (chartType === 'pie') {
        config.data = {
            labels: labels.length > 0 ? labels : data[0].map((_, i) => `Item ${i + 1}`),
            datasets: [{
                data: data[0],
                backgroundColor: this.generateChartColors(data[0].length)
            }]
        };
    } else {
        config.data = {
            labels: hasLabels ? labels : data.map((_, i) => i + 1),
            datasets: data.map((series, index) => ({
                label: hasLabels && labels[index] ? labels[index] : `Series ${index + 1}`,
                data: series,
                borderColor: this.generateChartColor(index),
                backgroundColor: this.generateChartColor(index, 0.2),
                fill: chartType === 'line'
            }))
        };
    }

    return config;
}
    
    generateChartColor(index, opacity = 1) {
        const colors = [
            `rgba(255, 99, 132, ${opacity})`,
            `rgba(54, 162, 235, ${opacity})`,
            `rgba(255, 206, 86, ${opacity})`,
            `rgba(75, 192, 192, ${opacity})`,
            `rgba(153, 102, 255, ${opacity})`,
            `rgba(255, 159, 64, ${opacity})`
        ];
        return colors[index % colors.length];
    }
    
    generateChartColors(count) {
        return Array.from({ length: count }, (_, i) => this.generateChartColor(i));
    }

    setupChartDrag() {
        let isDragging = false;
        let currentDragElement = null;
        let initialX = 0;
        let initialY = 0;
        let xOffset = 0;
        let yOffset = 0;
    
        document.addEventListener('mousedown', (e) => {
            const header = e.target.closest('.chart-header');
            if (header && header.parentElement.classList.contains('chart-window')) {
                currentDragElement = header.parentElement;
                isDragging = true;
                
                initialX = e.clientX - xOffset;
                initialY = e.clientY - yOffset;
                
                // Bring to front
                const maxZ = Math.max(...Array.from(document.querySelectorAll('.chart-window'))
                    .map(el => parseInt(window.getComputedStyle(el).zIndex) || 1000));
                currentDragElement.style.zIndex = maxZ + 1;
            }
        });
    
        document.addEventListener('mousemove', (e) => {
            if (isDragging && currentDragElement) {
                e.preventDefault();
                
                const currentX = e.clientX - initialX;
                const currentY = e.clientY - initialY;
                
                xOffset = currentX;
                yOffset = currentY;
                
                currentDragElement.style.left = `${currentX}px`;
                currentDragElement.style.top = `${currentY}px`;
            }
        });
    
        document.addEventListener('mouseup', () => {
            isDragging = false;
            currentDragElement = null;
        });
    }
    


    //Drag handle events
    startDrag(startRow, startCol, startX, startY) {
        this.dragging = {
            startRow,
            startCol,
            currentRow: startRow,
            currentCol: startCol,
            startX,
            startY,
            active: true
        };

        this.dragHighlight = document.createElement('div');
        this.dragHighlight.className = 'drag-highlight';
        this.gridContainer.appendChild(this.dragHighlight);

        document.addEventListener('mousemove', this.handleDragMove);
        document.addEventListener('mouseup', this.handleDragEnd);
    }

    handleDragMove(e) {
        if (!this.dragging?.active) return;

        const cell = this.getCell(0, 0);
        const cellWidth = cell.offsetWidth;
        const cellHeight = cell.offsetHeight;

        const deltaX = e.clientX - this.dragging.startX;
        const deltaY = e.clientY - this.dragging.startY;

        const deltaCol = Math.round(deltaX / cellWidth);
        const deltaRow = Math.round(deltaY / cellHeight);

        const newRow = Math.max(0, Math.min(this.rows - 1, this.dragging.startRow + deltaRow));
        const newCol = Math.max(0, Math.min(this.columns - 1, this.dragging.startCol + deltaCol));

        if (newRow !== this.dragging.currentRow || newCol !== this.dragging.currentCol) {
            this.dragging.currentRow = newRow;
            this.dragging.currentCol = newCol;
            this.updateDragHighlight();
        }
    }

    updateDragHighlight() {
        const { startRow, startCol, currentRow, currentCol } = this.dragging;
        const minRow = Math.min(startRow, currentRow);
        const maxRow = Math.max(startRow, currentRow);
        const minCol = Math.min(startCol, currentCol);
        const maxCol = Math.max(startCol, currentCol);

        const startCell = this.getCell(minRow, minCol);
        const endCell = this.getCell(maxRow, maxCol);

        const gridRect = this.gridContainer.getBoundingClientRect();
        const startRect = startCell.getBoundingClientRect();
        const endRect = endCell.getBoundingClientRect();

        this.dragHighlight.style.left = `${startRect.left - gridRect.left}px`;
        this.dragHighlight.style.top = `${startRect.top - gridRect.top}px`;
        this.dragHighlight.style.width = `${endRect.right - startRect.left}px`;
        this.dragHighlight.style.height = `${endRect.bottom - startRect.top}px`;
    }

    handleDragEnd() {
        if (!this.dragging?.active) return;

        const { startRow, startCol, currentRow, currentCol } = this.dragging;
        const minRow = Math.min(startRow, currentRow);
        const maxRow = Math.max(startRow, currentRow);
        const minCol = Math.min(startCol, currentCol);
        const maxCol = Math.max(startCol, currentCol);

        // Expand grid if needed
        while (this.rows <= maxRow) this.addRow();
        while (this.columns <= maxCol) this.addColumn();

        const sourceData = this.cellData[`${startRow},${startCol}`] || {};

        for (let row = minRow; row <= maxRow; row++) {
            for (let col = minCol; col <= maxCol; col++) {
                if (row === startRow && col === startCol) continue;

                const deltaRow = row - startRow;
                const deltaCol = col - startCol;
                const newData = { ...sourceData };

                if (newData.formula) {
                    newData.formula = this.adjustFormula(newData.formula, deltaRow, deltaCol);
                    try {
                        const result = this.evaluateFormula(
                            newData.formula.substring(1),
                            row,
                            col
                        );
                        newData.value = result.result;
                        newData.displayValue = result.result;
                        newData.references = result.references;
                    } catch (error) {
                        newData.value = '#ERROR!';
                        newData.displayValue = '#ERROR!';
                    }
                }

                this.cellData[`${row},${col}`] = newData;
                this.displayCellValue(this.getCell(row, col), newData);
            }
        }

        this.dragHighlight.remove();
        document.removeEventListener('mousemove', this.handleDragMove);
        document.removeEventListener('mouseup', this.handleDragEnd);
        this.dragging.active = false;
    }

    adjustFormula(formula, deltaRow, deltaCol) {
        return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (_, colAbs, col, rowAbs, row) => {
            let newCol = col;
            if (!colAbs) {
                const colIndex = this.getColumnIndex(col);
                newCol = this.getColumnName(colIndex + deltaCol);
            }

            let newRow = parseInt(row);
            if (!rowAbs) {
                newRow += deltaRow;
            }

            return `${colAbs}${newCol}${rowAbs}${newRow}`;
        });
    }

    exportCSV() {
        const csvContent = [];
        for (let row = 0; row < this.rows; row++) {
            const rowData = [];
            for (let col = 0; col < this.columns; col++) {
                const cellId = `${row},${col}`;
                const cell = this.cellData[cellId];
                let value = '';
                if (cell) {
                    value = cell.displayValue !== undefined ? cell.displayValue : cell.value || '';
                    if (typeof value === 'string') {
                        if (/[,"\n]/.test(value)) {
                            value = `"${value.replace(/"/g, '""')}"`;
                        }
                    } else {
                        value = String(value);
                    }
                }
                rowData.push(value);
            }
            csvContent.push(rowData.join(','));
        }

        const csvString = csvContent.join('\r\n');
        const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
        const link = document.createElement('a');
        link.href = URL.createObjectURL(blob);
        link.download = 'spreadsheet.csv';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    importCSV(event) {
        const file = event.target.files[0];
        if (!file) return;

        const reader = new FileReader();
        reader.onload = (e) => {
            const csvText = e.target.result;
            const csvData = this.parseCsv(csvText);

            this.resetSheet();

            // Determine needed rows/columns
            const csvRows = csvData.length;
            const csvCols = csvData.reduce((max, row) => Math.max(max, row.length), 0);

            // Add required rows
            while (this.rows < csvRows) {
                this.addRow();
            }

            // Add required columns
            while (this.columns < csvCols) {
                this.addColumn();
            }

            // Populate cells
            csvData.forEach((row, rowIndex) => {
                row.forEach((value, colIndex) => {
                    this.updateCellValue(rowIndex, colIndex, value);
                });
            });

            event.target.value = '';
        };
        reader.readAsText(file);
    }

    parseCsv(csvText) {
        const rows = [];
        let currentRow = [];
        let currentField = '';
        let inQuotes = false;

        for (let i = 0; i < csvText.length; i++) {
            const char = csvText[i];

            if (char === '"') {
                if (inQuotes && csvText[i + 1] === '"') {
                    currentField += '"';
                    i++;
                } else {
                    inQuotes = !inQuotes;
                }
            } else if (char === ',' && !inQuotes) {
                currentRow.push(currentField);
                currentField = '';
            } else if ((char === '\n' || char === '\r') && !inQuotes) {
                currentRow.push(currentField);
                rows.push(currentRow);
                currentRow = [];
                currentField = '';
                if (char === '\r' && csvText[i + 1] === '\n') i++;
            } else {
                currentField += char;
            }
        }

        // Handle last field/row
        if (currentField !== '' || currentRow.length > 0) {
            currentRow.push(currentField);
            rows.push(currentRow);
        }

        return rows;
    }
    startResize(type, index, initialPos) {
        this.resizing = { type, index, initialPos };
        document.body.style.cursor = type === 'column' ? 'col-resize' : 'row-resize';
    }

    updateResize(clientX, clientY) {
        if (!this.resizing) return;

        const { type, index, initialPos } = this.resizing;

        if (type === 'column') {
            const delta = clientX - initialPos;
            const cells = document.querySelectorAll(`.cell[data-col="${index}"]`);
            const header = document.querySelectorAll('.column-header')[index];
            const currentWidth = cells[0].offsetWidth;
            const newWidth = Math.max(50, currentWidth + delta);

            cells.forEach(cell => {
                cell.style.width = `${newWidth}px`;
                cell.style.minWidth = `${newWidth}px`;
            });

            if (header) {
                header.style.width = `${newWidth}px`;
                header.style.minWidth = `${newWidth}px`;
            }

            this.resizing.initialPos = clientX;
        } else if (type === 'row') {
            const delta = clientY - initialPos;
            const cells = document.querySelectorAll(`.cell[data-row="${index}"]`);
            const row = document.querySelectorAll('.row')[index];
            const header = row.querySelector('.row-header');
            const currentHeight = cells[0].offsetHeight;
            const newHeight = Math.max(20, currentHeight + delta);

            cells.forEach(cell => {
                cell.style.height = `${newHeight}px`;
                cell.style.minHeight = `${newHeight}px`;
            });

            if (header) {
                header.style.height = `${newHeight}px`;
                header.style.minHeight = `${newHeight}px`;
            }

            this.resizing.initialPos = clientY;
        }
    }

    stopResize() {
        this.resizing = null;
        document.body.style.cursor = '';
    }

    selectCell(row, col) {
        if (this.editMode) {
            this.exitEditMode();
        }

        // Remove previous selection
        const prevSelected = document.querySelector('.cell.selected');
        if (prevSelected) {
            prevSelected.classList.remove('selected');
        }

        this.selectedCell = { row, col };
        const cell = this.getCell(row, col);
        cell.classList.add('selected');

        // Update cell position indicator
        this.cellPosition.textContent = `${this.getColumnName(col)}${row + 1}`;

        // Update formula bar
        const cellId = `${row},${col}`;
        const cellData = this.cellData[cellId] || {};
        this.formulaInput.value = cellData.formula || cellData.value || '';

        // Scroll cell into view if needed
        cell.scrollIntoView({ block: 'nearest', inline: 'nearest' });


        const dragHandle = document.createElement('div');
        dragHandle.className = 'drag-handle';
        cell.appendChild(dragHandle);
    }

    enterEditMode() {
        const { row, col } = this.selectedCell;
        const cell = this.getCell(row, col);
        const cellId = `${row},${col}`;
        const cellData = this.cellData[cellId] || {};

        // Show formula in cell when editing
        cell.textContent = cellData.formula || cellData.value || '';
        this.editMode = true;
        cell.focus();

        // Select all text in the cell
        const selection = window.getSelection();
        const range = document.createRange();
        range.selectNodeContents(cell);
        selection.removeAllRanges();
        selection.addRange(range);
    }

    exitEditMode(cancel = false) {
        const { row, col } = this.selectedCell;
        const cell = this.getCell(row, col);

        if (!cancel) {
            const value = cell.textContent.trim();
            this.updateCellValue(row, col, value);
            this.saveCell(row, col);
        } else {
            // Restore the cell's previous value
            const cellId = `${row},${col}`;
            const cellData = this.cellData[cellId] || {};
            this.displayCellValue(cell, cellData);
        }

        this.editMode = false;
        cell.blur();
    }

    updateCellValue(row, col, value) {
        const cell = this.getCell(row, col);
        const cellId = `${row},${col}`;

        // Store old value and references for detecting circular references
        const oldCellData = this.cellData[cellId] || {};
        const oldReferences = oldCellData.references || [];

        // Initialize new cell data
        let cellData = {
            value: value,
            displayValue: value,
            type: 'string',
            references: [],
            formula: null,
            styling: oldCellData.styling || {}
        };

        // Handle formulas
        if (value && value.startsWith('=')) {
            cellData.formula = value;

            try {
                const { result, references } = this.evaluateFormula(value.substring(1), row, col);
                cellData.value = result;
                cellData.displayValue = result;
                cellData.type = typeof result;
                cellData.references = references;

                // Update formula bar with the formula
                this.formulaInput.value = value;
            } catch (error) {
                cellData.value = '#ERROR!';
                cellData.displayValue = '#ERROR!';
                cellData.type = 'error';
                console.error('Formula error:', error);
            }
        } else {
            // Try to detect types for non-formula values
            if (!isNaN(value) && value !== '') {
                cellData.value = parseFloat(value);
                cellData.type = 'number';
            }
        }

        // Save cell data
        this.cellData[cellId] = cellData;

        // Remove this cell from old referenced cells' dependents
        oldReferences.forEach(ref => {
            // Update cells that were referencing this cell
            const referringCells = this.findCellsThatReference(ref);
            referringCells.forEach(refCell => {
                if (refCell !== cellId) {
                    this.recalculateCell(refCell);
                }
            });
        });

        // Update cells that reference this cell
        const referringCells = this.findCellsThatReference(cellId);
        referringCells.forEach(refCell => {
            this.recalculateCell(refCell);
        });

        // Display the value in the cell
        this.displayCellValue(cell, cellData);
    }

    findCellsThatReference(cellId) {
        const referringCells = [];

        Object.keys(this.cellData).forEach(key => {
            if (this.cellData[key].references && this.cellData[key].references.includes(cellId)) {
                referringCells.push(key);
            }
        });

        return referringCells;
    }

    recalculateCell(cellId) {
        const [row, col] = cellId.split(',').map(Number);
        const cellData = this.cellData[cellId];

        if (cellData && cellData.formula) {
            this.updateCellValue(row, col, cellData.formula);
        }
    }

    saveCell(row, col) {
        const cellId = `${row},${col}`;
        const cellData = this.cellData[cellId] || {};

        // Update formula bar with the formula if it exists, otherwise the value
        this.formulaInput.value = cellData.formula || cellData.value || '';
    }

    displayCellValue(cell, cellData) {
        // Clear existing content and apply styling
        cell.textContent = cellData.displayValue !== undefined ? cellData.displayValue : '';

        // Apply styling if available
        const styling = cellData.styling || {};
        cell.style.fontWeight = styling.bold ? 'bold' : 'normal';
        cell.style.fontStyle = styling.italic ? 'italic' : 'normal';
        if (styling.fontFamily) cell.style.fontFamily = styling.fontFamily;
        if (styling.fontSize) cell.style.fontSize = styling.fontSize + 'px';
        if (styling.color) cell.style.color = styling.color;
        if (styling.backgroundColor) cell.style.backgroundColor = styling.backgroundColor;
    }

    evaluateFormula(formula, currentRow, currentCol) {
        // Reset call stack detection
        this.lastCalledCell = null;
        let references = [];

        // Define regex patterns for various functions
        const sumPattern = /SUM\(([A-Z]+[0-9]+:[A-Z]+[0-9]+)\)/gi;
        const avgPattern = /AVERAGE\(([A-Z]+[0-9]+:[A-Z]+[0-9]+)\)/gi;
        const maxPattern = /MAX\(([A-Z]+[0-9]+:[A-Z]+[0-9]+)\)/gi;
        const minPattern = /MIN\(([A-Z]+[0-9]+:[A-Z]+[0-9]+)\)/gi;
        const countPattern = /COUNT\(([A-Z]+[0-9]+:[A-Z]+[0-9]+)\)/gi;
        const trimPattern = /TRIM\(([A-Z]+[0-9]+)\)/gi;
        const upperPattern = /UPPER\(([A-Z]+[0-9]+)\)/gi;
        const lowerPattern = /LOWER\(([A-Z]+[0-9]+)\)/gi;

        // Process SUM function
        formula = formula.replace(sumPattern, (match, range) => {
            const { values, refs } = this.getCellsInRange(range, currentRow, currentCol);
            references = [...references, ...refs];
            return values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0);
        });

        // Process AVERAGE function
        formula = formula.replace(avgPattern, (match, range) => {
            const { values, refs } = this.getCellsInRange(range, currentRow, currentCol);
            references = [...references, ...refs];
            const numValues = values.filter(val => !isNaN(parseFloat(val))).length;
            return numValues > 0
                ? values.reduce((sum, val) => sum + (parseFloat(val) || 0), 0) / numValues
                : 0;
        });

        // Process MAX function
        formula = formula.replace(maxPattern, (match, range) => {
            const { values, refs } = this.getCellsInRange(range, currentRow, currentCol);
            references = [...references, ...refs];
            const numValues = values.filter(val => !isNaN(parseFloat(val)));
            return numValues.length > 0
                ? Math.max(...numValues.map(val => parseFloat(val)))
                : 0;
        });

        // Process MIN function
        formula = formula.replace(minPattern, (match, range) => {
            const { values, refs } = this.getCellsInRange(range, currentRow, currentCol);
            references = [...references, ...refs];
            const numValues = values.filter(val => !isNaN(parseFloat(val)));
            return numValues.length > 0
                ? Math.min(...numValues.map(val => parseFloat(val)))
                : 0;
        });

        // Process COUNT function
        formula = formula.replace(countPattern, (match, range) => {
            const { values, refs } = this.getCellsInRange(range, currentRow, currentCol);
            references = [...references, ...refs];
            return values.filter(val => !isNaN(parseFloat(val))).length;
        });

        // Process TRIM function
        formula = formula.replace(trimPattern, (match, cellRef) => {
            const { row, col } = this.parseCellReference(cellRef);
            const cellId = `${row},${col}`;
            references.push(cellId);
            const cellData = this.cellData[cellId] || {};
            return `"${String(cellData.value || '').trim()}"`;
        });

        // Process UPPER function
        formula = formula.replace(upperPattern, (match, cellRef) => {
            const { row, col } = this.parseCellReference(cellRef);
            const cellId = `${row},${col}`;
            references.push(cellId);
            const cellData = this.cellData[cellId] || {};
            return `"${String(cellData.value || '').toUpperCase()}"`;
        });

        // Process LOWER function
        formula = formula.replace(lowerPattern, (match, cellRef) => {
            const { row, col } = this.parseCellReference(cellRef);
            const cellId = `${row},${col}`;
            references.push(cellId);
            const cellData = this.cellData[cellId] || {};
            return `"${String(cellData.value || '').toLowerCase()}"`;
        });

        const cellRefPattern = /([A-Z]+)([0-9]+)/g;
        formula = formula.replace(cellRefPattern, (match, colName, rowNum) => {
            const col = this.getColumnIndex(colName);
            const row = parseInt(rowNum) - 1;
            const cellId = `${row},${col}`;

            // Check for circular reference
            if (row === currentRow && col === currentCol) {
                throw new Error('Circular reference detected');
            }

            // Add to references
            references.push(cellId);

            // Get cell value
            const cellData = this.cellData[cellId] || {};
            if (cellData.type === 'string') {
                return `"${cellData.value || ''}"`;
            }
            return cellData.value || 0;
        });

        // Evaluate the formula
        try {
            // Use Function constructor to safely evaluate the formula
            const result = new Function(`return ${formula}`)();
            return { result, references };
        } catch (error) {
            console.error('Error evaluating formula:', error);
            throw new Error('Invalid formula');
        }
    }

    getCellsInRange(range, currentRow, currentCol) {
        const [start, end] = range.split(':');
        const startCell = this.parseCellReference(start);
        const endCell = this.parseCellReference(end);

        const values = [];
        const refs = [];

        for (let row = startCell.row; row <= endCell.row; row++) {
            for (let col = startCell.col; col <= endCell.col; col++) {
                // Check for circular reference
                if (row === currentRow && col === currentCol) {
                    throw new Error('Circular reference detected');
                }

                const cellId = `${row},${col}`;
                refs.push(cellId);

                const cellData = this.cellData[cellId] || {};
                values.push(cellData.value || 0);
            }
        }

        return { values, refs };
    }

    parseCellReference(ref) {
        const match = ref.match(/([A-Z]+)([0-9]+)/);
        if (!match) {
            throw new Error(`Invalid cell reference: ${ref}`);
        }

        const colName = match[1];
        const rowNum = parseInt(match[2]);

        return {
            row: rowNum - 1,
            col: this.getColumnIndex(colName)
        };
    }

    getColumnIndex(colName) {
        let col = 0;
        for (let i = 0; i < colName.length; i++) {
            col = col * 26 + (colName.charCodeAt(i) - 64);
        }
        return col - 1;
    }

    getColumnName(index) {
        let columnName = '';
        let temp = index + 1;

        while (temp > 0) {
            const remainder = (temp - 1) % 26;
            columnName = String.fromCharCode(65 + remainder) + columnName;
            temp = Math.floor((temp - remainder) / 26);
        }

        return columnName;
    }

    getCell(row, col) {
        return document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
    }

    // Styling functions
    toggleBold() {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.bold = !cellData.styling.bold;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.fontWeight = cellData.styling.bold ? 'bold' : 'normal';
    }

    toggleItalic() {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.italic = !cellData.styling.italic;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.fontStyle = cellData.styling.italic ? 'italic' : 'normal';
    }

    setFontFamily(fontFamily) {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.fontFamily = fontFamily;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.fontFamily = fontFamily;
    }

    setFontSize(fontSize) {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.fontSize = fontSize;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.fontSize = `${fontSize}px`;
    }

    setFontColor(color) {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.color = color;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.color = color;
    }

    setBackgroundColor(color) {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        let cellData = this.cellData[cellId] || {};

        if (!cellData.styling) cellData.styling = {};
        cellData.styling.backgroundColor = color;

        this.cellData[cellId] = cellData;

        const cell = this.getCell(row, col);
        cell.style.backgroundColor = color;
    }

    showContextMenu(x, y) {
        this.removeContextMenu();

        const menu = document.createElement('div');
        menu.className = 'context-menu';

        const menuItems = [
            { text: 'Cut', action: () => this.cutCell() },
            { text: 'Copy', action: () => this.copyCell() },
            { text: 'Paste', action: () => this.pasteCell() },
            { text: 'Delete', action: () => this.clearCell() },
            { text: 'Insert Row Above', action: () => this.insertRowAbove() },
            { text: 'Insert Row Below', action: () => this.insertRowBelow() },
            { text: 'Insert Column Left', action: () => this.insertColumnLeft() },
            { text: 'Insert Column Right', action: () => this.insertColumnRight() },
            { text: 'Delete Row', action: () => this.deleteRow() },
            { text: 'Delete Column', action: () => this.deleteColumn() }
        ];

        menuItems.forEach(item => {
            const menuItem = document.createElement('div');
            menuItem.className = 'context-menu-item';
            menuItem.textContent = item.text;
            menuItem.addEventListener('click', () => {
                item.action();
                this.removeContextMenu();
            });
            menu.appendChild(menuItem);
        });

        // Position the menu
        menu.style.left = `${x}px`;
        menu.style.top = `${y}px`;

        // Adjust menu position if it goes out of viewport
        const viewportWidth = window.innerWidth;
        const viewportHeight = window.innerHeight;

        setTimeout(() => {
            const menuWidth = menu.offsetWidth;
            const menuHeight = menu.offsetHeight;

            if (x + menuWidth > viewportWidth) {
                menu.style.left = `${x - menuWidth}px`;
            }

            if (y + menuHeight > viewportHeight) {
                menu.style.top = `${y - menuHeight}px`;
            }
        }, 0);

        document.body.appendChild(menu);
        this.contextMenu = menu;
    }

    removeContextMenu() {
        if (this.contextMenu) {
            this.contextMenu.remove();
            this.contextMenu = null;
        }
    }

    // Clipboard operations
    copyCell() {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;
        localStorage.setItem('clipboard', JSON.stringify(this.cellData[cellId] || {}));
    }

    cutCell() {
        this.copyCell();
        this.clearCell();
    }

    pasteCell() {
        const clipboard = JSON.parse(localStorage.getItem('clipboard') || '{}');
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;

        // Update cell data with clipboard content
        this.cellData[cellId] = JSON.parse(JSON.stringify(clipboard)); // Deep copy

        // If the clipboard contains a formula, update it to point to the new cell
        if (this.cellData[cellId].formula) {
            // Keep the formula as is for now - user can edit if needed
        }

        // Update cell display
        const cell = this.getCell(row, col);
        this.displayCellValue(cell, this.cellData[cellId]);
        this.formulaInput.value = this.cellData[cellId].formula || this.cellData[cellId].value || '';
    }

    clearCell() {
        const { row, col } = this.selectedCell;
        const cellId = `${row},${col}`;

        // Clear cell data
        delete this.cellData[cellId];

        // Clear cell display
        const cell = this.getCell(row, col);
        cell.textContent = '';
        cell.style = '';

        // Clear formula bar
        this.formulaInput.value = '';

        // Update cells that reference this cell
        const referringCells = this.findCellsThatReference(cellId);
        referringCells.forEach(refCell => {
            this.recalculateCell(refCell);
        });
    }

    // Row and column operations
    addRow() {
        this.insertRowBelow();
    }

    addColumn() {
        this.insertColumnRight();
    }

    insertRowAbove() {
        const { row } = this.selectedCell;
        this.insertRow(row);
    }

    insertRowBelow() {
        const { row } = this.selectedCell;
        this.insertRow(row + 1);
    }

    insertRow(rowIndex) {
        // Update data structure for the new row
        const newCellData = {};

        // Shift existing data down
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            if (r >= rowIndex) {
                newCellData[`${r + 1},${c}`] = this.cellData[key];
            } else {
                newCellData[key] = this.cellData[key];
            }
        });

        this.cellData = newCellData;

        // Update UI
        const rowElement = document.createElement('div');
        rowElement.className = 'row';

        const rowHeader = document.createElement('div');
        rowHeader.className = 'row-header';

        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'row-resize-handle';
        resizeHandle.dataset.row = this.rows;
        rowHeader.appendChild(resizeHandle);

        rowElement.appendChild(rowHeader);

        for (let j = 0; j < this.columns; j++) {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.contentEditable = true;
            cell.dataset.row = this.rows;
            cell.dataset.col = j;
            rowElement.appendChild(cell);
        }

        this.rows++;

        // Insert the new row at the specified position
        const rows = document.querySelectorAll('.row');
        if (rowIndex < rows.length) {
            const referenceRow = rows[rowIndex];
            referenceRow.parentNode.insertBefore(rowElement, referenceRow);
        } else {
            this.gridContainer.appendChild(rowElement);
        }

        // Update row numbers
        const rowHeaders = document.querySelectorAll('.row-header');
        for (let i = 0; i < rowHeaders.length; i++) {
            rowHeaders[i].textContent = i + 1;
            const resizeHandle = rowHeaders[i].querySelector('.row-resize-handle');
            if (resizeHandle) {
                resizeHandle.dataset.row = i;
            }
        }

        // Update row data attributes
        const allRows = document.querySelectorAll('.row');
        for (let i = 0; i < allRows.length; i++) {
            const cells = allRows[i].querySelectorAll('.cell');
            cells.forEach(cell => {
                cell.dataset.row = i;
            });
        }

        // Update formulas that might reference cells in the shifted rows
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            const cellData = this.cellData[key];

            if (cellData.formula) {
                this.updateCellValue(r, c, cellData.formula);
            }
        });
    }

    insertColumnLeft() {
        const { col } = this.selectedCell;
        this.insertColumn(col);
    }

    insertColumnRight() {
        const { col } = this.selectedCell;
        this.insertColumn(col + 1);
    }

    insertColumn(colIndex) {
        // Update data structure for the new column
        const newCellData = {};

        // Shift existing data right
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            if (c >= colIndex) {
                newCellData[`${r},${c + 1}`] = this.cellData[key];
            } else {
                newCellData[key] = this.cellData[key];
            }
        });

        this.cellData = newCellData;

        // Add new column header
        const columnHeader = document.createElement('div');
        columnHeader.className = 'column-header';

        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'column-resize-handle';
        resizeHandle.dataset.column = this.columns;
        columnHeader.appendChild(resizeHandle);

        const headers = document.querySelectorAll('.column-header');
        if (colIndex < headers.length) {
            const referenceHeader = headers[colIndex];
            this.columnHeaders.insertBefore(columnHeader, referenceHeader);
        } else {
            this.columnHeaders.appendChild(columnHeader);
        }

        // Add new cell to each row
        const rows = document.querySelectorAll('.row');
        rows.forEach((row, rowIndex) => {
            const cell = document.createElement('div');
            cell.className = 'cell';
            cell.contentEditable = true;
            cell.dataset.row = rowIndex;
            cell.dataset.col = this.columns;

            const cells = row.querySelectorAll('.cell');
            if (colIndex < cells.length) {
                const referenceCell = cells[colIndex];
                row.insertBefore(cell, referenceCell);
            } else {
                row.appendChild(cell);
            }
        });

        this.columns++;

        // Update column letters
        const columnHeaders = document.querySelectorAll('.column-header');
        for (let i = 0; i < columnHeaders.length; i++) {
            if (i > 0) { // Skip the corner cell
                columnHeaders[i].textContent = this.getColumnName(i - 1);
            }
            const resizeHandle = columnHeaders[i].querySelector('.column-resize-handle');
            if (resizeHandle) {
                resizeHandle.dataset.column = i - 1;
            }
        }

        // Update column data attributes
        const allRows = document.querySelectorAll('.row');
        for (let i = 0; i < allRows.length; i++) {
            const cells = allRows[i].querySelectorAll('.cell');
            for (let j = 0; j < cells.length; j++) {
                cells[j].dataset.col = j;
            }
        }

        // Update formulas that might reference cells in the shifted columns
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            const cellData = this.cellData[key];

            if (cellData.formula) {
                this.updateCellValue(r, c, cellData.formula);
            }
        });
    }

    deleteRow() {
        const { row } = this.selectedCell;

        // Update data structure
        const newCellData = {};

        // Remove data for deleted row and shift data up
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            if (r < row) {
                newCellData[key] = this.cellData[key];
            } else if (r > row) {
                newCellData[`${r - 1},${c}`] = this.cellData[key];
            }
            // Row === rowIndex data is simply dropped
        });

        this.cellData = newCellData;

        // Update UI
        const rows = document.querySelectorAll('.row');
        if (row < rows.length) {
            rows[row].remove();
        }

        this.rows--;

        // Update row numbers
        const rowHeaders = document.querySelectorAll('.row-header');
        for (let i = 0; i < rowHeaders.length; i++) {
            rowHeaders[i].textContent = i + 1;
            const resizeHandle = rowHeaders[i].querySelector('.row-resize-handle');
            if (resizeHandle) {
                resizeHandle.dataset.row = i;
            }
        }
        
        const allRows = document.querySelectorAll('.row');
        for (let i = 0; i < allRows.length; i++) {
            const cells = allRows[i].querySelectorAll('.cell');
            cells.forEach(cell => {
                cell.dataset.row = i;
            });
        }

        this.selectCell(Math.min(row, this.rows - 1), this.selectedCell.col);
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            const cellData = this.cellData[key];

            if (cellData.formula) {
                this.updateCellValue(r, c, cellData.formula);
            }
        });
    }

    deleteColumn() {
        const { col } = this.selectedCell;

        const newCellData = {};
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            if (c < col) {
                newCellData[key] = this.cellData[key];
            } else if (c > col) {
                newCellData[`${r},${c - 1}`] = this.cellData[key];
            }
        });

        this.cellData = newCellData;

        const headers = document.querySelectorAll('.column-header');
        if (col + 1 < headers.length) { 
            headers[col + 1].remove();
        }

        // Remove column cells
        const rows = document.querySelectorAll('.row');
        rows.forEach(row => {
            const cells = row.querySelectorAll('.cell');
            if (col < cells.length) {
                cells[col].remove();
            }
        });

        this.columns--;

        // Update column letters
        const columnHeaders = document.querySelectorAll('.column-header');
        for (let i = 0; i < columnHeaders.length; i++) {
            if (i > 0) { // Skip the corner cell
                columnHeaders[i].textContent = this.getColumnName(i - 1);
            }
            const resizeHandle = columnHeaders[i].querySelector('.column-resize-handle');
            if (resizeHandle) {
                resizeHandle.dataset.column = i - 1;
            }
        }

        // Update column data attributes
        const allRows = document.querySelectorAll('.row');
        for (let i = 0; i < allRows.length; i++) {
            const cells = allRows[i].querySelectorAll('.cell');
            for (let j = 0; j < cells.length; j++) {
                cells[j].dataset.col = j;
            }
        }

        this.selectCell(this.selectedCell.row, Math.min(col, this.columns - 1));
        Object.keys(this.cellData).forEach(key => {
            const [r, c] = key.split(',').map(Number);
            const cellData = this.cellData[key];

            if (cellData.formula) {
                this.updateCellValue(r, c, cellData.formula);
            }
        });
    }

    resetSheet() {
        this.cellData = {};
        const cells = document.querySelectorAll('.cell');
        cells.forEach(cell => {
            cell.textContent = '';
            cell.style = '';
        });
        this.formulaInput.value = '';
        this.selectCell(0, 0);
    }
}


// Initialize the spreadsheet when the document is fully loaded
document.addEventListener('DOMContentLoaded', () => {
    window.spreadsheetApp = new SpreadsheetApp();
});