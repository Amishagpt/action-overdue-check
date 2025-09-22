class ExcelAnalyzer {
    constructor() {
        this.dragActive = false;
        this.isAnalyzing = false;
        this.result = null;
        this.initializeEventListeners();
    }

    initializeEventListeners() {
        const uploadZone = document.getElementById('uploadZone');
        const fileInput = document.getElementById('fileInput');

        // Drag and drop events
        uploadZone.addEventListener('dragenter', this.handleDrag.bind(this));
        uploadZone.addEventListener('dragover', this.handleDrag.bind(this));
        uploadZone.addEventListener('dragleave', this.handleDrag.bind(this));
        uploadZone.addEventListener('drop', this.handleDrop.bind(this));

        // File input change
        fileInput.addEventListener('change', this.handleFileInput.bind(this));
    }

    handleDrag(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const uploadZone = document.getElementById('uploadZone');
        
        if (e.type === 'dragenter' || e.type === 'dragover') {
            if (!this.dragActive) {
                this.dragActive = true;
                uploadZone.classList.add('drag-active');
            }
        } else if (e.type === 'dragleave') {
            this.dragActive = false;
            uploadZone.classList.remove('drag-active');
        }
    }

    handleDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        
        const uploadZone = document.getElementById('uploadZone');
        this.dragActive = false;
        uploadZone.classList.remove('drag-active');

        const files = Array.from(e.dataTransfer.files);
        const excelFile = files.find(file => 
            file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
        );

        if (excelFile) {
            this.analyzeExcel(excelFile);
        } else {
            this.showToast('Please upload an Excel file (.xlsx or .xls)', 'error');
        }
    }

    handleFileInput(e) {
        const file = e.target.files[0];
        if (file) {
            this.analyzeExcel(file);
        }
    }

    isAssigned(value) {
        if (value === null || value === undefined || value === '') return false;
        const str = String(value).toLowerCase().trim();
        return ['yes', 'true', 'assigned', 'done', '1'].includes(str) || 
               (str !== 'no' && str !== 'false' && str !== 'unassigned' && str !== '0' && str !== '');
    }

    parseDate(dateValue) {
        if (!dateValue) return null;
        
        // Handle Excel serial dates
        if (typeof dateValue === 'number') {
            const excelEpoch = new Date(1900, 0, 1);
            const days = dateValue - 2; // Excel starts from 1900-01-01, but has leap year bug
            return new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000);
        }
        
        // Handle date strings
        const parsed = new Date(dateValue);
        return isNaN(parsed.getTime()) ? null : parsed;
    }

    getTodayInKolkata() {
        const now = new Date();
        const kolkataTime = new Date(now.toLocaleString("en-US", { timeZone: "Asia/Kolkata" }));
        return new Date(kolkataTime.getFullYear(), kolkataTime.getMonth(), kolkataTime.getDate());
    }

    findColumns(worksheet) {
        const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
        let actionCol = null;
        let dueDateCol = null;

        // Check first row for headers
        for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
            const cell = worksheet[cellAddress];
            if (!cell) continue;

            const header = String(cell.w || cell.v || '').toLowerCase().trim();
            if (header.includes('action') && !actionCol) {
                actionCol = XLSX.utils.encode_col(col);
            } else if ((header.includes('due') && header.includes('date')) || header === 'due date' && !dueDateCol) {
                dueDateCol = XLSX.utils.encode_col(col);
            }
        }

        return { actionCol, dueDateCol };
    }

    setAnalyzingState(analyzing) {
        this.isAnalyzing = analyzing;
        const uploadIcon = document.getElementById('uploadIcon');
        const uploadTitle = document.getElementById('uploadTitle');
        const fileInput = document.getElementById('fileInput');
        const browseBtn = document.querySelector('.browse-btn');

        if (analyzing) {
            uploadIcon.innerHTML = '<div class="spinner"></div>';
            uploadIcon.classList.add('spinning');
            uploadTitle.textContent = 'Analyzing...';
            fileInput.disabled = true;
            browseBtn.disabled = true;
        } else {
            uploadIcon.innerHTML = `
                <svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
                    <polyline points="7,10 12,15 17,10"></polyline>
                    <line x1="12" y1="15" x2="12" y2="3"></line>
                </svg>
            `;
            uploadIcon.classList.remove('spinning');
            uploadTitle.textContent = 'Drop Excel file here';
            fileInput.disabled = false;
            browseBtn.disabled = false;
        }
    }

    showError(message) {
        const errorCard = document.getElementById('errorCard');
        const errorMessage = document.getElementById('errorMessage');
        
        errorMessage.textContent = message;
        errorCard.style.display = 'block';
        
        // Hide results if shown
        document.getElementById('resultsSection').style.display = 'none';
    }

    hideError() {
        document.getElementById('errorCard').style.display = 'none';
    }

    showResults(result) {
        this.hideError();
        
        const resultsSection = document.getElementById('resultsSection');
        
        // Update summary
        document.getElementById('summaryText').textContent = result.summary;
        
        // Update stats
        document.getElementById('totalCount').textContent = result.total_rows;
        document.getElementById('assignedCount').textContent = result.assigned_count;
        document.getElementById('assignedPct').textContent = `${result.assigned_pct}% of total`;
        document.getElementById('overdueCount').textContent = result.overdue_count;
        document.getElementById('overduePct').textContent = `${result.overdue_pct_of_assigned}% of assigned`;
        
        // Update details
        document.getElementById('analysisDate').textContent = result.today_iso;
        document.getElementById('timezone').textContent = result.timezone;
        document.getElementById('actionColumn').textContent = result.columns_used.action;
        document.getElementById('dueDateColumn').textContent = result.columns_used.due_date;
        
        // Update notes
        const notesSection = document.getElementById('notesSection');
        const notesList = document.getElementById('notesList');
        
        if (result.notes && result.notes.length > 0) {
            notesList.innerHTML = '';
            result.notes.forEach(note => {
                const li = document.createElement('li');
                li.textContent = note;
                notesList.appendChild(li);
            });
            notesSection.style.display = 'block';
        } else {
            notesSection.style.display = 'none';
        }
        
        // Update JSON output
        document.getElementById('jsonOutput').textContent = JSON.stringify(result, null, 2);
        
        resultsSection.style.display = 'block';
    }

    showToast(message, type = 'success') {
        const toast = document.getElementById('toast');
        const toastMessage = document.getElementById('toastMessage');
        
        toastMessage.textContent = message;
        toast.className = `toast ${type}`;
        toast.classList.add('show');
        
        setTimeout(() => {
            toast.classList.remove('show');
        }, 3000);
    }

    async analyzeExcel(file) {
        this.setAnalyzingState(true);
        this.hideError();

        try {
            const buffer = await file.arrayBuffer();
            const workbook = XLSX.read(buffer, { type: 'buffer' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];
            const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

            const { actionCol, dueDateCol } = this.findColumns(worksheet);
            
            if (!actionCol) {
                throw new Error('Action column not found. Please ensure your Excel file has an "Action" column.');
            }

            const notes = [];
            if (!dueDateCol) {
                notes.push('Due Date column not found. Overdue analysis skipped.');
            }

            const today = this.getTodayInKolkata();
            let totalRows = 0;
            let assignedCount = 0;
            let overdueCount = 0;

            // Analyze data (skip header row)
            for (let row = 1; row <= range.e.r; row++) {
                const actionCellAddr = `${actionCol}${row + 1}`;
                const actionCell = worksheet[actionCellAddr];
                
                if (!actionCell && !dueDateCol) continue; // Skip completely empty rows
                
                totalRows++;
                const isActionAssigned = this.isAssigned(actionCell?.w || actionCell?.v);
                
                if (isActionAssigned) {
                    assignedCount++;
                    
                    if (dueDateCol) {
                        const dueDateCellAddr = `${dueDateCol}${row + 1}`;
                        const dueDateCell = worksheet[dueDateCellAddr];
                        const dueDate = this.parseDate(dueDateCell?.w || dueDateCell?.v);
                        
                        if (dueDate && dueDate < today) {
                            overdueCount++;
                        }
                    }
                }
            }

            const assignedPct = totalRows > 0 ? (assignedCount / totalRows) * 100 : 0;
            const overduePctOfAssigned = assignedCount > 0 ? (overdueCount / assignedCount) * 100 : 0;

            const summary = `Total: ${totalRows} | Assigned: ${assignedCount} (${assignedPct.toFixed(0)}%) | Overdue: ${overdueCount} (${overduePctOfAssigned.toFixed(0)}%)`;

            const analysisResult = {
                total_rows: totalRows,
                assigned_count: assignedCount,
                assigned_pct: Math.round(assignedPct * 10) / 10,
                overdue_count: overdueCount,
                overdue_pct_of_assigned: Math.round(overduePctOfAssigned * 10) / 10,
                today_iso: today.toISOString().split('T')[0],
                timezone: 'Asia/Kolkata',
                columns_used: {
                    action: actionCol,
                    due_date: dueDateCol || 'Not found'
                },
                notes,
                summary
            };

            this.result = analysisResult;
            this.showResults(analysisResult);
            this.showToast(`Analyzed ${totalRows} rows successfully`);

        } catch (err) {
            const errorMessage = err.message || 'Failed to analyze Excel file';
            this.showError(errorMessage);
            this.showToast(errorMessage, 'error');
        } finally {
            this.setAnalyzingState(false);
        }
    }
}

// Initialize the analyzer when DOM is loaded
document.addEventListener('DOMContentLoaded', () => {
    new ExcelAnalyzer();
});