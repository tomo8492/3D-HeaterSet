AgGrid_Custom_Header = '''
class SelectableHeader {
    init(params) {
        this.params = params;
        this.eGui = document.createElement('div');
        this.eGui.style.display = 'flex';
        this.eGui.style.alignItems = 'center';
        this.eGui.style.gap = '5px';
        this.eGui.style.cursor = 'pointer';
        
        this.dot = document.createElement('span');
        this.dot.innerHTML = '●';
        this.dot.style.color = params.selected ? '#4CAF50' : '#ccc';
        
        const label = document.createElement('span');
        label.textContent = params.displayName;
        
        this.eGui.appendChild(this.dot);
        this.eGui.appendChild(label);
        
        // Toggle on Alt+Click
        this.eGui.addEventListener('click', (e) => {
            if (e.altKey) {
                params.selected = !params.selected;
                this.dot.style.color = params.selected ? '#4CAF50' : '#ccc';
                
                emitEvent('columnSelected', {
                    column: params.column.colId,
                    selected: params.selected
                });
            }
        });
    }
    
    getGui() {
        return this.eGui;
    }
}
'''