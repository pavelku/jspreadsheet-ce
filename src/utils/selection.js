import dispatch from "./dispatch.js";
import { getFreezeWidth } from "./freeze.js";
import { getCellNameFromCoords } from "./helpers.js";
import { setHistory } from "./history.js";
import { updateCell, updateFormula, updateFormulaChain, updateTable } from "./internal.js";
import { getColumnNameFromId, getIdFromColumnName } from "./internalHelpers.js";
import { updateToolbar } from "./toolbar.js";

export const updateCornerPosition = function() {
    const obj = this;

    // If any selected cells
    if (!obj.highlighted || !obj.highlighted.length) {
        obj.corner.style.top = '-2000px';
        obj.corner.style.left = '-2000px';
    } else {
        // Get last cell
        const last = obj.highlighted[obj.highlighted.length-1].element;
        const lastX = last.getAttribute('data-x');

        const contentRect = obj.content.getBoundingClientRect();
        const x1 = contentRect.left;
        const y1 = contentRect.top;

        const lastRect = last.getBoundingClientRect();
        const x2 = lastRect.left;
        const y2 = lastRect.top;
        const w2 = lastRect.width;
        const h2 = lastRect.height;

        const x = (x2 - x1) + obj.content.scrollLeft + w2 - 4;
        const y = (y2 - y1) + obj.content.scrollTop + h2 - 4;

        // Place the corner in the correct place
        obj.corner.style.top = y + 'px';
        obj.corner.style.left = x + 'px';

        if (obj.options.freezeColumns) {
            const width = getFreezeWidth.call(obj);
            // Only check if the last column is not part of the merged cells
            if (lastX > obj.options.freezeColumns-1 && x2 - x1 + w2 < width) {
                obj.corner.style.display = 'none';
            } else {
                if (obj.options.selectionCopy != false) {
                    obj.corner.style.display = '';
                }
            }
        } else {
            if (obj.options.selectionCopy != false) {
                obj.corner.style.display = '';
            }
        }
    }

    updateToolbar(obj);
}

export const resetSelection = function(blur) {
    const obj = this;

    let previousStatus;

    // Remove style
    if (!obj.highlighted || !obj.highlighted.length) {
        previousStatus = 0;
    } else {
        previousStatus = 1;

        for (let i = 0; i < obj.highlighted.length; i++) {
            obj.highlighted[i].element.classList.remove('highlight');
            obj.highlighted[i].element.classList.remove('highlight-left');
            obj.highlighted[i].element.classList.remove('highlight-right');
            obj.highlighted[i].element.classList.remove('highlight-top');
            obj.highlighted[i].element.classList.remove('highlight-bottom');
            obj.highlighted[i].element.classList.remove('highlight-selected');

            const px = parseInt(obj.highlighted[i].element.getAttribute('data-x'));
            const py = parseInt(obj.highlighted[i].element.getAttribute('data-y'));

            // Check for merged cells
            let ux, uy;

            if (obj.highlighted[i].element.getAttribute('data-merged')) {
                const colspan = parseInt(obj.highlighted[i].element.getAttribute('colspan'));
                const rowspan = parseInt(obj.highlighted[i].element.getAttribute('rowspan'));
                ux = colspan > 0 ? px + (colspan - 1) : px;
                uy = rowspan > 0 ? py + (rowspan - 1): py;
            } else {
                ux = px;
                uy = py;
            }

            // Remove selected from headers
            for (let j = px; j <= ux; j++) {
                if (obj.headers[j]) {
                    obj.headers[j].classList.remove('selected');
                }
            }

            // Remove selected from rows
            for (let j = py; j <= uy; j++) {
                if (obj.rows[j]) {
                    obj.rows[j].element.classList.remove('selected');
                }
            }
        }
    }

    // Reset highlighted cells
    obj.highlighted = [];

    // Reset
    obj.selectedCell = null;
    // obj.startSelCol = obj.endSelCol = obj.startSelRow = obj.endSelRow = undefined;
    // console.log('reset style');

    // Hide corner
    obj.corner.style.top = '-2000px';
    obj.corner.style.left = '-2000px';

    if (blur == true && previousStatus == 1) {
        dispatch.call(obj, 'onblur', obj);
    }

    return previousStatus;
}

/**
 * Update selection based on two cells
 */
export const updateSelection = function(el1, el2, origin) {
    const obj = this;

    const x1 = el1.getAttribute('data-x');
    const y1 = el1.getAttribute('data-y');

    let x2, y2;
    if (el2) {
        x2 = el2.getAttribute('data-x');
        y2 = el2.getAttribute('data-y');
    } else {
        x2 = x1;
        y2 = y1;
    }

    updateSelectionFromCoords.call(obj, x1, y1, x2, y2, origin);
}

export const removeCopyingSelection = function() {
    const copying = document.querySelectorAll('.jss_worksheet .copying');
    for (let i = 0; i < copying.length; i++) {
        copying[i].classList.remove('copying');
        copying[i].classList.remove('copying-left');
        copying[i].classList.remove('copying-right');
        copying[i].classList.remove('copying-top');
        copying[i].classList.remove('copying-bottom');
    }
}

export const updateSelectionFromCoords = function(x1, y1, x2, y2, origin) {
    const obj = this;

    console.log('--updateSelectionFromCoords-- y1 = [', y1, '] y2 = [', y2, '] x1 = [', x1, '] x2 = [', x2, '], scrollDirection ', obj.scrollDirection, ', mouseOverDirection = ', obj.mouseOverDirection);

    var selectWholeColumn = false;
    var isRowSelected = false;

    if (y2 == obj.totalItemsInQuery && y1 == 0) {
        selectWholeColumn = true;
    }

    // select column
    if (y1 == null) {
        // console.log('oncolumn click , total items in query = ', obj.totalItemsInQuery);
        y1 = 0;
        y2 = obj.rows.length - 1;; // 

        if (x1 == null) {
            return;
        }
    } else if (x1 == null || x1 == 0) {
        // select row
        isRowSelected = true;
        x1 = 0;
        x2 = obj.options.data[0].length - 1;
    }

    // Same element
    if (x2 == null) {
        x2 = x1;
    }
    if (y2 == null) {
        y2 = y1;
    }

    // Selection must be within the existing data
    if (x1 >= obj.headers.length) {
        x1 = obj.headers.length - 1;
    }
    if (y1 >= obj.rows.length) {        
        y1 = obj.rows.length - 1;
    }
    if (x2 >= obj.headers.length) {
        x2 = obj.headers.length - 1;
    }
    if (y2 >= obj.rows.length) {
        y2 = obj.rows.length - 1;
    }

    // Limits
    let borderLeft = null;
    let borderRight = null;
    let borderTop = null;
    let borderBottom = null;

    // Origin & Destination
    let px, ux;

    if (parseInt(x1) < parseInt(x2)) {
        px = parseInt(x1);
        ux = parseInt(x2);
    } else {
        px = parseInt(x2);
        ux = parseInt(x1);
    }

    let py, uy;

    if (parseInt(y1) < parseInt(y2)) {
        py = parseInt(y1);
        uy = parseInt(y2);
    } else {
        py = parseInt(y2);
        uy = parseInt(y1);
    }

    // console.log('py = ', py, ', uy = ')

    // Verify merged columns
    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            if (obj.records[j][i] && obj.records[j][i].element.getAttribute('data-merged')) {
                const x = parseInt(obj.records[j][i].element.getAttribute('data-x'));
                const y = parseInt(obj.records[j][i].element.getAttribute('data-y'));
                const colspan = parseInt(obj.records[j][i].element.getAttribute('colspan'));
                const rowspan = parseInt(obj.records[j][i].element.getAttribute('rowspan'));

                if (colspan > 1) {
                    if (x < px) {
                        px = x;
                    }
                    if (x + colspan > ux) {
                        ux = x + colspan - 1;
                    }
                }

                if (rowspan) {
                    if (y < py) {
                        py = y;

                    }
                    if (y + rowspan > uy) {
                        uy = y + rowspan - 1;
                    }
                }
            }
        }
    }

    // Vertical limits
    for (let j = py; j <= uy; j++) {
        if (obj.rows[j].element.style.display != 'none') {
            if (borderTop == null) {
                borderTop = j;
            }
            borderBottom = j;
        }
    }

    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            // Horizontal limits
            if (!obj.options.columns || !obj.options.columns[i] || obj.options.columns[i].type != 'hidden') {
                if (borderLeft == null) {
                    borderLeft = i;
                }
                borderRight = i;
            }
        }
    }

    // Create borders
    if (! borderLeft) {
        borderLeft = 0;
    }
    if (! borderRight) {
        borderRight = 0;
    }

    const ret = dispatch.call(obj, 'onbeforeselection', obj, borderLeft, borderTop, borderRight, borderBottom, origin);
    if (ret === false) {
        return false;
    }

    // Reset Selection
    const previousState = obj.resetSelection();

    // Keep selected cell
    obj.selectedCell = [x1, y1, x2, y2];

    // Add selected cell
    if (obj.records[y1][x1]) {
        obj.records[y1][x1].element.classList.add('highlight-selected');
    }

    // Redefining styles
    for (let i = px; i <= ux; i++) {
        for (let j = py; j <= uy; j++) {
            if (obj.rows[j].element.style.display != 'none' && obj.records[j][i].element.style.display != 'none') {
                obj.records[j][i].element.classList.add('highlight');
                obj.highlighted.push(obj.records[j][i]);
            }
        }
    }

    for (let i = borderLeft; i <= borderRight; i++) {
        if ((!obj.options.columns || !obj.options.columns[i] || obj.options.columns[i].type != 'hidden') && obj.cols[i].colElement.style && obj.cols[i].colElement.style.display != 'none') {
            // Top border
            if (obj.records[borderTop] && obj.records[borderTop][i]) {
                obj.records[borderTop][i].element.classList.add('highlight-top');
            }
            // Bottom border
            if (obj.records[borderBottom] && obj.records[borderBottom][i]) {
                obj.records[borderBottom][i].element.classList.add('highlight-bottom');
            }
            // Add selected from headers
            obj.headers[i].classList.add('selected');
        }
    }

    for (let j = borderTop; j <= borderBottom; j++) {
        if (obj.rows[j] && obj.rows[j].element.style.display != 'none') {
            // Left border
            obj.records[j][borderLeft].element.classList.add('highlight-left');
            // Right border
            obj.records[j][borderRight].element.classList.add('highlight-right');
            // Add selected from rows
            obj.rows[j].element.classList.add('selected');
        }
    }

    obj.selectedContainer = [ borderLeft, borderTop, borderRight, borderBottom ];

    // Handle events
    if (previousState == 0) {
        dispatch.call(obj, 'onfocus', obj);

        removeCopyingSelection();
    }

    // console.log('before onselection obj.startSelCol = ', obj.startSelCol, ', obj.endSelCol = ', obj.endSelCol, ', obj.startSelRow = ', obj.startSelRow, ', obj.endSelRow = ', obj.endSelRow);
    dispatch.call(obj, 'onselection', obj, borderLeft, borderTop, borderRight, borderBottom, origin);

    // vyvolano mysi
    if (origin) {
        console.log('mouse');
        // kliknu na libovolnou bunku (bez mouse move)
        if (origin.type == "mousedown" && !origin.shiftKey){

            console.log('!! first mouse, mouseOverDirection set none');

            // prvni klik nastav start sel row
            const startRowIndex = obj.getRowData(y1)[0];
            const endRowIndex = obj.getRowData(y2)[0];   

            if (!selectWholeColumn) {
                obj.startSelCol = x1;
                obj.startSelRow = startRowIndex;            
                obj.endSelRow = startRowIndex;            
                obj.endSelCol = x2 ? x2 : x1;
                obj.mouseOverDirection = "none";
                obj.keyOverDirection = "none";                                         
            }
            else {
                obj.startSelCol = x1;
                obj.endSelCol = x2 ? x2 : x1;
                obj.startSelRow = 1;     
                obj.endSelRow = obj.totalItemsInQuery;
                obj.mouseOverDirection = "down";
                obj.keyOverDirection = "none";                                         
            }

            // obj.endSelCol = x2;    
            // obj.endSelRow = !selectWholeColumn ? endRowIndex : obj.totalItemsInQuery;
            console.log('New Selection = [', obj.startSelRow , ',', obj.endSelRow, ']');
        }
        // pohyb mysi
        else if (origin.type == "mouseover" || (origin.type == "mousedown" && origin.shiftKey)) {            
            obj.startSelCol = x1;
            obj.endSelCol = x2;     

            if (origin.type == "mouseover")
            {
                console.log('!! mouseover');
                // mysi jedu dolu
                if (obj.mouseOverDirection == "down") {
                    const endRowIndex = obj.getRowData(y2)[0];    
                    obj.endSelRow = !selectWholeColumn ? endRowIndex : obj.totalItemsInQuery;                
                }
                // mysi jedu nahoru
                else if (obj.mouseOverDirection == "up") {
                    const startRowIndex = obj.getRowData(y2)[0];
                    obj.startSelRow = !selectWholeColumn ? startRowIndex : 1;                
                }
                // vybral jsem oblast ze shora dolu a pak jedu nahoru
                else if (obj.mouseOverDirection == "sellDownAndThanUp") {
                    const endRowIndex = obj.getRowData(y2)[0];    
                    obj.endSelRow = !selectWholeColumn ? endRowIndex : obj.totalItemsInQuery;
                }
                // vybral jsem oblast ze zdola nahoru a pak jedu dolu
                else if (obj.mouseOverDirection == "sellUpnAndThanDown") { 
                    const startRowIndex = obj.getRowData(y2)[0];    
                    obj.startSelRow = startRowIndex;
                }      
                else if (obj.mouseOverDirection == "none") { 
                    const startRowIndex = obj.getRowData(y1)[0];    
                    obj.startSelRow = obj.startSelRow = startRowIndex;
                }
            }     

            // TODO comment mousedown with shift key
            if (origin.type == "mousedown" && origin.shiftKey)
            {
                // druhy klik -> nastav end position
                var data = obj.getData();                
                obj.endSelRow = obj.getRowData(y2)[0];    

                if (obj.startSelRow < obj.endSelRow)
                {
                    obj.mouseOverDirection = 'down';                    
                }
                else {
                    obj.mouseOverDirection = 'up';                    
                }

                if (obj.startSelRow > obj.endSelRow) {
                    const tmpStart = obj.startSelRow;
                    obj.startSelRow = obj.endSelRow;
                    obj.endSelRow = tmpStart;
                }

                const firstRowPos = data[0][0];
                const endRowPos = data[data.length-1][0];
                const startPos = Math.max(firstRowPos, obj.startSelRow);
                const endPos = Math.min(endRowPos, obj.endSelRow);

                const startRowIndex = getDataByNrPos(data, startPos <= endPos ? startPos : endPos, 0);
                const endRowIndex = getDataByNrPos(data, startPos < endPos ? endPos : startPos, 0); // startRowIndex   
                
                if (obj.mouseOverDirection == "down") {
                    obj.selectedCell[1] = startRowIndex;
                    obj.selectedCell[3] = endRowIndex;                    
                }
                else if (obj.mouseOverDirection == "up") {
                    obj.selectedCell[1] = endRowIndex;
                    obj.selectedCell[3] = startRowIndex;                    
                }

                obj.keyDirectionDone = true;
                obj.preventOnSelection = true;
                console.log('!!!! mousedown in updateFromCoords DOWN WITH SCHIFT KEY, direction  = ', obj.mouseOverDirection, 'firstRowPos = ', firstRowPos, ', endRowPos = ', endRowPos, ' obj.startSelRow = ', obj.startSelRow, ', obj.endSelRow = ', obj.endSelRow);                
                refreshSelection.call(obj);    
                return;
            }

            // console.log('OnSelect MODE AFTER - obj.startSelRow = ', obj.startSelRow, ' obj.endSelRow = ', obj.endSelRow);
        }
        else {
            resetMousePos();
        }
    }
    // pohyb na klavesnici
    else {
        console.log('--updateSelectionFromCoords-- keyboard input obj.keyDirection = ', obj.keyDirection, 'obj.keyDirectionDone = ', obj.keyDirectionDone, ', obj.keyOverDirection = ', obj.keyOverDirection);
        // if (obj.keyDirection != -1) {
        if (!obj.keyDirectionDone) {
            obj.startSelCol = x1;
            obj.endSelCol = x2;     
            // 0 = doleva
            // 1 = nahoru
            // 2 = doprava
            // 3 = dolu
            const endRowIndex = obj.getRowData(y2)[0]; 
            const startRowIndex = obj.getRowData(y1)[0];            

            // vybrana oblast ze shora dolu
            if (y1 < y2) {   
                if (obj.keyDirection == 1 || obj.keyDirection == 3) {
                    const lastVisibleRow = obj.getRowData(obj.rows.length - 1)[0];            
                    console.log('??? lastVisibleRow = ', lastVisibleRow);
                    if (obj.endSelRow <= lastVisibleRow) {
                        obj.endSelRow = endRowIndex;
                        console.log('--updateSelectionFromCoords-- 1 - keyboard input endSelRow set = ', obj.endSelRow);
                    }
                    else {
                        if (obj.keyDirection == 1) {                            
                            obj.endSelRow = obj.endSelRow - 1;
                        }
                        else if (obj.endSelRow < obj.totalItemsInQuery) {
                            obj.endSelRow = obj.endSelRow + 1;
                        }

                        console.log('??? --updateSelectionFromCoords-- 1-1 Without visible location - keyboard input endSelRow set = ', obj.endSelRow);
                    }
                }
            }
            // vybrana oblast ze zdola nahoru
            else if (y1 > y2) {  
                if (obj.keyDirection == 1 || obj.keyDirection == 3) {
                    const firstVisibleRow = obj.getRowData(0)[0];            
                    console.log('??? firstVisibleRow = ', firstVisibleRow);
                    if (obj.startSelRow >= firstVisibleRow) {
                        obj.startSelRow = endRowIndex;   
                        console.log('--updateSelectionFromCoords-- 2 - keyboard input endSelRow set = ', obj.endSelRow);             
                    }
                    else {
                        if (obj.keyDirection == 1) {
                            if (obj.startSelRow > 1) {
                                obj.startSelRow = obj.startSelRow - 1;
                            }
                        }
                        else if (obj.startSelRow > 1) {
                            obj.startSelRow = obj.startSelRow + 1;
                        }
                        console.log('??? --updateSelectionFromCoords-- 2-1 Without visible location - keyboard input endSelRow set = ', obj.endSelRow);             
                    }
                }
            }
            // vybrana jedna radka - pomoci sipek
            else {
                if (obj.shifKey)
                {
                    if (obj.keyDirection == 1) {
                        obj.endSelRow = endRowIndex;                                
                        console.log('--updateSelectionFromCoords-- 3 - keyboard input endSelRow set = ', obj.endSelRow);             
                    }
                    else if (obj.keyDirection == 3) {
                        obj.startSelRow = startRowIndex;   
                        console.log('--updateSelectionFromCoords-- 4 - keyboard input startSelRow set = ', obj.startSelRow);             
                    }
                    else {
                        obj.endSelRow = obj.startSelRow = endRowIndex;                         
                        console.log('--updateSelectionFromCoords-- 5 - keyboard input startSelRow and endSelRow set = ', obj.startSelRow);             
                    }
                }
                else {
                    obj.endSelRow = obj.startSelRow = endRowIndex;                         
                }
            }            
        }
        else {
            obj.preventOnSelection = false;
        }
        // }
    }    

    console.log('--updateSelectionFromCoords-- at the end  startPos = [', obj.startSelRow, ',', obj.startSelCol, '], endPos = [', obj.endSelRow, ',', obj.endSelCol, ']');

    // Find corner cell
    updateCornerPosition.call(obj);
}



export const chooseSelection = function (startPos, endPos) {
    const obj = this;

    var data = obj.getData();
   
    
    const startRowIndex = getDataByNrPos(data, startPos <= endPos ? startPos : endPos, 0);
    const endRowIndex = getDataByNrPos(data, startPos < endPos ? endPos : startPos, 0); // startRowIndex   
    console.log('choose selection ', obj.keyDirection, ', obj.keyOverDirection = ', obj.keyOverDirection, ', obj.mouseOverDirection = ', obj.mouseOverDirection, ', obj.keyDirectionDone = ', obj.keyDirectionDone);    
    
    if (obj.mouseOverDirection == "down" || obj.mouseOverDirection == "sellDownAndThanUp" || obj.keyOverDirection == "down") { // || obj.keyOverDirection == "3") {
        obj.selectedCell[1] = startRowIndex;
        obj.selectedCell[3] = endRowIndex;        
    }
    else if (obj.mouseOverDirection == "up" || obj.mouseOverDirection == "sellUpnAndThanDown" || obj.keyOverDirection == "up") { //|| obj.keyDirection == 1) {
        obj.selectedCell[1] = endRowIndex;
        obj.selectedCell[3] = startRowIndex;        
    }
    else if (obj.keyOverDirection == "equal" || obj.mouseOverDirection == "none") {
        obj.selectedCell[1] = startRowIndex;
        obj.selectedCell[3] = startRowIndex;        
    }

    // scroll down -> zacatek se odroluje
    if (obj.mouseOverDirection == "down") {
        if (startRowIndex != 0 && startRowIndex < 5 && endRowIndex < 3)
        {
            console.log('!!! obezle');
            obj.selectedCell[1] = 0;
            obj.keyDirectionDone = true;
        }
    }


    if (!obj.keyDirectionDone) {

        const y1 = obj.selectedCell[1];
        const y2 = obj.selectedCell[3];

        obj.startSelCol = obj.selectedCell[0];
        obj.endSelCol = obj.selectedCell[2];

        const endRowIndex = obj.getRowData(y2)[0]; 
        const startRowIndex = obj.getRowData(y1)[0]; 

        console.log('chooseSelection pohyb klavesnici smer = ', obj.keyDirection);

        // vybrana oblast ze shora dolu
        if (obj.keyOverDirection == "down") {// (y1 < y2) {   
            if (obj.keyDirection == 1 || obj.keyDirection == 3) {
                obj.endSelRow = endRowIndex;
                console.log('chooseSelection 1 - endSelRow set = ', obj.endSelRow);
            }
        }
        // vybrana oblast ze zdola nahoru
        else if (obj.keyOverDirection == "up") { // if (y1 > y2) {  
            if (obj.keyDirection == 1 || obj.keyDirection == 3) {
                obj.startSelRow = endRowIndex;                
                console.log('chooseSelection 2 - startSelRow set = ', obj.startSelRow);
            }
        }
        // vybrana jedna radka - pomoci sipek
        else if (obj.keyDirection == 3 || obj.keyDirection == 1) {
            if (obj.keyDirection == 1) {
                obj.endSelRow = endRowIndex;                                
                console.log('--chooseSelection-- 3 - keyboard input endSelRow set = ', obj.endSelRow);             
            }
            else if (obj.keyDirection == 3) {
                obj.startSelRow = startRowIndex;   
                console.log('--chooseSelection-- 4 - keyboard input startSelRow set = ', obj.startSelRow);             
            }           
            // obj.endSelRow = obj.startSelRow = endRowIndex;                         
            // console.log('chooseSelection 3 - endSelRow set = ', obj.endSelRow);       
        }

        obj.keyDirectionDone = true;
        obj.preventOnSelection = true;
    }

    refreshSelection.call(obj);    
}

const getDataByNrPos = (data, curPosNr, startIndex) =>{
    for (let j = startIndex; j < data.length; j++) {
        if (data[j][0] == curPosNr)
            return j;
    } 

    return -1;
}

const resetMousePos = ()  => {
    const obj = this;
    obj.startSelCol = obj.endSelCol = obj.startSelRow = obj.endSelRow = -1;
    obj.oldStartSelRow = obj.oldEndSelRow = -1;
}


/**
 * Get selected column numbers
 *
 * @return array
 */
export const getSelectedColumns = function(visibleOnly) {
    const obj = this;

    if (!obj.selectedCell) {
        return [];
    }

    const result = [];

    for (let i = Math.min(obj.selectedCell[0], obj.selectedCell[2]); i <= Math.max(obj.selectedCell[0], obj.selectedCell[2]); i++) {
        if (!visibleOnly || obj.headers[i].style.display != 'none') {
            result.push(i);
        }
    }

    return result;
}

/**
 * Refresh current selection
 */
export const refreshSelection = function() {
    const obj = this;

    if (obj.selectedCell) {
        obj.updateSelectionFromCoords(obj.selectedCell[0], obj.selectedCell[1], obj.selectedCell[2], obj.selectedCell[3]);
    }
}

/**
 * Remove copy selection
 *
 * @return void
 */
export const removeCopySelection = function() {
    const obj = this;

    // Remove current selection
    for (let i = 0; i < obj.selection.length; i++) {
        obj.selection[i].classList.remove('selection');
        obj.selection[i].classList.remove('selection-left');
        obj.selection[i].classList.remove('selection-right');
        obj.selection[i].classList.remove('selection-top');
        obj.selection[i].classList.remove('selection-bottom');
    }

    obj.selection = [];
}

const doubleDigitFormat = function(v) {
    v = ''+v;
    if (v.length == 1) {
        v = '0'+v;
    }
    return v;
}

/**
 * Helper function to copy data using the corner icon
 */
export const copyData = function(o, d) {
    const obj = this;

    // Get data from all selected cells
    const data = obj.getData(true, true);

    // Selected cells
    const h = obj.selectedContainer;

    // Cells
    const x1 = parseInt(o.getAttribute('data-x'));
    const y1 = parseInt(o.getAttribute('data-y'));
    const x2 = parseInt(d.getAttribute('data-x'));
    const y2 = parseInt(d.getAttribute('data-y'));

    // Records
    const records = [];
    let breakControl = false;

    let rowNumber, colNumber;

    if (h[0] == x1) {
        // Vertical copy
        if (y1 < h[1]) {
            rowNumber = y1 - h[1];
        } else {
            rowNumber = 1;
        }
        colNumber = 0;
    } else {
        if (x1 < h[0]) {
            colNumber = x1 - h[0];
        } else {
            colNumber = 1;
        }
        rowNumber = 0;
    }

    // Copy data procedure
    let posx = 0;
    let posy = 0;

    for (let j = y1; j <= y2; j++) {
        // Skip hidden rows
        if (obj.rows[j] && obj.rows[j].element.style.display == 'none') {
            continue;
        }

        // Controls
        if (data[posy] == undefined) {
            posy = 0;
        }
        posx = 0;

        // Data columns
        if (h[0] != x1) {
            if (x1 < h[0]) {
                colNumber = x1 - h[0];
            } else {
                colNumber = 1;
            }
        }
        // Data columns
        for (let i = x1; i <= x2; i++) {
            // Update non-readonly
            if (obj.records[j][i] && ! obj.records[j][i].element.classList.contains('readonly') && obj.records[j][i].element.style.display != 'none' && breakControl == false) {
                // Stop if contains value
                if (! obj.selection.length) {
                    if (obj.options.data[j][i] != '') {
                        breakControl = true;
                        continue;
                    }
                }

                // Column
                if (data[posy] == undefined) {
                    posx = 0;
                } else if (data[posy][posx] == undefined) {
                    posx = 0;
                }

                // Value
                let value = data[posy][posx];

                if (value && ! data[1] && obj.parent.config.autoIncrement != false) {
                    if (obj.options.columns && obj.options.columns[i] && (!obj.options.columns[i].type || obj.options.columns[i].type == 'text' || obj.options.columns[i].type == 'number')) {
                        if ((''+value).substr(0,1) == '=') {
                            const tokens = value.match(/([A-Z]+[0-9]+)/g);

                            if (tokens) {
                                const affectedTokens = [];
                                for (let index = 0; index < tokens.length; index++) {
                                    const position = getIdFromColumnName(tokens[index], 1);
                                    position[0] += colNumber;
                                    position[1] += rowNumber;
                                    if (position[1] < 0) {
                                        position[1] = 0;
                                    }
                                    const token = getColumnNameFromId([position[0], position[1]]);

                                    if (token != tokens[index]) {
                                        affectedTokens[tokens[index]] = token;
                                    }
                                }
                                // Update formula
                                if (affectedTokens) {
                                    value = updateFormula(value, affectedTokens)
                                }
                            }
                        } else {
                            if (value == Number(value)) {
                                value = Number(value) + rowNumber;
                            }
                        }
                    } else if (obj.options.columns && obj.options.columns[i] && obj.options.columns[i].type == 'calendar') {
                        const date = new Date(value);
                        date.setDate(date.getDate() + rowNumber);
                        value = date.getFullYear() + '-' + doubleDigitFormat(parseInt(date.getMonth() + 1)) + '-' + doubleDigitFormat(date.getDate()) + ' ' + '00:00:00';
                    }
                }

                records.push(updateCell.call(obj, i, j, value));

                // Update all formulas in the chain
                updateFormulaChain.call(obj, i, j, records);
            }
            posx++;
            if (h[0] != x1) {
                colNumber++;
            }
        }
        posy++;
        rowNumber++;
    }

    // Update history
    setHistory.call(obj, {
        action:'setValue',
        records:records,
        selection:obj.selectedCell,
    });

    // Update table with custom configuration if applicable
    updateTable.call(obj);

    // On after changes
    const onafterchangesRecords = records.map(function(record) {
        return {
            x: record.x,
            y: record.y,
            value: record.newValue,
            oldValue: record.oldValue,
        };
    });

    dispatch.call(obj, 'onafterchanges', obj, onafterchangesRecords);
}

export const hash = function(str) {
    let hash = 0, i, chr;

    if (str.length === 0) {
        return hash;
    } else {
        for (i = 0; i < str.length; i++) {
          chr = str.charCodeAt(i);
          hash = ((hash << 5) - hash) + chr;
          hash |= 0;
        }
    }
    return hash;
}

/**
 * Move coords to A1 in case overlaps with an excluded cell
 */
export const conditionalSelectionUpdate = function(type, o, d) {
    const obj = this;

    if (type == 1) {
        if (obj.selectedCell && ((o >= obj.selectedCell[1] && o <= obj.selectedCell[3]) || (d >= obj.selectedCell[1] && d <= obj.selectedCell[3]))) {
            obj.resetSelection();
            return;
        }
    } else {
        if (obj.selectedCell && ((o >= obj.selectedCell[0] && o <= obj.selectedCell[2]) || (d >= obj.selectedCell[0] && d <= obj.selectedCell[2]))) {
            obj.resetSelection();
            return;
        }
    }
}

/**
 * Get selected rows numbers
 *
 * @return array
 */
export const getSelectedRows = function(visibleOnly) {
    const obj = this;

    if (!obj.selectedCell) {
        return [];
    }

    const result = [];

    for (let i = Math.min(obj.selectedCell[1], obj.selectedCell[3]); i <= Math.max(obj.selectedCell[1], obj.selectedCell[3]); i++) {
        if (!visibleOnly || obj.rows[i].element.style.display != 'none') {
            result.push(i);
        }
    }

    return result;
}

export const selectAll = function() {
    const obj = this;

    if (! obj.selectedCell) {
        obj.selectedCell = [];
    }

    obj.selectedCell[0] = 1;
    obj.selectedCell[1] = 0;
    obj.selectedCell[2] = obj.headers.length - 1;
    obj.selectedCell[3] = obj.totalItemsInQuery;// obj.records.length - 1;

    obj.updateSelectionFromCoords(obj.selectedCell[0], obj.selectedCell[1], obj.selectedCell[2], obj.selectedCell[3]);
}

export const getSelection = function() {
    const obj = this;

    if (!obj.selectedCell) {
        return null;
    }

    return [
        Math.min(obj.selectedCell[0], obj.selectedCell[2]),
        Math.min(obj.selectedCell[1], obj.selectedCell[3]),
        Math.max(obj.selectedCell[0], obj.selectedCell[2]),
        Math.max(obj.selectedCell[1], obj.selectedCell[3]),
    ];
}

export const getSelected = function(columnNameOnly) {
    const obj = this;

    const selectedRange = getSelection.call(obj);

    if (!selectedRange) {
        return [];
    }

    const cells = [];

    for (let y = selectedRange[1]; y <= selectedRange[3]; y++) {
        for (let x = selectedRange[0]; x <= selectedRange[2]; x++) {
            if (columnNameOnly) {
                cells.push(getCellNameFromCoords(x, y));
            } else {
                cells.push(obj.records[y][x]);
            }
        }
    }

    return cells;
}

export const getRange = function() {
    const obj = this;

    const selectedRange = getSelection.call(obj);

    if (!selectedRange) {
        return '';
    }

    const start = getCellNameFromCoords(selectedRange[0], selectedRange[1]);
    const end = getCellNameFromCoords(selectedRange[2], selectedRange[3]);

    if (start === end) {
        return obj.options.worksheetName + '!' + start;
    }

    return obj.options.worksheetName + '!' + start + ':' + end;
}

export const isSelected = function(x, y) {
    const obj = this;

    const selection = getSelection.call(obj);

    return x >= selection[0] && x <= selection[2] && y >= selection[1] && y <= selection[3];
}

export const getHighlighted = function() {
    const obj = this;

    const selection = getSelection.call(obj);

    if (selection) {
        return [selection];
    }

    return [];
}