import * as React from 'react';
import GC from '@mescius/spread-sheets';
import '@mescius/spread-sheets-print';
import '@mescius/spread-sheets-io';
import '@mescius/spread-sheets-shapes';
import '@mescius/spread-sheets-charts';
import '@mescius/spread-sheets-slicers';
import '@mescius/spread-sheets-pivot-addon';
import '@mescius/spread-sheets-reportsheet-addon';
import "@mescius/spread-sheets-tablesheet";
import "@mescius/spread-sheets-ganttsheet";
import '@mescius/spread-sheets-print';
import '@mescius/spread-sheets-pdf';
import { SpreadSheets, Worksheet } from '@mescius/spread-sheets-react';
import './styles.css';

window.GC = GC;

export function AppFunc() {
    const [spread, setSpread] = React.useState(null);
    const [selectedFile, setSelectedFile] = React.useState(null);
    const [navigableHits, setNavigableHits] = React.useState([]);
    const [currentHitIndex, setCurrentHitIndex] = React.useState(0);
    const [isWorkbookReadyToUse, setWorkbookReady] = React.useState(false);
    const openOptions = {
        sjs: {},
        ssjson: {},
        xlsx: {},
        csv: {},
    }
    const selectedTextRef = React.useRef(null);
    const isPartialCellSelectionRef = React.useRef(false);
    const selectedCellPosition = React.useRef({ left: 0, top: 0 });
    const HiddenRowColumnMap = React.useRef({});
    const wasSheetHidden = React.useRef({});
    const showHiddenDataRef = React.useRef(false);

    function initSpread(spread) {

        setSpread(spread);
        //init Status Bar
        var statusBar = new GC.Spread.Sheets.StatusBar.StatusBar(document.getElementById('statusBar'));
        statusBar.bind(spread);
        document.getElementById('vp').style.height = '80vh'

        if (!spread) {
            return;
        }

    }

    function open() {
        var file = selectedFile;
        if (!file) {
            return;
        }

        var fileType = getFileType(file);
        var options = deepClone(openOptions[fileType]);

        if (fileType === FileType.SJS) {
            spread.open(file, function () { }, function (err) {
                console.log(err);
            }, options);
        } else {
            spread.import(file, function () { }, function () { }, options);
        }
    }

    function onSelectedFileChange(e) {
        let selectedFile = e.target.files[0];
        let openFileType = getFileType(selectedFile);
        setSelectedFile(selectedFile);
    }

    function HighlightText(searchResults, cellText, row, col, activeSheet, spread, isExtracted) {
        const highlightCommand = {
            canUndo: true,
            execute: function (spread, options, isUndo) {
                const Commands = GC.Spread.Sheets.Commands;
                if (isUndo) {
                    Commands.undoTransaction(spread, options);
                    return true;
                } else {
                    Commands.startTransaction(spread, options);

                    //the whole text cell is matched so just highlight that simply
                    if (searchResults.length === 1 && searchResults[0].text === cellText) {
                        activeSheet.getCell(row, col).foreColor("#FFDF00")
                    } else {
                        const cellContent = { richText: [] };
                        let lastIndex = 0;
                        searchResults.forEach((result, index) => {

                            if (result.index < lastIndex) {
                                result.index = lastIndex
                            } else {
                                //push not highlighted text
                                cellContent.richText.push({ text: cellText.substring(lastIndex, result.index) });
                            }

                            //push highlighted text
                            lastIndex = result.index + result.text.length;
                            cellContent.richText.push({ style: { foreColor: isExtracted ? "green" : "#FFDF00" }, text: cellText.substring(result.index, lastIndex) });
                        });
                        if (lastIndex < cellText.length) {
                            const remainingText = cellText.substring(lastIndex);
                            cellContent.richText.push({ text: remainingText });
                        }
                        activeSheet.setValue(row, col, cellContent);
                    }

                    //hover effect
                    activeSheet.comments.add(row, col, isExtracted ? "Extracted" : "Search_Result");
                    const activeComment = activeSheet.comments.get(row, col)
                    activeComment.width(isExtracted ? 110 : 150)
                    activeComment.height(35)
                    activeComment.fontSize('14' + "pt");
                    activeComment.fontWeight('bold');
                    activeComment.borderWidth(0);
                    activeComment.backColor(isExtracted ? "green" : '#FFDF00');
                    activeComment.zIndex(10000000000000);


                    Commands.endTransaction(spread, options);
                    return true;
                }
            }
        };

        const commandManager = spread.commandManager();
        commandManager.register('highlightCommand-' + row + '-' + col, highlightCommand);
        commandManager.execute({
            cmd: 'highlightCommand-' + row + '-' + col,
            sheetName: spread.getActiveSheet().name(),
            customID: 'highlightCommand-' + row + '-' + col
        });
    }

    function search() {
        console.log("searching", new Date().toLocaleTimeString())
        let navigableHits = [];
        let searchStrings = document.getElementById('search-text').value;
        if (!searchStrings || searchStrings.length === 0) {
            searchStrings = 'Hall,Smith,Boyd,Trevor,Curtis,Brian,jones rachael,859-86-8326,211-43-1582,713-62-9309';
        }
        searchStrings = searchStrings.split(",")

        const activeSheet = spread.getActiveSheet()
        const range = activeSheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data);
        if (!range) {
            return;
        }
        while (spread.undoManager().getUndoStack().length > 0) {
            spread.undoManager().undo();
        }

        spread.suspendPaint();
        for (var i = range.row; i < range.row + range.rowCount; i++) {
            for (var j = range.col; j < range.col + range.colCount; j++) {
                let text = activeSheet.getValue(i, j);
                if (typeof text !== 'string')
                    text = activeSheet.getText(i, j);
                const searchResults = findMatches(text, searchStrings, false);
                if (searchResults.length > 0) {
                    navigableHits.push({ row: i, col: j, sheetIndex: spread.getActiveSheetIndex() });
                    HighlightText(searchResults, text, i, j, activeSheet, spread);
                }
            }
        }
        setNavigableHits(navigableHits)
        spread.resumePaint();
        console.log("done", new Date().toLocaleTimeString())

    }

    function getTextForSelectedCells() {
        const activeSheet = spread.getActiveSheet();
        const selections = activeSheet.getSelections();
        const textList = [];
        selections.forEach(s => {
            for (let r = s.row; r <= s.row + s.rowCount - 1; r++) {
                for (let c = s.col; c <= s.col + s.colCount - 1; c++) {
                    textList.push(activeSheet.getText(r, c));
                }
            }
        })
        return textList
    }

    function getContextMenu(selectedText) {
        //selected Text is only passed down to this function if its a partial cell text selection
        isPartialCellSelectionRef.current = selectedText && selectedText.length > 0;
        selectedTextRef.current = selectedText
        return [
            {
                text: "Extract",
                name: "extract",
                command: () => {
                    const activeSheet = spread.getActiveSheet()
                    const row = activeSheet.getActiveRowIndex()
                    const col = activeSheet.getActiveColumnIndex()
                    if (selectedText) {
                        HighlightText([{ text: selectedText, index: activeSheet.getText(row, col).indexOf(selectedText) }], activeSheet.getText(row, col), row, col, activeSheet, spread, true);
                        hideCustomContextMenu()
                    }
                    else {
                        activeSheet.getCell(row, col).foreColor("green")
                        const activeComment = activeSheet.comments.get(row, col)
                        if (activeComment) {
                            activeComment.backColor('green');
                            activeComment.text("Extracted")
                            activeComment.width(110)
                        }
                    }
                },
                workArea: "viewport"
            },
            {
                text: "Copy",
                name: 'copy',
                command: () => {
                    //simply copy the selected text if user is manually highlighting some text.
                    if (selectedText)
                        navigator.clipboard.writeText(selectedText)
                    else {
                        const textToCopyList = getTextForSelectedCells();
                        navigator.clipboard.writeText(textToCopyList.join(", "));
                    }
                    hideCustomContextMenu()
                },
                workArea: "viewport"
            }
        ]
    }


    function showCustomContextMenu(options, x, y) {
        // closes previous opened context menu
        hideCustomContextMenu();
        const contextMenuHost = createContextMenu(options, x, y);
        const hostElement = document.querySelector('[gcuielement="gcLayerContainer"]');
        hostElement.appendChild(contextMenuHost);
    }

    // hides custom context menu
    function hideCustomContextMenu() {
        const element = document.querySelector('.custom-context-menu-container');
        if (element)
            element.remove();
    }

    function createContextMenu(options, x, y) {
        const container = document.createElement("div");
        container.style.position = "absolute";
        container.style.zIndex = '2100';
        container.style.left = `${x}px`;
        container.style.top = `${y}px`;
        container.classList.add('custom-context-menu-container');

        options.forEach((option) => {
            const optionDiv = document.createElement('div');
            optionDiv.classList.add('custom-context-menu-item');
            option.disable && optionDiv.classList.add('disabled');
            optionDiv.textContent = option.text;
            optionDiv.onclick = !option.disable && option.command;
            container.appendChild(optionDiv);
        });
        return container;
    }

    function extractData() {
        const activeSheetIndex = spread.getActiveSheetIndex();
        const sheetIndexList = [activeSheetIndex, ...spread.sheets.map((_, index) => index).filter(index => index !== activeSheetIndex)];
        for (let index = 0; index < sheetIndexList.length; index++) {
            const sheetIndex = sheetIndexList[index];
            const sheet = spread.sheets[sheetIndex];
            const range = sheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data);
            const nonEmptyColumns = getNonEmptyCols(sheet);
            const groupedColumns = groupConsecutiveColumns(nonEmptyColumns)
            console.log(`Sheet-${sheetIndex} Data`, groupedColumns.map(group => sheet.getArray(0, group[0], range.row + range.rowCount, group[1])))
        }
    }

    function fitRowsToPageHeight() {
        const changeAmount = 0.1;
        const activeSheet = spread.getActiveSheet();
        const rowCount = activeSheet.getRowCount(GC.Spread.Sheets.SheetArea.rowHeader);
        let lastVisibleRowOnScreen = activeSheet.getViewportBottomRow(1) + 1;
        let newZoom = 1;
        while (newZoom > 0.25) {
            if (rowCount <= lastVisibleRowOnScreen)
                break;
            newZoom -= changeAmount;
            activeSheet.zoom(newZoom);
            lastVisibleRowOnScreen = activeSheet.getViewportBottomRow(1) + 1;
        }
        activeSheet.zoom(newZoom);
    }

    function toggleShowHiddenData(showHiddenData) {
        function showHiddenRowsAndColumn(visible) {
            spread.sheets.forEach((sheet, sheetIndex) => {
                if (wasSheetHidden.current[sheetIndex]) {
                    sheet.visible(visible)
                }
                if (HiddenRowColumnMap.current[Number(sheetIndex)]) {
                    const { hiddenColumns, hiddenRows } = HiddenRowColumnMap.current[Number(sheetIndex)];
                    hiddenColumns.forEach(columnIndex => {
                        sheet.setColumnVisible(columnIndex, visible, GC.Spread.Sheets.SheetArea.colHeader);
                    })
                    hiddenRows.forEach(rowIndex => {
                        sheet.setRowVisible(rowIndex, visible, GC.Spread.Sheets.SheetArea.rowHeader);
                    })
                }
            });
        }

        const hasProcessedHiddenData = Object.keys(HiddenRowColumnMap.current).length > 0;
        setTimeout(() => {
            spread.suspendPaint();
            spread.suspendEvent();
            //if we already have hidden rows and column information, just use that.
            if (hasProcessedHiddenData)
                showHiddenRowsAndColumn(showHiddenData)
            else if (showHiddenData)
                //go through the sheets and unhide hidden rows and columns and also generate the hidden rows and column map.
                spread.sheets.forEach((sheet, sheetIndex) => {
                    if (!sheet.visible()) {
                        wasSheetHidden.current[sheetIndex] = true;
                        sheet.visible(true)
                    }
                    //unhide all columns
                    const columnCount = sheet.getColumnCount(GC.Spread.Sheets.SheetArea.colHeader);
                    for (let i = 0; i < columnCount; i++) {
                        if (!sheet.getColumnVisible(i, GC.Spread.Sheets.SheetArea.colHeader)) {
                            if (!HiddenRowColumnMap.current[sheetIndex])
                                HiddenRowColumnMap.current[sheetIndex] = { hiddenColumns: [i], hiddenRows: [] };
                            else
                                HiddenRowColumnMap.current[sheetIndex].hiddenColumns.push(i);
                            sheet.setColumnVisible(i, true, GC.Spread.Sheets.SheetArea.colHeader)
                        }
                    }
                    //unhide all rows
                    const rowCount = sheet.getRowCount(GC.Spread.Sheets.SheetArea.rowHeader);
                    for (let i = 0; i < rowCount; i++) {
                        if (!sheet.getRowVisible(i, GC.Spread.Sheets.SheetArea.rowHeader)) {
                            if (!HiddenRowColumnMap.current[sheetIndex])
                                HiddenRowColumnMap.current[sheetIndex] = { hiddenColumns: [], hiddenRows: [i] };
                            else
                                HiddenRowColumnMap.current[sheetIndex].hiddenRows.push(i);
                            sheet.setRowVisible(i, true, GC.Spread.Sheets.SheetArea.rowHeader)
                        }
                    }
                });
            spread.resumePaint();
            spread.resumeEvent();
        }, 100)
    }

    //file opening
    React.useEffect(() => {
        if (spread)
            fetch('00000029.xlsx')
                .then(res => res.blob())
                .then((blob) => {
                    const file = new File([blob], 'excel.xlsx', { type: blob.type });
                    spread.import(file, () => {
                        setWorkbookReady(true)
                        search();
                    }, (error) => {
                        console.log('error', error);
                    });
                });
    }, [spread]);

    //hit navigation
    React.useEffect(() => {
        if (navigableHits.length > 0) {
            const hit = navigableHits[currentHitIndex];
            if (hit) {
                spread.setActiveSheetIndex(hit.sheetIndex);
                spread.getActiveSheet().setActiveCell(hit.row, hit.col);
                spread.getActiveSheet().showColumn(hit.col, GC.Spread.Sheets.HorizontalPosition.left);
                spread.getActiveSheet().showRow(hit.row, GC.Spread.Sheets.VerticalPosition.top);
            }
        }
    }, [navigableHits, currentHitIndex])

    //custom context menu
    React.useEffect(() => {

        if (!spread || !isWorkbookReadyToUse) {
            return;
        }
        let isSelectingText = false;
        const spreadHost = document.querySelector('.sample-spreadsheets');

        function onSpreadHostMouseDown(e) {
            const editingElement = document.querySelector('[gcuielement="gcEditingInput"]')
            const sheet = spread.getActiveSheet();
            if (editingElement && sheet.isEditing() && (editingElement === e.target || editingElement.contains(e.target))) {
                isSelectingText = true;
            }
        }

        function onSpreadHostMouseUp() {
            const editingElement = document.querySelector('[gcuielement="gcEditingInput"]')
            const sheet = spread.getActiveSheet();
            if (sheet.isEditing() && editingElement && isSelectingText) {
                const selection = window.getSelection();
                const selectedText = selection.toString();

                if (selectedText !== '') {
                    //generates a new custom menu that will show up after user is done selecting some text.
                    let customMenuX, customMenuY;
                    try {
                        const range = selection.getRangeAt(0);
                        const clientRects = range.getClientRects();
                        const lastHighLightRect = clientRects[clientRects.length - 1];

                        const spreadJSEl = document.querySelector(".viewer_document");
                        const spreadJSElClientRect = spreadJSEl.getBoundingClientRect();
                        customMenuX = (lastHighLightRect.x - spreadJSElClientRect.x); // Add offset to avoid blocking highlight text
                        customMenuY = (lastHighLightRect.y - spreadJSElClientRect.y) + 20;
                    } catch (error) {
                        const cellRect = sheet.getCellRect(sheet.getActiveRowIndex(), sheet.getActiveColumnIndex(), 1, 1);
                        customMenuX = cellRect.x;
                        customMenuY = cellRect.y + cellRect.height;
                    }

                    showCustomContextMenu(
                        getContextMenu(selectedText),
                        customMenuX,
                        customMenuY
                    );
                }
                isSelectingText = false;
            }
        }

        function onMouseDown(event) {
            const target = event.target;
            const contextMenuContainer = document.querySelector('.custom-context-menu-container');
            if (target !== contextMenuContainer && contextMenuContainer && !contextMenuContainer.contains(target)) {
                hideCustomContextMenu();
            }
        }

        function onMouseUp(e) {
            selectedCellPosition.current = {
                left: e.x,
                top: e.y
            }
        }

        //setting the context menu that opens on right click
        const newContextMenu = getContextMenu();
        const wrapper = document.querySelector < HTMLDivElement > ('#gc-dialog1 > div > div:nth-child(1)');
        if (wrapper && spread.contextMenu.menuData.length && spread.contextMenu.menuData[0].disable == true && newContextMenu[0].disable == false) {

            const newElements = spread.contextMenu.menuView.createMenuItemElement(newContextMenu[0])
            const element = newElements[0];

            element.onclick = function () {
                typeof spread.contextMenu.menuData[0].command == 'function' ? spread.contextMenu.menuData[0].command() : () => { };

                //remove dialog
                document.querySelector('#gc-dialog1').remove();
                document.querySelector('.gc-overlay-gc-dialog1').remove();
            };
            //
            wrapper.onmouseenter = () => {
                wrapper.classList.remove('gc-ui-contextmenu-disable-hover');
                wrapper.classList.add('gc-ui-contextmenu-hover');
                wrapper.classList.add('ui-state-hover');
            }

            wrapper.firstChild.remove();
            wrapper.appendChild(element);
        }
        spread.contextMenu.menuData = newContextMenu;
        //setting the context menu that opens on text selection
        spreadHost.addEventListener('mousedown', onSpreadHostMouseDown);
        spreadHost.addEventListener('mouseup', onSpreadHostMouseUp);

        document.addEventListener('mousedown', onMouseDown, true);
        document.addEventListener('mouseup', onMouseUp)

        return () => {
            if (isWorkbookReadyToUse) {
                spreadHost.removeEventListener('mousedown', onSpreadHostMouseDown);
                spreadHost.removeEventListener('mouseup', onSpreadHostMouseUp);
                document.removeEventListener('mousedown', onMouseDown, true);
                document.removeEventListener('mouseup', onMouseUp)
            }
        }

    }, [spread, isWorkbookReadyToUse]);

    return <div class="sample-tutorial">
        <div class="sample-container">
            <div class="sample-spreadsheets">
                <SpreadSheets workbookInitialized={spread => initSpread(spread)}>
                    <Worksheet>
                    </Worksheet>
                </SpreadSheets>
            </div>
            <div id="statusBar"></div>
        </div>
        <div className="options-container">
            <div className="option-row">
                <div class="inputContainer">
                    <input id="selectedFile" type="file" accept=".sjs, .xlsx, .xlsm, .ssjson, .json, .csv" onChange={onSelectedFileChange} />
                    <button class="settingButton" id="open" onClick={open}>Open</button>

                </div>
                <div class="inputContainer">
                    <input type='text' id='search-text' style={{
                        border: '1px solid black',
                        marginRight: '10px'
                    }} />
                    <button class="settingButton" id="search" onClick={search}>Search</button>
                    <button class="settingButton" id="prev" style={{ marginRight: '8px' }} onClick={() => setCurrentHitIndex(currentHitIndex - 1)}>Previous Hit</button>
                    <button class="settingButton" id="next" onClick={() => setCurrentHitIndex(currentHitIndex + 1)}>Next Hit</button>
                    <button class="settingButton" id="extract" style={{ marginRight: '8px' }} onClick={extractData}>Extract Data</button>
                    <button class="settingButton" id="fitRowsToPageHeight" onClick={fitRowsToPageHeight}>Fit to Height</button>
                    <button class="settingButton" id="toggleShowHiddenData" onClick={() => {
                        showHiddenDataRef.current = !showHiddenDataRef.current;
                        toggleShowHiddenData(showHiddenDataRef.current)
                    }}>Show/Hide Data</button>
                </div>
            </div>
        </div>
    </div>;
}

function findMatches(str, words, matchWholeWord) {
    const flags = 'gi';
    const matches = [];

    for (const word of words) {
        const pattern = matchWholeWord ? `\\b${word}\\b` : word;
        const regex = new RegExp(pattern, flags);
        let match;

        while ((match = regex.exec(str)) !== null) {
            matches.push({ text: match[0], index: match.index });
        }
    }

    if (!matchWholeWord && matches.length === 0) {
        for (const word of words) {
            const splitWord = word.split(' '); // split word into individual words, for names with spaces in different columns
            if (splitWord.some(wStr => wStr.toLowerCase() === str.toLowerCase())) {
                matches.push({ text: str, index: 0 });
                break;
            }
        }
    }

    return matches.length > 0 ? matches.sort((a, b) => a.index - b.index) : matches;
}


// #808080
export const WordMarkingTabArray = ['#000000', '#0000FF', '#FF00FF', '#808080', '#008000', '#00FFFF', '#00FF00', '#800000',
    '#000080', '#808000', '#800080', '#FF6A00', '#C0C0C0', '#008080', '#FFFF00'];


export function getNonEmptyCols(sheet) {
    let json = sheet.toJSON()
    let dataTable = json.data.dataTable && Object.keys(json.data.dataTable);
    let nonEmptyColumns = [];
    dataTable && dataTable.forEach((row) => {
        let rowArray = Object.keys(json.data.dataTable[row]);
        rowArray.forEach((col) => {
            if (!nonEmptyColumns.includes(Number(col))) {
                nonEmptyColumns.push(Number(col));
            }
        });
    });
    return nonEmptyColumns.sort((a, b) => Number(a) - Number(b));
}

export function groupConsecutiveColumns(arr) {
    if (!arr.length) return [];

    arr.sort((a, b) => a - b); // Ensure the array is sorted
    const result = [];
    let start = arr[0];

    for (let i = 1; i < arr.length; i++) {
        if (arr[i] !== arr[i - 1] + 1) {
            result.push([start, arr[i - 1]]);
            start = arr[i];
        }
    }
    result.push([start, arr[arr.length - 1]]); // Push the last range

    return result;
}