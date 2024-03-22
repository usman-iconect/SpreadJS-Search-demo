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
    const [openOptions, setOpenOptions] = React.useState({
        sjs: {},
        ssjson: {},
        xlsx: {},
        csv: {},
    });
    const [saveOptions, setSaveOptions] = React.useState({
        sjs: {},
        ssjson: {},
        xlsx: {},
        csv: {},
    });

    function initSpread(spread) {

        setSpread(spread);
        spread.contextMenu.menuData = [
            {
                text: "New Person",
                name: "newPerson",
                command: () => {
                    console.log("new person clicked")
                },
                workArea: "viewport"
            },
            {
                text: "Current",
                name: "Current",
                command: () => {
                    console.log("Current clicked")
                },
                workArea: "viewport"
            },
            {
                text: "Copy",
                name: "Copy",
                command: () => {
                    console.log("Copy clicked")
                },
                workArea: "viewport"
            },
            {
                text: "Close",
                name: "Close",
                command: () => {
                    console.log("close clicked")
                },
                workArea: "viewport"
            },
        ]

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

    function HighlightText(searchResults, cellText, row, col, activeSheet, spread) {
        const highlightCommand = {
            canUndo: true,
            execute: function (spread, options, isUndo) {
                const Commands = GC.Spread.Sheets.Commands;
                if (isUndo) {
                    Commands.undoTransaction(spread, options);
                    return true;
                } else {
                    Commands.startTransaction(spread, options);
                    // activeSheet.getCell(row, col).backColor("green")

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
                            cellContent.richText.push({ style: { foreColor: "#FFDF00" }, text: cellText.substring(result.index, lastIndex) });
                        });
                        if (lastIndex < cellText.length) {
                            const remainingText = cellText.substring(lastIndex);
                            cellContent.richText.push({ text: remainingText });
                        }
                        activeSheet.setValue(row, col, cellContent);
                    }

                    //hover effect
                    activeSheet.comments.add(row, col, containsNumber(cellText) ? "SSN" : "Person");
                    const activeComment = activeSheet.comments.get(row, col)
                    activeComment.width(80)
                    activeComment.height(35)
                    activeComment.fontSize('14' + "pt");
                    activeComment.fontWeight('bold');
                    activeComment.borderWidth(0);
                    activeComment.backColor('#FFDF00');
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
        // let searchString = ["373087151310005", "778122350629261", "539604577512086", "570410512495429", "880898401883481", "855558342263732", "853530326251646", "823350331938527", "508858507779861", "1936650886647"];
        // let searchString = ['misfire', 'legal', 'claim', 'alleged', 'infection', 'design', 'bowel', 'device', 'records', 'erosion', 'patient', 'mesh', 'bard']
        let searchStrings = document.getElementById('search-text').value;
        if (!searchStrings || searchStrings.length === 0) {
            searchStrings = 'Hall,Smith,Boyd,Trevor,Curtis,Brian,859-86-8326,211-43-1582,713-62-9309';
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
                    HighlightText(searchResults, text, i, j, activeSheet, spread);
                }
            }
        }
        spread.resumePaint();
        console.log("done", new Date().toLocaleTimeString())

    }

    function openExportTab() {
        // Open a new tab
        const exportTab = window.open('/viewer.html', '_blank');

        // Listen for messages from the export tab
        window.addEventListener('message', (event) => {
            console.log("ecent coming", event)
            if (event.origin !== window.location.origin) return;

            // Handle messages from the export tab
            if (event.data === 'exportCompleted') {
                // Do something when export is completed, e.g., close the export tab
                exportTab.close();
            }
        });
    }

    React.useEffect(() => {
        if (spread)
            fetch('00000029.xlsx')
                .then(res => res.blob())
                .then((blob) => {
                    const file = new File([blob], 'excel.xlsx', { type: blob.type });
                    spread.import(file, () => {
                        search();
                    }, (error) => {
                        console.log('error', error);
                    });
                });
    }, [spread]);

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
                    <button class="settingButton" id="serach" onClick={search}>Search</button>

                </div>
                <button class="settingButton" id="serach" onClick={() => {
                    spread.savePDF(function (blob) {

                        console.log(blob)

                    }, function ({ errorMessage }) {
                        console.log(errorMessage);
                    }, {
                        title: 'Test Title',
                        author: 'Test Author',
                        subject: 'Test Subject',
                        keywords: 'Test Keywords',
                        creator: 'test Creator'
                    });
                }}>Export PDF</button>
                <button style={{ marginLeft: 10 }} class="settingButton" id="serach" onClick={() => {
                    spread.suspendPaint();
                    const sheet = spread.sheets[1];
                    const printInfo = sheet.printInfo();
                    printInfo.showBorder(false);
                    printInfo.showGridLine(true);
                    spread.resumePaint();
                    setTimeout(() => {
                        spread.print()
                    }, 0)
                }}>Print</button>
                <button style={{ marginLeft: 10 }} class="settingButton" onClick={openExportTab}>New Tab Print</button>
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

    return matches.length > 0 ? matches.sort((a, b) => a.index - b.index) : matches;
}


// #808080
export const WordMarkingTabArray = ['#000000', '#0000FF', '#FF00FF', '#808080', '#008000', '#00FFFF', '#00FF00', '#800000',
    '#000080', '#808000', '#800080', '#FF6A00', '#C0C0C0', '#008080', '#FFFF00'];

function containsNumber(str) {
    const numberRegex = /\d/;
    return numberRegex.test(str);
}