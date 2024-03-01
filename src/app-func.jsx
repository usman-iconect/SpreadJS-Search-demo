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
function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
}

function getElementId(mode, fileType, propName) {
    return mode + '-' + fileType + '-' + propName;
}

function getFileType(file) {
    if (!file) {
        return;
    }

    var fileName = file.name;
    var extensionName = fileName.substring(fileName.lastIndexOf(".") + 1);

    if (extensionName === 'sjs') {
        return FileType.SJS;
    } else if (extensionName === 'xlsx' || extensionName === 'xlsm') {
        return FileType.Excel;
    } else if (extensionName === 'ssjson' || extensionName === 'json') {
        return FileType.SSJson;
    } else if (extensionName === 'csv') {
        return FileType.Csv;
    }
}


var defaultOpenOptions = [
    { propName: "openMode", type: "select", displayText: "OpenMode", options: [{ name: 'normal', value: 0 }, { name: 'lazy', value: 1 }, { name: 'incremental', value: 2 }], default: 0 },
    { propName: "includeStyles", type: "boolean", default: true },
    { propName: "includeFormulas", type: "boolean", default: true },
    { propName: "fullRecalc", type: "boolean", default: false },
    { propName: "dynamicReferences", type: "boolean", default: true },
    { propName: "calcOnDemand", type: "boolean", default: false },
    { propName: "includeUnusedStyles", type: "boolean", default: true },
];

var importXlsxOptions = [
    { propName: "openMode", type: "select", displayText: "OpenMode", options: [{ name: 'normal', value: 0 }, { name: 'lazy', value: 1 }, { name: 'incremental', value: 2 }], default: 0 },
    { propName: "includeStyles", type: "boolean", default: true },
    { propName: "includeFormulas", type: "boolean", default: true },
    { propName: "frozenColumnsAsRowHeaders", type: "boolean", default: false },
    { propName: "frozenRowsAsColumnHeaders", type: "boolean", default: false },
    { propName: "fullRecalc", type: "boolean", default: false },
    { propName: "dynamicReferences", type: "boolean", default: true },
    { propName: "calcOnDemand", type: "boolean", default: false },
    { propName: "includeUnusedStyles", type: "boolean", default: true },
];

var importSSJsonOptions = [
    { propName: "includeStyles", type: "boolean", default: true },
    { propName: "includeFormulas", type: "boolean", default: true },
    { propName: "frozenColumnsAsRowHeaders", type: "boolean", default: false },
    { propName: "frozenRowsAsColumnHeaders", type: "boolean", default: false },
    { propName: "fullRecalc", type: "boolean", default: false },
    { propName: "incrementalLoading", type: "boolean", default: false }
];

var importCsvOptions = [
    { propName: "encoding", type: "string", default: "UTF-8" },
    { propName: "rowDelimiter", type: "string", default: "\r\n" },
    { propName: "columnDelimiter", type: "string", default: "," }
];


var FileType = {
    SJS: 'sjs',
    Excel: 'xlsx',
    SSJson: 'ssjson',
    Csv: 'csv',
}

export function AppFunc() {
    const [spread, setSpread] = React.useState(null);
    const [selectedFile, setSelectedFile] = React.useState(null);
    const [openFileType, setOpenFileType] = React.useState('');
    const [saveFileType, setSaveFileType] = React.useState(FileType.SJS);
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

    function newPerson() {
        console.log("new person clicked")
    }

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
    }
    function open() {
        var file = selectedFile;
        if (!file) {
            return;
        }

        var fileType = getFileType(file);
        var options = deepClone(openOptions[fileType]);

        if (fileType === FileType.SJS) {
            spread.open(file, function () { }, function () { }, options);
        } else {
            spread.import(file, function () { }, function () { }, options);
        }
    }

    function onSelectedFileChange(e) {
        let selectedFile = e.target.files[0];
        let openFileType = getFileType(selectedFile);
        setSelectedFile(selectedFile);
        setOpenFileType(openFileType)
    }
    function onSaveFileTypeChange(e) {
        let saveFileType = e.target.value;
        setSaveFileType(saveFileType);
    }
    function onPropChange(mode, fileType, prop, e) {
        let element = e.target;

        var value;
        if (prop.type === 'boolean') {
            value = element.checked;
        } else if (prop.type === 'number') {
            value = +element.value;
        } else if (prop.type === 'string') {
            value = element.value;
        } else if (prop.type === 'select') {
            value = +element.value;
        }

        if (mode === 'open') {
            openOptions[fileType][prop.propName] = value;
            setOpenOptions(openOptions);
        } else if (mode === 'save') {
            saveOptions[fileType][prop.propName] = value;
            setSaveOptions(saveOptions);
        }
    }
    function createOptions(options, fileType, mode) {
        let selectFileType = mode === 'open' ? openFileType : saveFileType;
        let display = selectFileType === fileType ? '' : 'none';

        return <div className={fileType} style={{ display }}>
            {options.map((prop) => createProp(mode, fileType, prop))}
        </div>;
    }
    function createProp(mode, fileType, prop) {
        let id = getElementId(mode, fileType, prop.propName);

        if (prop.type === 'select') {
            return <item className='item'>
                <label for={id}>{prop.displayText || prop.propName}</label>
                <select id={id} defaultValue={prop.default} onChange={(e) => onPropChange(mode, fileType, prop, e)}>
                    {prop.options.map((p) => <option value={p.value}>{p.name}</option>)}
                </select>
            </item>
        } else if (prop.type === 'boolean') {
            return <item className='item'>
                <input id={id} type='checkbox' defaultChecked={prop.default} onChange={(e) => onPropChange(mode, fileType, prop, e)}></input>
                <label for={id}>{prop.displayText || prop.propName}</label>
            </item>
        } else if (prop.type === 'number') {
            return <item className='item'>
                <label for={id}>{prop.displayText || prop.propName}</label>
                <input id={id} type='number' defaultValue={prop.default} onChange={(e) => onPropChange(mode, fileType, prop, e)}></input>
            </item>
        } else {
            return <item className='item'>
                <label for={id}>{prop.displayText || prop.propName}</label>
                <input id={id} type='text' defaultValue={prop.default} onChange={(e) => onPropChange(mode, fileType, prop, e)}></input>
            </item>
        }
    }

    function setCellTypeCallback(hyperLink, data) {
        // h.text("set sheet tab style");
        hyperLink.linkToolTip(data.text.includes("@") ? "email" : "Person");
        hyperLink.activeOnClick(true);
        hyperLink.linkColor("black");
        hyperLink.onClickAction(() => {
            console.log("cell clicked", data);
        });
        return hyperLink;
    }

    function search() {
        console.log("searching", new Date().toLocaleTimeString())
        // let searchString = ["373087151310005", "778122350629261", "539604577512086", "570410512495429", "880898401883481", "855558342263732", "853530326251646", "823350331938527", "508858507779861", "1936650886647"];
        // let searchString = ['misfire', 'legal', 'claim', 'alleged', 'infection', 'design', 'bowel', 'device', 'records', 'erosion', 'patient', 'mesh', 'bard']
        let searchString = document.getElementById('search-text').value
        const activeSheet = spread.getActiveSheet()
        const range = activeSheet.getUsedRange(GC.Spread.Sheets.UsedRangeType.data);
        if (!range) {
            return;
        }

        spread.suspendPaint();

        for (var i = range.row; i < range.row + range.rowCount; i++) {
            for (var j = range.col; j < range.col + range.colCount; j++) {
                const text = activeSheet.getText(i, j);
                if (text == searchString || text.includes(searchString)) {
                    highlightText(searchString, text, i, j, activeSheet);
                }
                else if (text != searchString) {
                    activeSheet.getCell(i, j).backColor(undefined);
                }

            }

        }
        spread.resumePaint();
        console.log("done", new Date().toLocaleTimeString())
    }

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

                    <div class="open-options">
                        {[
                            createOptions(defaultOpenOptions, FileType.SJS, 'open'),
                            createOptions(importXlsxOptions, FileType.Excel, 'open'),
                            createOptions(importSSJsonOptions, FileType.SSJson, 'open'),
                            createOptions(importCsvOptions, FileType.Csv, 'open'),
                        ]}
                    </div>
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
                <button style={{marginLeft: 10}} class="settingButton" id="serach" onClick={() => {
                    setTimeout(() => {
                        spread.print()
                    }, 0)
                }}>Print</button>

            </div>
        </div>
    </div>;
}

function highlightText(searchString, text, row, col, activeSheet) {

    if (text.length == searchString.length) {
        activeSheet.getCell(row, col).foreColor("rgb(252, 28, 3)");
    } else {
        const normalText = text.split(searchString);
        const cellContent = { richText: [] };

        for (let i = 0; i < normalText.length; i++) {
            if (i == 0 && normalText[i] == "") {
                cellContent.richText.push({ style: { foreColor: "rgb(252, 28, 3)" }, text: searchString });
                continue;
            } else if (i == (normalText.length - 1)) {
                if (normalText[i] == "") {
                    break;
                } else {
                    cellContent.richText.push({ text: normalText[i] });
                }
            } else {
                cellContent.richText.push({ text: normalText[i] });
                cellContent.richText.push({ style: { foreColor: "rgb(252, 28, 3)" }, text: searchString });
            }

        }
        activeSheet.setValue(row, col, cellContent);
    }

}