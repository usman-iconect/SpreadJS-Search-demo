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

function mapExportFileType(fileType) {
    if (fileType === FileType.SSJson) {
        return GC.Spread.Sheets.FileType.ssjson;
    } else if (fileType === FileType.Csv) {
        return GC.Spread.Sheets.FileType.csv;
    }
    return GC.Spread.Sheets.FileType.excel;
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
    function initSpread(spread) {

        setSpread(spread);
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
    function save() {
        var fileType = saveFileType;
        var fileName = 'export.' + fileType;
        var options = deepClone(saveOptions[fileType]);

        if (fileType === FileType.SJS) {
            spread.save(function (blob) { saveAs(blob, fileName); }, function () { }, options);
        } else {
            options.fileType = mapExportFileType(fileType);
            spread.export(function (blob) { saveAs(blob, fileName); }, function () { }, options);
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


    function search(){
        console.log("searching", new Date().toLocaleTimeString())
        // let searchString = ["373087151310005", "778122350629261", "539604577512086", "570410512495429", "880898401883481", "855558342263732", "853530326251646", "823350331938527", "508858507779861", "1936650886647"];
        // let searchString = ['misfire', 'legal', 'claim', 'alleged', 'infection', 'design', 'bowel', 'device', 'records', 'erosion', 'patient', 'mesh', 'bard']
        let searchString = document.getElementById('search-text').value.split(" ")
        spread.suspendPaint();
        const activeSheet = spread.sheets[1]
        for (var i = 0; i < activeSheet.getRowCount() ; i++) {
            for (var j = 0; j < activeSheet.getColumnCount() ; j++)
            {
                var text = activeSheet.getText(i,j).toLowerCase();
                let isAHit = false
                for (const word of searchString) {
                    if (text.includes(word.toLowerCase().trim())) {
                        isAHit = true
                    }
                }
                if (isAHit) {
                    activeSheet.getCell(i, j).backColor("lightgreen");
                }
                else
                {
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
                    }}/>                  
                    <button class="settingButton" id="serach" onClick={search}>Search</button>

                </div>
            </div>
        </div>
    </div>;
}