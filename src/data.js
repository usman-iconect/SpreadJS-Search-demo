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