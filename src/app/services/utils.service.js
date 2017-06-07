import angular from 'angular';
import * as XLSX from 'xlsx-style';

class Utils {
    constructor() {
    }

    createTableData(props) {
        let obj = {};

        obj.colData = props['name'];

        if (props['colSpan']) {
            obj.colSpan = props['colSpan']
        } else {
            obj.colSpan = 1;
        }

        if (props['rowSpan']) {
            obj.rowSpan = props['rowSpan']
        } else {
            obj.rowSpan = 1;
        }

        if (props['dataElement']) {
            obj.dataElement = props['dataElement']
        } else {
            obj.dataElement = '';
        }

        if (props['categoryOptionCombo']) {
            obj.categoryOptionCombo = props['categoryOptionCombo']
        } else {
            obj.categoryOptionCombo = '';
        }
        if (props['dataElementCell']) {
            obj.dataElementCell = props['dataElementCell']
        } else {
            obj.dataElementCell = false;
        }

        if (props['dataEntryCell']) {
            obj.dataEntryCell = props['dataEntryCell']
        } else {
            obj.dataEntryCell = false;
        }

        if (props['panelCell']) {
            obj.panelCell = props['panelCell']
        } else {
            obj.panelCell = false;
        }

        if (props['formulaCell']) {
            obj.formulaCell = props['formulaCell']
        } else {
            obj.formulaCell = false;
        }

        if (props['sheetName']) {
            obj.sheetName = props['sheetName']
        } else {
            obj.sheetName = '';
        }

        if (props['rows']) {
            obj.rows = props['rows']
        } else {
            obj.rows = '';
        }

        return obj
    }

    createTableRow(rowData) {
        let tr = [];
        for (let i = 0; i < rowData.length; i++) {
            tr.push(this.createTableData(rowData[i]));
        }
        return tr;
    }

    createDataRows(number, val, isDataElement, isDataEntry, isPanel) {
        let data = [];

        for (let i = 0; i < number; i++) {
            data.push({
                colSpan: 1,
                colData: val,
                rowSpan: 1,
                dataElementCell: isDataElement,
                dataEntryCell: isDataEntry,
                panelCell: isPanel
            });
        }
        return data;
    }

    s2ab(s) {
        let buf = new ArrayBuffer(s.length);
        let view = new Uint8Array(buf);
        for (let i = 0; i != s.length; ++i) {
            view[i] = s.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    dateToYMD(date) {
        let d = date.getDate();
        let m = date.getMonth() + 1;
        let y = date.getFullYear();
        return '' + y + '-' + (m <= 9 ? '0' + m : m) + '-' + (d <= 9 ? '0' + d : d);
    }

}


export default angular.module('services.utils', [])
    .service('Utils', Utils)
    .name;