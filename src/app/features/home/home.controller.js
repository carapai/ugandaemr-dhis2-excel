import * as xlsx from "xlsx";
import * as xlsxs from "xlsx-style";
import saveAs from "save-as";
export default class HomeController {
    constructor($scope, Data, Restangular, Filters, Utils, $uibModal, dataSets) {

        this.items = [];
        this.dataSets = Restangular.stripRestangular(dataSets);
        this.filters = Filters;
        this.utils = Utils;
        this.uimodal = $uibModal;
        let d = new Date();
        this.yearValue = d.getFullYear();
        this.periodType = "Quarterly";

        this.data = Data;

        this.restagular = Restangular;

        this.elements = [];
        this.cellsGot = [];

        this.selectedDataset = null;
        this.selectedDatasetCategories = null;
        this.selectedPeriod = null;
        this.selectedDatasetCategories = null;
        this.selectedOrganisationUnit = null;
        this.showDatasetCategories = false;
        this.datasetCategories = null;
        this.form = null;
        this.excel = null;
        this.createdTable = null;

        if (!this.dataSets) {
            let modalInstance = this.uimodal.open({
                animation: true,
                ariaLabelledBy: 'modal-title',
                ariaDescribedBy: 'modal-body',
                template: require('./alert-modal.html'),
                controller: 'ModalController',
                controllerAs: 'alert',
                size: 'sm',
                backdrop: false,
                resolve: {
                    items: function () {
                        return "You have not been assigned any datasets, contact system administrator";
                    }
                }
            });
            modalInstance.result.then(() => {
            }, () => {
            });
        }

        this.panelStyling = {
            "fill": {
                "fgColor": {
                    "rgb": "3c3c3c"
                }
            },
            "font": {
                "name": "Times New Roman",
                bold: true,
                italic: true,
                outline: true,
                shadow: true,
                vertAlign: "superscript",
                "sz": 24,
                "color": {
                    "rgb": "FFFF00"
                }
            }
        };

        this.dataElementStyling = {
            "alignment": {
                "horizontal": "left",
                "vertical": "center",
                "wrapText": 1
            }
        };


        this.dataEntryStyling = {
            "border": {
                "left": {
                    "style": "thin",
                    "color": {
                        "auto": 1
                    }
                },
                "right": {
                    "style": "thin",
                    "color": {
                        "auto": 1
                    }
                },
                "top": {
                    "style": "thin",
                    "color": {
                        "auto": 1
                    }
                },
                "bottom": {
                    "style": "thin",
                    "color": {
                        "auto": 1
                    }
                }
            }
        };

        this.headersStyling = {
            "alignment": {
                "horizontal": "center",
                "vertical": "center",
                "wrapText": 1
            },
            "fill": {
                "fgColor": {
                    "rgb": "eaf7fb"
                }
            }
        };

        this.monthNames = [
            "January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
            "November", "December"
        ];

        $scope.$watch(() => this.excel, (newVal) => {
                if (newVal) {
                    this.cellsGot = [];
                    this.wb = xlsx.read(newVal.base64, {
                        type: 'base64',
                        WTF: false
                    });

                    let work_sheet = this.wb.Sheets["Main"];
                    let properties_work_sheet = this.wb.Sheets["Properties"];
                    let props_sheet = xlsx.utils.sheet_to_json(properties_work_sheet);

                    let datasetString = _.filter(props_sheet, (d) => {
                        return _.has(d, 'dataSet');
                    })[0].dataSet;

                    let organizationCell = _.filter(props_sheet, (d) => {
                        return _.has(d, 'organization');
                    })[0].organization;

                    let periodCell = _.filter(props_sheet, (d) => {
                        return _.has(d, 'period');
                    })[0].period;

                   /* this.dataSetCategoryCombo = _.filter(props_sheet, (d) => {
                        return _.has(d, 'categoryCombination');
                    })[0].categoryCombination;*/

                    let datasetCategoryOptionCells = _.filter(props_sheet, (d) => {
                        return _.has(d, 'category');
                    });

                    let dataCells = _.filter(props_sheet, (d) => {
                        return _.has(d, 'entryCell');
                    });

                    this.importedOrganisationUnit = work_sheet[organizationCell].v;
                    this.importedPeriod = work_sheet[periodCell].v;

                    let datasetCategoryOptionsValues = _.map(datasetCategoryOptionCells, (datasetCategoryOptionCell) => {
                        return work_sheet[datasetCategoryOptionCell.category].v;
                    });

                    if (!this.importedPeriod) {
                        let dt = xlsx.SSF.parse_date_code(period, {
                            date1904: false
                        });
                        this.importedPeriod = dt.m < 10 ? dt.y + '0' + dt.m : dt.y + '' + dt.m;
                    }

                    this.selectedDatasetCategories = datasetCategoryOptionsValues.join(",");
                    this.importedDataset = datasetString.split(',')[1];
                    this.importedDatasetId = datasetString.split(',')[0];

                    _.forEach(dataCells, (dataCell) => {
                        let cells = dataCell.entryCell.split('-');
                        this.cellsGot = [...this.cellsGot, {
                            cell: cells[0],
                            dataElement: cells[1],
                            categoryOptionCombo: cells[2],
                            cellValue: work_sheet[cells[0]]
                        }];
                    });

                    Data.getOne('organisationUnits', this.importedOrganisationUnit).then((orgUnit) => {
                        this.realOrganizationUnit = Restangular.stripRestangular(orgUnit);
                    }, (error) => {
                        let modalInstance = this.uimodal.open({
                            animation: true,
                            ariaLabelledBy: 'modal-title',
                            ariaDescribedBy: 'modal-body',
                            template: require('./alert-modal.html'),
                            controller: 'ModalController',
                            controllerAs: 'alert',
                            size: 'sm',
                            backdrop: false,
                            resolve: {
                                items: function () {
                                    return "Organization Unit UID Specified In The Excel Not Found, Please Correct It Before You Can Continue";
                                }
                            }
                        });
                        modalInstance.result.then(() => {
                            this.cellsGot = []
                        }, () => {
                        });
                    });

                    this.data.getOne('dataSets', this.importedDatasetId, {fields: 'dataSetElements[dataElement[id,name,displayName,categoryCombo[id,name,uuid,displayName,categoryOptionCombos[id,name,displayName,categoryCombo[id,name,displayName],categoryOptions[id,name,displayName]],categories[id,name,displayName,categoryCombos[id,name,displayName],categoryOptions[id,name,uuid,displayName]]]]]'}).then((dataSet) => {
                        let elements = _.map(this.restagular.stripRestangular(dataSet)['dataSetElements'], 'dataElement');
                        this.dataElementsFound = _.groupBy(elements, 'id');
                        let categoryOptionCombos = [];
                        _.forEach(elements, (element) => {
                            categoryOptionCombos = [...categoryOptionCombos, ...element.categoryCombo.categoryOptionCombos]
                        });
                        this.categoryOptionCombosFound = _.groupBy(_.uniqBy(categoryOptionCombos, 'id'), 'id');
                    });
                }
            }
        );
    }

    Workbook() {
        this.SheetNames = [];
        this.Sheets = {};
        this.Custprops = {};
    }

    open(insertedRecords) {
        let modalInstance = this.uimodal.open({
            animation: true,
            ariaLabelledBy: 'modal-title',
            ariaDescribedBy: 'modal-body',
            template: require('./modal.html'),
            controller: 'ModalController',
            controllerAs: 'alert',
            size: 'sm',
            backdrop: false,
            resolve: {
                items: function () {
                    return insertedRecords;
                }
            }
        });
        modalInstance.result.then(function () {
        }, function () {
        });
    }

    showOrganizationUnits() {
        this.elements = [];
        this.selectedPeriod = null;
        this.selectedDatasetCategories = null;
        this.selectedOrganisationUnit = null;
        this.table = null;
        this.showDatasetCategories = false;
        this.datasetCategories = null;
        this.form = null;
        this.createdTable = null;
        this.periodType = this.selectedDataset.periodType;

        this.selectedDataset.organisationUnits = [_.head(this.selectedDataset.organisationUnits)];

        this.data.getOne('dataSets/' + this.selectedDataset.id, 'form', {
            ou: this.selectedDataset.organisationUnits[0].id,
            metaData: true
        }).then((form) => {
            this.form = this.restagular.stripRestangular(form);
        });

        this.data.getOne('dataSets', this.selectedDataset.id, {fields: 'dataSetElements[dataElement[id,name,displayName,categoryCombo[id,name,uuid,displayName,categoryOptionCombos[id,name,displayName,categoryCombo[id,name,displayName],categoryOptions[id,name,displayName]],categories[id,name,displayName,categoryCombos[id,name,displayName],categoryOptions[id,name,uuid,displayName]]]]]'}).then((dataSet) => {
            this.selectedDataset.dataElements = _.map(this.restagular.stripRestangular(dataSet)['dataSetElements'], 'dataElement');
            let categoryOptionCombos = [];
            let categoryOptions = [];

            _.forEach(this.selectedDataset.dataElements, (element) => {
                categoryOptionCombos = [...categoryOptionCombos, ...element.categoryCombo.categoryOptionCombos]
            });
            _.forEach(_.uniqBy(categoryOptionCombos, 'id'), (categoryOptionCombo) => {
                categoryOptions = [...categoryOptions, ...categoryOptionCombo.categoryOptions]
            });

            this.selectedDataset.categoryOptions = _.uniqBy(categoryOptions, 'id');
        });
    }

    searchName(list, name) {
        let found = false;
        for (let i = 0; i < list.length; i++) {
            if (list[i].name.toLowerCase().trim().indexOf(name.toLowerCase().trim()) >= 0) {
                found = true;
                break;
            }
        }
        return found;
    }

    showPeriods() {
        this.selectedDatasetCategories = null;
        this.table = null;
        this.tableRows = null;
        this.showDatasetCategories = false;
        this.datasetCategories = null;
        this.selectedPeriod = null;
        this.createdTable = null;
        this.getPeriodArray();
    }

    getPeriodArray() {
        this.dataPeriods = this.filters.getPeriods(this.periodType, this.yearValue);
    }

    nextYear() {
        this.yearValue = parseInt(this.yearValue) + 1;
        this.getPeriodArray();
    }

    previousYear() {
        this.yearValue = parseInt(this.yearValue) - 1;
        this.getPeriodArray();
    }

    getText(el) {
        let text = '';
        if (el.hasChildNodes()) {
            for (let i = 0, l = el.childNodes.length; i < l; i++) {
                if (el.childNodes[i].nodeType === Element.TEXT_NODE) {
                    text += el.childNodes[i].nodeValue;
                }
            }
            return text.trim();
        }
    }

    showOthers() {
        let filteredCategoryOptions = [];
        _.forEach(this.categoryOptions, (catOption) => {
            let userGroupAccesses = _.map(catOption.userGroupAccesses, 'id');
            if (this.userGroups.length > 0 && userGroupAccesses.length > 0 && (_.intersection(this.userGroups, userGroupAccesses)).length > 0 && this.checkDate(this.periodType, this.selectedPeriod.id, catOption.startDate, catOption.endDate)) {
                filteredCategoryOptions.push(catOption.id)
            }
        });

        this.filteredCategoryOptions = filteredCategoryOptions;
        if (this.form.categoryCombo) {
            _.forEach(this.form.categoryCombo.categories, (category) => {
                if (category.label === 'Project') {
                    let options = _.remove(category.categoryOptions, (co) => {
                        return _.indexOf(this.filteredCategoryOptions, co.id) !== -1;
                    });
                    category.categoryOptions = options;
                }
            });

            this.datasetCategories = this.form.categoryCombo.categories;
            this.showDatasetCategories = true;
            this.selectedDatasetCategories = null;
            this.table = null;
            this.tableRows = null;

        } else {
            this.showDataSets();
        }
    }

    removeComments(html) {
        return ('' + html)
            .replace/*HTMLComments*/(/<!-[\S\s]*?-->/gm, '')
            .replace/*JSBlockComments*/(/\/\*[\S\s]*?\*\//gm, '')
            .replace/*JSLineComments*/(/^.*?\/\/.*/gm, '$1');
    }

    removeJs(html) {
        let SCRIPT_REGEX = /<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi;
        while (SCRIPT_REGEX.test(html)) {
            html = html.replace(SCRIPT_REGEX, "");
        }

        return html;
    }

    removeCss(html) {
        let STYLE_REGEX = /<style\b[^<]*(?:(?!<\/style>)<[^<]*)*<\/style>/gi;
        while (STYLE_REGEX.test(html)) {
            html = html.replace(STYLE_REGEX, "");
        }

        return html;
    }

    removeTag(html, tag) {
        let element = html.getElementsByTagName(tag), index;

        for (index = element.length - 1; index >= 0; index--) {
            element[index].parentNode.removeChild(element[index]);
        }

        return html;
    }

    removeByClass(html, className) {
        let element = html.getElementsByClassName(className), index;
        for (index = element.length - 1; index >= 0; index--) {
            element[index].parentNode.removeChild(element[index]);
        }
        return html;
    }

    clean(node) {
        for (let n = 0; n < node.childNodes.length; n++) {
            let child = node.childNodes[n];
            if
            (
                child.nodeType === 8
                ||
                (child.nodeType === 3 && !/\S/.test(child.nodeValue))
            ) {
                node.removeChild(child);
                n--;
            }
            else if (child.nodeType === 1) {
                this.clean(child);
            }
        }

        return node;
    }

    createTableHtml(table) {
        let tableRows = [];
        for (let i = 0, row; row = table.rows[i]; i++) {
            let tds = [];
            for (let j = 0, col; col = row.cells[j]; j++) {
                try {
                    let foundObject = angular.fromJson(col.innerText);
                    if (foundObject.id.endsWith('-val')) {
                        let otherFields = foundObject.id.split("-");
                        tds = [...tds, {
                            name: '',
                            colSpan: col.colSpan || 1,
                            rowSpan: col.rowSpan || 1,
                            dataEntryCell: true,
                            dataElement: otherFields[0],
                            categoryOptionCombo: otherFields[1]
                        }]
                    } else {
                        tds = [...tds, {
                            name: '',
                            colSpan: col.colSpan || 1,
                            rowSpan: col.rowSpan || 1,
                            dataEntryCell: false
                        }];
                    }
                } catch (e) {
                    if (j === 0) {
                        tds = [...tds, {
                            name: col.innerText.replace(/<(?:.|\n)*?>/gm, ''),
                            colSpan: col.colSpan || 1,
                            rowSpan: col.rowSpan || 1,
                            dataElementCell: true
                        }]
                    } else {
                        tds = [...tds, {
                            name: col.innerText.replace(/<(?:.|\n)*?>/gm, ''),
                            colSpan: col.colSpan || 1,
                            rowSpan: col.rowSpan || 1
                        }]
                    }
                }
            }
            tableRows = [...tableRows, this.utils.createTableRow(tds)]
        }
        return tableRows;
    }

    findElements(node) {
        for (let n = 0; n < node.childNodes.length; n++) {
            let child = node.childNodes[n];
            if (child.tagName === 'TABLE') {
                for (let i = 0, row; row = child.rows[i]; i++) {
                    for (let j = 0, col; col = row.cells[j]; j++) {
                        let input = col.querySelector("input,textarea,select,radio");
                        if (input !== null) {
                            let inputAttributes = {};
                            for (let i = 0; i < input.attributes.length; i++) {
                                let attrib = input.attributes[i];
                                inputAttributes[attrib.name] = attrib.value;
                            }
                            col.innerHTML = angular.toJson(inputAttributes);
                        }
                    }
                }
                this.elements = [...this.elements, {
                    val: child
                }]
            } else if (child.hasChildNodes()) {
                this.findElements(child);
            } else {
                let val = child.nodeValue;
                if (val !== null) {
                    if (String(val).indexOf("$") === -1) {

                        let table = document.createElement('table');
                        let tr = document.createElement('tr');

                        tr.appendChild(document.createElement('td'));
                        tr.cells[0].appendChild(document.createTextNode(val));

                        table.appendChild(tr);

                        this.elements = [...this.elements, {val: table}];
                    }
                }
            }
        }
    }

    createHeaders(maximumLength) {
        let headers = [];
        headers = [...headers, this.utils.createTableRow([{
            name: 'This is an automatically created template.  Do not edit or change the layout',
            colSpan: 1,
            panelCell: true
        }])];

        headers = [
            ...headers,
            this.utils.createTableRow([{name: 'Dataset', css: '', dataElementCell: true}, {
                name: this.selectedDataset.displayName,
                colSpan: 1,
                dataElementCell: true
            }])];

        headers = [...headers, this.utils.createTableRow([
            {
                name: 'Organization',
                css: '',
                dataElementCell: true
            }, {
                name: this.selectedOrganisationUnit.displayName + '-' + this.selectedOrganisationUnit.id + '-Organizations',
                colSpan: 1,
                dataEntryCell: true,
                formulaCell: true
            }])
        ];
        if (this.selectedPeriod) {
            headers = [...headers, this.utils.createTableRow([
                {
                    name: 'Period',
                    css: '', dataElementCell: true,
                },
                {
                    name: this.selectedPeriod.name + '-' + this.selectedPeriod.id + '-Periods',
                    colSpan: 1,
                    dataEntryCell: true,
                    formulaCell: true
                }
            ])];
        }
        if (this.datasetCategories) {
            _.forEach(this.datasetCategories, (val) => {
                let label = this.selectedDatasetCategories[val.label].label;
                headers = [...headers, this.utils.createTableRow([
                    {
                        name: val.label,
                        dataElementCell: true,
                    },
                    {
                        name: ((label.split("(")[0]).split(":")[0]).trim() + '-' + this.selectedDatasetCategories[val.label].id + '-' + val.label + '-DatasetCategories',
                        colSpan: 1,
                        formulaCell: true,
                        dataEntryCell: true
                    }
                ])];
            });
        }
        return headers;
    }

    showDataSets() {
        this.createdTable = null;
        this.elements = [];
        if (this.selectedDataset.dataEntryForm) {
            let div = document.createElement('div');

            let html = this.selectedDataset.dataEntryForm.htmlCode;

            div.innerHTML = this.removeCss(this.removeJs(this.removeComments(html)));

            div = this.clean(this.removeTag(this.removeTag(this.removeTag(div, "link"), "meta"), "title"));
            div = this.removeByClass(this.removeByClass(div, "hidden"), "input-group");

            this.findElements(div);
            this.elements = _.map(this.elements, (element) => {
                let current = element.val;
                return {val: this.createTableHtml(current), el: current};
            });
            let headers = this.createHeaders();

            let tableEl = document.createElement("table");

            for (let i = 0; i < headers.length; i++) {
                let rowEl = tableEl.insertRow();
                let currentRow = headers[i];
                for (let j = 0; j < currentRow.length; j++) {
                    let cell = rowEl.insertCell();
                    cell.textContent = currentRow[j].colData;
                }
            }

            let x = document.createElement("TABLE");
            x.append(tableEl.tBodies[0]);
            _.forEach(this.elements, (element) => {
                _.forEach(element.el.tBodies, (tBody) => {
                    x.append(tBody)
                });
            });
            this.createdTable = x;
        } else {
            let tableRows = [];
            let fields = [];

            _.forEach(this.form.groups, (group) => {
                _.forEach(group.fields, (field) => {
                    fields.push(field);
                });
            });

            this.fields = _.groupBy(fields, 'dataElement');

            let maximumLength = 0;

            _.forEach(this.fields, (field) => {
                if (field.length > maximumLength) {
                    maximumLength = field.length;
                }
            });

            this.maximumLength = maximumLength;

            tableRows.push(this.utils.createTableRow([{
                name: 'This is an automatically created template.  Do not edit or change the layout',
                colSpan: maximumLength + 1,
                panelCell: true
            }]));

            tableRows.push(this.utils.createTableRow([{name: 'Dataset', css: ''}, {
                name: this.selectedDataset.displayName,
                colSpan: maximumLength,
                dataElementCell: true
            }]));
            tableRows.push(this.utils.createTableRow([
                {
                    name: 'Organization',
                    css: ''
                }, {
                    name: this.selectedOrganisationUnit.displayName + '-' + this.selectedOrganisationUnit.id + '-Organizations',
                    colSpan: maximumLength,

                    dataEntryCell: true
                }
            ]));
            if (this.selectedPeriod) {
                tableRows.push(this.utils.createTableRow([
                    {
                        name: 'Period',
                        css: ''
                    },
                    {
                        name: this.selectedPeriod.name + '-' + this.selectedPeriod.id + '-Periods',
                        colSpan: maximumLength,
                        dataEntryCell: true
                    }
                ]));
            }
            if (this.datasetCategories) {
                _.forEach(this.datasetCategories, (val) => {
                    let label = this.selectedDatasetCategories[val.label].label;
                    tableRows.push(this.utils.createTableRow([
                        {
                            name: val.label
                        },
                        {
                            name: ((label.split("(")[0]).split(":")[0]).trim() + '-' + this.selectedDatasetCategories[label].id + '-' + label + '-DatasetCategories',
                            colSpan: maximumLength,
                            dataEntryCell: true
                        }
                    ]));
                });
            }
            // Group DataElements based the CategoryComboId
            let categoryCombos = _.groupBy(this.selectedDataset.dataElements, 'categoryCombo.id');
            // Loop through the grouped DataElements
            _.forEach(categoryCombos, (dataElements) => {
                let cats = this.processCategories(dataElements[0].categoryCombo.categories);
                _.forEach(cats, (category) => {
                    const opts = category.categoryOptions;
                    let total = _.reduce(opts, (sum, n) => {
                        if (!n.colSpan) {
                            n.colSpan = 1;
                        }
                        return sum + n.colSpan;
                    }, 0);
                    tableRows.push(_.concat(this.utils.createTableRow([{
                        name: '',
                        colSpan: (maximumLength - total) + 1
                    }]), this.utils.createTableRow(category.categoryOptions)));
                });

                _.forEach(dataElements, (dataElement) => {
                    let dataValueCells = this.fields[dataElement.id];
                    dataElement.colSpan = (maximumLength - dataValueCells.length) + 1;
                    dataElement.dataElementCell = true;
                    dataElement.name = dataElement.displayName;

                    _.forEach(dataValueCells, (dataValueCell) => {
                        dataValueCell.dataEntryCell = true;
                        dataValueCell.name = '';
                    });
                    const anotherArray = _.concat([dataElement], dataValueCells);
                    tableRows.push(this.utils.createTableRow(anotherArray));
                });
            });

            _.forEach(tableRows, (row) => {
                row[0].css = 'nrcindicatorName';
                row[0].dataElementCell = true;
                row[0].dataEntryCell = false;
            });

            let tableEl = document.createElement("table");

            for (let i = 0; i < tableRows.length; i++) {
                let rowEl = tableEl.insertRow();
                let currentRow = tableRows[i];
                for (let j = 0; j < currentRow.length; j++) {
                    let cell = rowEl.insertCell();
                    cell.colSpan = currentRow[j].colSpan;
                    cell.rowSpan = currentRow[j].rowSpan;
                    if (currentRow[j].dataElement && currentRow[j].categoryOptionCombo) {
                        cell.textContent = angular.toJson({
                            id: currentRow[j].dataElement + '-' + currentRow[j].categoryOptionCombo + '-val',
                            value: ''
                        })
                    } else {
                        cell.textContent = currentRow[j].colData;
                    }
                }
            }
            let x = document.createElement("TABLE");
            x.append(tableEl.tBodies[0]);
            this.elements = [{val: tableRows}];
            this.createdTable = x;
        }
    }

    displayData(last) {
        if (last) {
            this.showDataSets();
        }
    }

    editDistance(s1, s2) {
        s1 = s1.toLowerCase();
        s2 = s2.toLowerCase();

        let costs = new Array();
        for (let i = 0; i <= s1.length; i++) {
            let lastValue = i;
            for (let j = 0; j <= s2.length; j++) {
                if (i === 0)
                    costs[j] = j;
                else {
                    if (j > 0) {
                        let newValue = costs[j - 1];
                        if (s1.charAt(i - 1) !== s2.charAt(j - 1))
                            newValue = Math.min(Math.min(newValue, lastValue),
                                    costs[j]) + 1;
                        costs[j - 1] = lastValue;
                        lastValue = newValue;
                    }
                }
            }
            if (i > 0)
                costs[s2.length] = lastValue;
        }
        return costs[s2.length];
    }

    similarity(s1, s2) {
        let longer = s1;
        let shorter = s2;
        if (s1.length < s2.length) {
            longer = s2;
            shorter = s1;
        }
        let longerLength = longer.length;
        if (longerLength == 0) {
            return 1.0;
        }
        return (longerLength - this.editDistance(longer, shorter)) / parseFloat(longerLength);
    }

    processCategories(categories) {

        let boys = [];
        for (let i = 0; i < categories.length; i++) {
            if (i <= 0) {
                boys = [...categories]
            } else {
                let currentOptions = categories[i].categoryOptions;
                let previousOptions = categories[i - 1].categoryOptions;
                let currentLength = currentOptions.length;
                let previousLength = previousOptions.length;
                let current = [];
                for (let j = 0; j < previousLength; j++) {
                    let prev = previousOptions[j];

                    for (let k = 0; k < currentLength; k++) {
                        let id;
                        if (prev.combo) {
                            id = prev.combo + ',' + prev.id;
                        } else {
                            id = prev.id
                        }
                        current = [...current, _.merge({combo: id}, currentOptions[k])];
                    }
                }

                let element = categories[i];

                let newObj = Object.assign({}, element, {categoryOptions: current})

                boys = [...categories.slice(0, i), newObj, ...categories.slice(i + 1)]

            }
            if (boys[i + 1]) {
                let cats = [];
                _.forEach(boys[i].categoryOptions, (opt) => {
                    opt.colSpan = boys[i + 1].categoryOptions.length;
                    cats.push(opt);
                });
                boys[i].categoryOptions = cats;
            }
        }
        return boys;
    }

    download() {
        let dataSetCategoryCombo = "";
        if (this.form.categoryCombo) {
            dataSetCategoryCombo = this.form.categoryCombo.id;
        }

        let defaultCellStyle = {
            font: {name: "Verdana", sz: 11, color: "FF00FF88"},
            fill: {fgColor: {rgb: "FFFFAA00"}}
        };

        let wb = xlsx.utils.book_new();

        let sh = xlsx.utils.table_to_sheet(this.createdTable,{sheet: "Main"});
        sh['!protect'] = {
            password: "password",
            /* enable formatting rows and columns */
            formatRows: false,
            formatColumns: false,
            formatCells: false,
            /* disable editing objects and scenarios */
            objects: true,
            scenarios: true
        };

        let dataEntryCells = [];
        let otherEntryCells = [];
        let organizationCell = '';
        let periodCell = '';
        _.forEach(sh, (cell, key) => {
            let cellValue = cell.v;
            if (!key.startsWith('!')) {
                if (cellValue) {
                    try {
                        let foundObject = angular.fromJson(cellValue);
                        if (foundObject.id.endsWith('-val')) {
                            let otherFields = foundObject.id.split("-");
                            dataEntryCells = [...dataEntryCells, key + '-' + otherFields[0] + '-' + otherFields[1]];
                            cell.v = '';
                            cell.s = this.dataEntryStyling;
                        } else {
                            cell.v = '';
                        }
                    } catch (e) {
                        if (this.searchName(this.selectedDataset.categoryOptions, String(cellValue))) {
                            cell.s = this.headersStyling;
                        }

                        if (this.searchName(this.selectedDataset.dataElements, String(cellValue))) {
                            cell.s = this.dataElementStyling;
                        }
                        let cellParts = String(cellValue).split("-");
                        if (String(cellValue).endsWith('-Organizations')) {
                            organizationCell = key;
                            cell.v = cellParts[1];
                            cell.s = this.dataEntryStyling;
                        }
                        if (String(cellValue).endsWith('-Periods')) {
                            periodCell = key;
                            cell.v = cellParts[1];
                            cell.s = this.dataEntryStyling;
                        }
                        if (String(cellValue).endsWith('-DatasetCategories')) {
                            otherEntryCells = [...otherEntryCells, key];
                            cell.v = cellParts[1];
                            cell.s = this.dataEntryStyling;
                        }
                    }
                }
            }
        });
        /*Sheet Names*/
        let mainSheetName = "Main";
        let propSheetName = "Properties";

        /*Add Data Cells to Custom Properties*/
        let arrays = _.chunk(dataEntryCells, 1);

        arrays[0][1] = this.selectedDataset.id + "," + this.selectedDataset.displayName;
        arrays[0][2] = dataSetCategoryCombo;
        arrays[0][3] = organizationCell;
        arrays[0][4] = periodCell;

        _.forEach(otherEntryCells, (others, i) => {
            arrays[i][5] = others;
        });

        let props = xlsx.utils.aoa_to_sheet([
            ["entryCell", 'dataSet', 'categoryCombination', 'organization', 'period', 'category'],
            ...arrays
        ]);

        xlsx.utils.book_append_sheet(wb, sh, mainSheetName);
        xlsx.utils.book_append_sheet(wb, props, propSheetName);
        xlsx.utils.book_set_sheet_visibility(wb, "Properties", xlsx.utils.consts.SHEET_HIDDEN);

        let wbout = xlsxs.write(wb, {
            bookType: 'xlsx',
            bookSST: true,
            type: 'binary'
        });

        saveAs(new Blob([this.utils.s2ab(wbout)], {type: "application/octet-stream"}), this.selectedDataset.displayName + ".xlsx");
    }

    download1() {
        var data = [
            [1, 2, 3],
            [true, false, null, "sheetjs"],
            ["foo", "bar", new Date("2014-02-19T14:30Z"), "0.3"],
            ["baz", null, "qux", 3.14159],
            ["hidden"],
            ["visible"]
        ];

        var ws_name = "SheetJS";

        var wscols = [
            {wch: 6}, // "characters"
            {wpx: 50}, // "pixels"
            ,
            {hidden: true} // hide column
        ];

        /* At 96 PPI, 1 pt = 1 px */
        var wsrows = [
            {hpt: 12}, // "points"
            {hpx: 16}, // "pixels"
            ,
            {hpx: 24},
            {hidden: true}, // hide row
            {hidden: false}
        ];

        console.log("Sheet Name: " + ws_name);
        console.log("Data: ");
        var i = 0;
        for (i = 0; i !== data.length; ++i) console.log(data[i]);
        console.log("Columns :");
        for (i = 0; i !== wscols.length; ++i) console.log(wscols[i]);

        var wb = xlsx.utils.book_new();

        /* convert an array of arrays in JS to a CSF spreadsheet */
        var ws = xlsx.utils.aoa_to_sheet(data, {cellDates: true});

        /* TEST: add worksheet to workbook */
        /*
         wb.SheetNames.push(ws_name);
         wb.Sheets[ws_name] = ws;
         */
        xlsx.utils.book_append_sheet(wb, ws, ws_name);

        /* TEST: simple formula */
        ws['C1'].f = "A1+B1";
        ws['C2'] = {t: 'n', f: "A1+B1"};

        /* TEST: single-cell array formula */
        /*
         ws['D1'] = {t:'n', f:"SUM(A1:C1*A1:C1)", F:"D1:D1"};
         */
        xlsx.utils.sheet_set_array_formula(ws, 'D1:D1', "SUM(A1:C1*A1:C1)");

        /* TEST: multi-cell array formula */
        /*
         ws['E1'] = {t:'n', f:"TRANSPOSE(A1:D1)", F:"E1:E4"};
         ws['E2'] = {t:'n', F:"E1:E4"};
         ws['E3'] = {t:'n', F:"E1:E4"};
         ws['E4'] = {t:'n', F:"E1:E4"};
         */
        xlsx.utils.sheet_set_array_formula(ws, 'E1:E4', "TRANSPOSE(A1:D1)");
        ws["!ref"] = "A1:E6";

        /* TEST: column props */
        ws['!cols'] = wscols;

        /* TEST: row props */
        ws['!rows'] = wsrows;

        /* TEST: hyperlink note: Excel does not automatically style hyperlinks */
        /*
         ws['A3'].l = { Target: "http://sheetjs.com", Tooltip: "Visit us <SheetJS.com!>" };
         */
        xlsx.utils.cell_set_hyperlink(ws['A3'], "http://sheetjs.com", "Visit us <SheetJS.com!>");

        /* TEST: built-in format */
        /*
         ws['B1'].z = "0%"; // Format Code 9
         */
        xlsx.utils.cell_set_number_format(ws['B1'], "0%");

        /* TEST: custom format */
        var custfmt = "\"This is \"\\ 0.0";
        /*
         ws['C2'].z = custfmt;
         */
        xlsx.utils.cell_set_number_format(ws['C2'], custfmt);

        /* TEST: page margins */
        ws['!margins'] = {left: 1.0, right: 1.0, top: 1.0, bottom: 1.0, header: 0.5, footer: 0.5};

        console.log("JSON Data:");
        console.log(xlsx.utils.sheet_to_json(ws, {header: 1}));

        /* TEST: hidden sheets */
        /*
         wb.SheetNames.push("Hidden");
         wb.Sheets["Hidden"] = xlsx.utils.aoa_to_sheet(["Hidden".split(""), [1,2,3]]);
         wb.Workbook = {Sheets:[]};
         wb.Workbook.Sheets[1] = {Hidden:1};
         */
        var data_2 = ["Hidden".split(""), [1, 2, 3]];
        xlsx.utils.book_append_sheet(wb, xlsx.utils.aoa_to_sheet(data_2), "Hidden");
        xlsx.utils.book_set_sheet_visibility(wb, "Hidden", xlsx.utils.consts.SHEET_HIDDEN);

        /* TEST: properties */
        wb.Props = {
            Title: "SheetJS Test",
            Subject: "Tests",
            Author: "Devs at SheetJS",
            Manager: "Sheet Manager",
            Company: "SheetJS",
            Category: "Experimentation",
            Keywords: "Test",
            Comments: "Nothing to say here",
            LastAuthor: "Not SheetJS",
            CreatedDate: new Date(2017, 1, 19)
        };

        /* TEST: comments */
        /*
         ws['A4'].c = [];
         ws['A4'].c.push({a:"SheetJS",t:"I'm a little comment, short and stout!\n\nWell, Stout may be the wrong word"});
         */
        xlsx.utils.cell_add_comment(ws['A4'], "I'm a little comment, short and stout!\n\nWell, Stout may be the wrong word", "SheetJS");

        /* TEST: sheet protection */
        ws['!protect'] = {
            password: "password",
            /* enable formatting rows and columns */
            formatRows: false,
            formatColumns: false,
            /* disable editing objects and scenarios */
            objects: true,
            scenarios: true
        };

        /* TEST: Workbook Properties */
        if (!wb.Workbook) wb.Workbook = {Sheets: [], WBProps: {}};
        if (!wb.Workbook.WBProps) wb.Workbook.WBProps = {};
        wb.Workbook.WBProps.filterPrivacy = true;

        console.log("Worksheet Model:");
        console.log(ws);

        var filenames = [
            ['sheetjs.xlsx', {bookSST: true}],
            ['sheetjs.xlsm'],
            ['sheetjs.xlsb'],
            ['sheetjs.xls', {bookType: 'biff2'}],
            ['sheetjs.xml.xls', {bookType: 'xlml'}],
            ['sheetjs.ods'],
            ['sheetjs.fods'],
            ['sheetjs.slk'],
            ['sheetjs.csv'],
            ['sheetjs.txt'],
            ['sheetjs.prn'],
            ['sheetjs.dif']
        ];

        let wbout = xlsx.write(wb, {
            bookType: 'xlsx',
            bookSST: true,
            type: 'binary'
        });
        saveAs(new Blob([this.utils.s2ab(wbout)], {type: "application/octet-stream"}), this.selectedDataset.displayName + ".xlsx")

        // filenames.forEach(function (r) {
        //     let wbout = xlsx.write(wb, {
        //         bookType: 'xlsx',
        //         bookSST: true,
        //         type: 'binary'
        //     });
        //     saveAs(new Blob([this.utils.s2ab(wbout)], {type: "application/octet-stream"}), this.selectedDataset.displayName + ".xlsx");
        //
        //     /* /!* write file *!/
        //      xlsx.writeFile(wb, r[0], r[1]);
        //      /!* test by reading back files *!/
        //      xlsx.readFile(r[0]);*/
        // });
    }

    onSubmit() {

        let date = new Date();
        let day = date.getDate();
        let monthIndex = date.getMonth();
        let year = date.getFullYear();

        let per = year + '-' + (monthIndex + 1) <= 9 ? '0' + (monthIndex + 1) : (monthIndex + 1) + '-' + day <= 9 ? '0' + day : day;
        let data = [];


        _.forEach(this.cellsGot, (cell) => {
            if (cell.categoryOptionCombo && cell.dataElement && cell.cellValue) {
                let ele = {
                    dataElement: cell.dataElement,
                    categoryOptionCombo: cell.categoryOptionCombo,
                    value: cell.cellValue.v
                };
                data = [...data, ele];
            }
        });

        if (data.length > 0) {
            let catOptions = this.selectedDatasetCategories.split(',');
            if (this.selectedDatasetCategories !== "") {
                this.data.getMany('categoryCombos', {
                    filter: 'id:in:[' + this.dataSetCategoryCombo + ']',
                    fields: 'categoryOptionCombos[id,categoryOptions[id,name]]'
                }).then((categoryCombos) => {
                    let dataCombos = this.restagular.stripRestangular(categoryCombos);
                    let categoryOptionCombos = _.flatten(_.map(dataCombos, 'categoryOptionCombos'));
                    for (let i = 0; i < categoryOptionCombos.length; i++) {
                        let opt = _.map(categoryOptionCombos[i]['categoryOptions'], function (o) {
                            return o.id
                        });
                        if (_.every(opt, (val) => {
                                return catOptions.indexOf(val) >= 0;
                            }) && catOptions.length == opt.length) {
                            let processedData = {
                                dataSet: this.importedDatasetId,
                                completeDate: this.utils.dateToYMD(new Date()),
                                period: this.importedPeriod,
                                orgUnit: this.importedOrganisationUnit,
                                attributeOptionCombo: categoryOptionCombos[i].id
                            };
                            processedData.dataValues = data;
                            this.data.post('dataValueSets', angular.toJson(processedData)).then((insertedRecords) => {
                                this.open(this.restagular.stripRestangular(insertedRecords));
                            });
                            break;
                        }
                    }
                });
            } else {
                let processedData = {
                    dataSet: this.importedDatasetId,
                    completeDate: this.utils.dateToYMD(new Date()),
                    period: this.importedPeriod,
                    orgUnit: this.importedOrganisationUnit
                };
                processedData.dataValues = data;
                this.data.post('dataValueSets', angular.toJson(processedData)).then((insertedRecords) => {
                    this.open(this.restagular.stripRestangular(insertedRecords));
                });
            }
        } else {
            this.message = 'No Data';
        }
        this.excel = null;
        this.cellsGot = [];
    };
}

HomeController.$inject = ['$scope', 'Data', 'Restangular', 'Filters', 'Utils', '$uibModal', 'dataSets'];
