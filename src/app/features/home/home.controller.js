import * as xlsx from "xlsx";
export default class HomeController {
    constructor($scope, Data, Restangular, Utils, $uibModal) {

        this.items = [];
        this.utils = Utils;
        this.uimodal = $uibModal;
        this.data = Data;
        this.restagular = Restangular;
        this.cellsGot = [];
        this.selectedDatasetCategories = null;
        this.excel = null;

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
    displayData(last) {
        if (last) {
            this.showDataSets();
        }
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
                    console.log(insertedRecords.conflicts);
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

HomeController.$inject = ['$scope', 'Data', 'Restangular', 'Utils', '$uibModal'];
