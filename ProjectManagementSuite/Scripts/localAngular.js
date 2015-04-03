// declare var containing angular module
var projApp = angular.module('ProjectApp', ['datatables']);
//
// define angular service to broadcast project id - inject service into project header controller & project line controller
//
projApp.factory('projectIdentityService', function ($rootScope) {
    var identityService = {};

    identityService.identity = '';

    identityService.prepForBroadcast = function (projid) {
        this.identity = projid;
        this.broadcastItem();
    };

    identityService.broadcastItem = function () {
        $rootScope.$broadcast('handleBroadcast');
    };

    return identityService;
});
//
// define datatables controller
//
projApp.controller('withAjaxCtrl', function ($scope, $http, $compile, DTOptionsBuilder, DTColumnBuilder, projectIdentityService) {
    //---------------------------------------------------------------------------------------------------------------------
    // pull up project details for a particular item - replace periods in item with underscores otherwise 404s will result
    //---------------------------------------------------------------------------------------------------------------------
    //
    $scope.getProjectData = function () {
        $http({ method: 'GET', url: 'api/itembystate/' + $scope.itemToGet.replace(/\./g,'_') }).
        success(function (data) {
            var jdat = JSON.parse(data);
            $scope.statedata = [];
            angular.forEach(jdat, function (value, key) {
                $scope.statedata.push(value);
            });
        }).
        error(function (data, status, headers, config) {
        });
    }
    //-----------------------------------------------------------------------------------------------------------------
    // clear old project by state for item table - clearProjectForItem
    //-----------------------------------------------------------------------------------------------------------------
    $scope.clearProjectForItem = function () {
        // clean all the table data off the modal
        $("#table-for-item-by-projects").children().remove();
        $scope.itemToGet = '';
        $scope.statedata = [];
    }
    //-----------------------------------------------------------------------------------------------------------------
    // pull up project header details in order to edit
    //-----------------------------------------------------------------------------------------------------------------
    //
    $scope.edit = function (ProjID) {
        $http({ method: 'GET', url: 'api/projects/' + ProjID }).
            success(function (data, status, headers, config) {
                $scope.editrecord = data[0];        // a list with a single element has been returned therefore reference data[0]
            }).
            error(function (data, status, headers, config) {
                // callback with failure details
            });
        // Then reload the data so that DT is refreshed
        $scope.dtOptions.reloadData();
    };
    //-----------------------------------------------------------------------------------------------------------------
    // send edited project header details back to server to update DB 
    //-----------------------------------------------------------------------------------------------------------------
    //
    $scope.updedit = function (record) {
        $http({ method: 'PUT', url: 'api/projects/' + record.ProjID, data: JSON.stringify(record) }).
            success(function (data, status, headers, config) {
                $scope.dtOptions.reloadData();
            }).
            error(function (data, status, headers, config) {
            });
    };
    //-----------------------------------------------------------------------------------------------------------------
    // delete out ProjID header and lines 
    //-----------------------------------------------------------------------------------------------------------------
    //
    $scope.delete = function (ProjID) {
        // are you sure ..... ?????
        alertify.confirm("Are you sure you want \n to delete ProjID: " + ProjID + " ?", function (e) {
            if (e) {
                // Delete some data and call server to make changes...
                $http({ method: 'DELETE', url: 'api/projects/' + ProjID }).
                     success(function (data, status, headers, config) {
                         // reload server data so that DT is refreshed
                         $scope.dtOptions.reloadData();
                     }).
                    error(function (data, status, headers, config) {
                        // callback with failure details
                    });
            } else {
                // user clicked "cancel"
            }
        });
    };
    //------------------------------------------------------------------------------------------------------------------
    // edit individual lines for ProjID 
    // broadcast project identity to project line controller - 'withProjectLines'
    //------------------------------------------------------------------------------------------------------------------
    //
    $scope.lines = function (ProjID) {
        $scope.message = 'Downloading Excel File of Project : ' + ProjID;
        // load project identity into broadcast service then broadcast ProjID
        projectIdentityService.prepForBroadcast(ProjID);
    };
    //------------------------------------------------------------------------------------------------------------------
    // upload spreadsheets to add new project data to DB - build http post function with Ajax - not Angularjs
    //------------------------------------------------------------------------------------------------------------------
    //
    $scope.uploadFiles = function () {
        // clear upload files placeholder div in modal
        $('#upload-files-placeholder').children().remove();
        // add back a new file upload template
        $('#upload-files-placeholder').append(Mustache.render($('#fileUploadTemplate').html()));
        // attach new upload file form to uplFileForm
        var upload = document.getElementById('selectedFiles');
        // set up eventListener function
        upload.addEventListener("change", function (event) {
            var files = upload.files,
                len = files.length;
            // for http : POST operation
            var uplform = new FormData();
            // load files from SELECTION process
            for (i = 0 ; i < len; i++) {
                var jdat = { 'filename': files[i].name };
                var addFiles = $("#files-placeholder").append(Mustache.render($('#fileUploadItemTemplate').html(), jdat));
                uplform.append("file" + i, files[i]);
            }
            //-----------------------------------------------------------------------------------------
            // send to api/fileupload - async GET web api task
            // WILL GENERATE INTERNAL SERVER ERROR IF FOLDER 'UPLOADS' DOES NOT EXIST ON WEBSITE
            //-----------------------------------------------------------------------------------------
            $http({
                method: 'POST',
                url: 'api/fileupload',
                data: uplform,
                headers: { 'Content-Type': undefined },
                transformRequest: function (data) { return data; }      // stop angularjs trying to convert data to JSON
            }).success(function (data, status, headers, config) {
                // advise of success or otherwise by alert for the moment
                for (i = 0; i < data.length; i++) {
                    alertify.alert(data[i]);
                }
                // refresh 
                $scope.dtOptions.reloadData();
                $('#shutDownUpload').trigger('click');                  // fire modal 'close' to ensure modal goes away !!!
                //
            }).error(function (data, status, headers, config) {
                // callback with failure details
            });
        },false);

    };
    //------------------------------------------------------------------------------------------------------------------
    // TEMPLATE GENERATOR FUNCTIONS
    //------------------------------------------------------------------------------------------------------------------
    //
    // sample model data - to serve as an example for user
    $scope.proj = [{ "projectNumber": "1450684", "projectName": "QUEST APARTMENTS BERRIMAH APR 14", "projectState": "SA", "projectType": "MEDIUM DENSITY", "mvxorders": [] }, { "projectNumber": "1322420", "projectName": "ROYAL NTH SHORE HOSP-CLINICAL SERV STG 3", "projectState": "NSW", "projectType": "Hospital", "mvxorders": [{ "order": "1002809691" }, { "order": "1002915480" }, { "order": "1002966523" }] }, { "projectNumber": "1380365", "projectName": "RENDEZVOUS HOTEL", "projectState": "NZ", "projectType": "Hotel", "mvxorders": [{ "order": "1002805423" }, { "order": "1002805424" }] }];

    $scope.resetTemplGen = function () {
        $scope.proj = [{ "projectNumber": "1450684", "projectName": "QUEST APARTMENTS BERRIMAH APR 14", "projectState": "SA", "projectType": "MEDIUM DENSITY", "mvxorders": [] }, { "projectNumber": "1322420", "projectName": "ROYAL NTH SHORE HOSP-CLINICAL SERV STG 3", "projectState": "NSW", "projectType": "Hospital", "mvxorders": [{ "order": "1002809691" }, { "order": "1002915480" }, { "order": "1002966523" }] }, { "projectNumber": "1380365", "projectName": "RENDEZVOUS HOTEL", "projectState": "NZ", "projectType": "Hotel", "mvxorders": [{ "order": "1002805423" }, { "order": "1002805424" }] }];
    };
    //
    // add in blank project
    $scope.addProject = function () {
        $scope.proj.push({
            projectNumber: '',
            projectName: '',
            projectState: '',
            projectType: '',
            mvxorders: []
        });
    };
    // remove movex order from project
    $scope.removeMovexOrder = function (parentidx, childidx) {
        $scope.proj[parentidx].mvxorders.splice(childidx, 1);
    };
    // add movex order to project
    $scope.addMovexOrder = function (parentidx, neword) {
        $scope.proj[parentidx].mvxorders.push({ order: [''] });
    };
    // remove project completely
    $scope.removeProject = function (parentidx) {
        $scope.proj.splice(parentidx, 1);
    };
    // POST new project header details to server
    $scope.sendJSONtoServer = function () {
        $http({ method: 'POST', url: 'api/template', data: JSON.stringify($scope.proj) }).
        success(function (data, status, headers, config) {
            //
            var div = document.getElementById('response');
            // use HTML-5 download file function of anchor link
            for (var key in data) {
                var newlink = document.createElement('a');
                newlink.setAttribute('href', data[key]);
                newlink.setAttribute('download', key);
                div.appendChild(newlink).click();
            }
            // finally close down the template modal
            $('#shutDownTempl').trigger('click');
            $scope.resetTemplGen();
        });
    };
    //------------------------------------------------------------------------------------------------------------------
    // generate issues log of projects with problem items
    //------------------------------------------------------------------------------------------------------------------
    $scope.generateProblemLog = function () {
        $http({ method: 'GET', url: 'api/itembystate' }).
        success(function (data, status, headers, config) {
            //
            var div = document.getElementById('problemitemlog');
            // use HTML-5 download file function of anchor link
            for (var key in data) {
                var newlink = document.createElement('a');
                newlink.setAttribute('href', data[key]);
                newlink.setAttribute('download', key);
                div.appendChild(newlink).click();
            }
        });
    };

    //------------------------------------------------------------------------------------------------------------------
    // Datatable options
    //------------------------------------------------------------------------------------------------------------------
    //
    $scope.dtOptions = DTOptionsBuilder.newOptions()
        .withOption('ajax', {
            url: 'api/projects',
            type: 'GET'
        })
        .withPaginationType('simple_numbers')
        .withOption('aLengthMenu', [10, 15, 20, 25])   // set array of page-lengths
        .withOption('iDisplayLength', 10)              // set default starting length
        // pick up row that was selected
        .withOption('createdRow', function (row, data, dataIndex) {
            // Recompiling so we can bind Angular directive to the DT
            $compile(angular.element(row).contents())($scope);
        })
            // Add Table tools compatibility
        .withTableTools('Content/DataTables-1.10.3/swf/copy_csv_xls_pdf.swf')
        .withTableToolsButtons([
            'copy',
             {
                 'sExtends': 'collection',
                 'sButtonText': 'Save',
                 'aButtons': ['csv', 'pdf']
             }]);


    $scope.dtColumns = [
        DTColumnBuilder.newColumn('ProjID').withTitle('ProjectID').withOption('width',20),
        DTColumnBuilder.newColumn('ProjNum').withTitle('Project Number'),
        DTColumnBuilder.newColumn('ProjName').withTitle('Project Name'),
        DTColumnBuilder.newColumn('ProjDesc').withTitle('Project Desc'),
        DTColumnBuilder.newColumn('State').withTitle('Project State').withOption('width',20),
        DTColumnBuilder.newColumn('ProjManFlag').withTitle('Manual Forecast').withOption('width',20),
        DTColumnBuilder.newColumn('MVXProjNum2').withTitle('MOVEX Orders'),
        DTColumnBuilder.newColumn(null).withTitle('Actions').notSortable()
                       .renderWith(function (data, type, full, meta) {
                           return '<button class="btn btn-warning" ng-click="edit(' + data.ProjID + ')" data-toggle="modal" data-target="#EditProjectModal">' +
                           '   <i class="fa fa-edit"></i>' +
                           '</button>&nbsp;' +
                           '<button class="btn btn-success" ng-click="lines(' + data.ProjID + ')">' +
                           '   <i class="fa fa-list"></i>' +
                           '</button>&nbsp;' +
                           '<button class="btn btn-danger" ng-click="delete(' + data.ProjID + ')">' +
                           '   <i class="fa fa-trash-o"></i>' +
                           '</button>';
                       })
    ];
});

// inject projectIdentityService into project lines controller
//
projApp.controller('withProjectLines', function ($scope,  projectIdentityService, $http) {
    $scope.$on('handleBroadcast', function () {
        $scope.bbc = "Downloading Project Number : " +  projectIdentityService.identity;
        $http({ method: 'GET', url: 'api/projectlines/' + projectIdentityService.identity }).
    success(function (data, status, headers, config) {
        //
        var div = document.getElementById('projectdownload');
        // use HTML-5 download file function of anchor link
        for (var key in data) {
            var newlink = document.createElement('a');
            newlink.setAttribute('href', data[key]);
            newlink.setAttribute('download', key);
            div.appendChild(newlink).click();
        }
        $scope.bbc = "";
    }).
    error(function (data, status, headers, config) {
        // callback with failure details
    });
    });
});


// activate wait when ajax fires - deactivate when ajax has finished
$(document).ajaxStart(function () {
    $('#loading').show();
}).ajaxStop(function () {
    $('#loading').hide();
});



