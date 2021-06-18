"use strict";
(function ($) {
    // angular stuff
	app
        // routing
        .config([
            "$routeProvider", "$mdThemingProvider", "$mdDateLocaleProvider", function ($routeProvider, $mdThemingProvider, $mdDateLocaleProvider) {
				// USS theming
                $mdThemingProvider.theme('default').primaryPalette('red', {
                    'default': '800',
                    'hue-1': '900',
                    'hue-2': '900'
                }).accentPalette('grey');
				
				// UK dates
				$mdDateLocaleProvider.formatDate = function(date) {
					return date == null ? null : (('0' + date.getDate()).slice(-2) + '/' + ('0' + (date.getMonth() + 1)).slice(-2) + '/' + date.getFullYear());
                };

                /* Any common app functions or variables here */
                
                // Routing for pages
                $routeProvider
                    .when('/Raise', {
                        redirectTo: '/<something>/0/0/null'
                    })
                    /* Can have multiple controllers, common to use a prefix then trap stage and editing id (or 0/0 if new item under same route) */
					.when('/<something>/:Stage/:Id/:Tab', {
                        /* Template view file */
                        templateUrl: window.appsubfolder + "Views/Department.html",
                        /* Controller of Forms or Home */
						controller: 'Forms',
                        resolve: {
                            listName: function () {
                                /* List Title for saving to */
                                return 'Departments';
                            },
                            spData: function () {
                                /* Override SP.Data.<this>ListItem where required, i.e. list title changed since creation */
                                return null;
                            },
                            additFuncs: function () {
                                return function (scope, webSvc, listSvc, $uibModal, $timeout, $q) {
                                    /* Any additional functions related to this route including new, i.e. Loading additional lists */
                                    scope.openViewModalExample = function (id) {
                                        $uibModal.open({
                                            animation: true,
                                            templateUrl: window.appsubfolder + "Views/Department.html",
						                    controller: 'Forms',
                                            resolve: {
                                                listName: function () {
                                                    return 'Departments';
                                                },
                                                spData: function () {
                                                    return null;
                                                },
                                                additFuncs: function () {
                                                    return null;
                                                },
                                                additFields: function () {
                                                    return null;
                                                },
                                                additSave: function () {
                                                    return null
                                                },
                                                $routeParams: function () {
                                                    return {
                                                        Id: id,
                                                        Stage: '-1'
                                                    }
                                                }
                                            }
                                        });
                                    }
                                    if (!scope.id)
                                        return;
                                    /* Any additional functions related to this route for loaded data ids */
                                };
                            },
                            additFields: function () {
                                /* Force specific field values on new and on save */
                                return [{field: 'Title', value: (new Date()).toJSON()}];;
                            },
                            additSave: function () {
                                return function (form, allowInvalid, stageViewed, scope) {
                                    /* During save function, additional logic such as changing status if conditions met, to invalidate form from here update scope.form.<field> */
                                }
                            }
                        }
                    })
					// home
                    .otherwise({
                        templateUrl: window.appsubfolder + "Views/Home.html",
                        controller: "Home",
						resolve: {
                            listName: function () {
                                /* List Title for primary loading from */
                                return 'Departments';
                                /* Can be a function call for variations */
                                return function (scope) {
                                    return '';
                                }
                            },
                            beforeLoad: function () {
                                return function (scope, webSvc, listSvc, $uibModal, $timeout) {
                                    /* Any additional functions related to this route */
                                };
                            },
                            filter: function () {
                                /* OData style filter */
                                return "Title ne ''";
                                /* Can be a function call for variations */
                                return function (scope) {
                                    return '';
                                }
                            },
                            select: function () {
                                /* OData style select, will automatically calculate expands */
                                return 'Id,Title,Author/Id,Author/Title,ExCo/Id,ExCo/Title,Owner/Id,Owner/Title,Representatives/Id,Representatives/Title,Site,Status,History';
                                /* Can be a function call for variations */
                                return function (scope) {
                                    return '';
                                }
                            },
                            order: function () {
                                /* OData style order */
                                return 'Id';
                                /* Can be a post data load as a JS order */
                                return function(a, b){
									return a.Id-b.Id;
								};
                            },
                            rows: function () {
                                /* Functions to process the loaded row data before it gets applied */
                                return function (r) {
                                    r.title = (r.Title || '').toLowerCase();
                                }
                            },
                            afterLoad: function () {
                                /* After the data has loaded some additional functions */
                                return function afterLoad(d, scope) {
                                }
                            },
                            search: function () {
                                /* On page search fields and filters */
                                return function (scope, r) {
                                    return (
                                        scope.Search == null
                                        || scope.Search == ''
                                        || ~r.title.indexOf(scope.Search.toLowerCase())
                                    )
                                }
                            },
                            refresh: function () {
                                /* Auto refresh the dashboard every 15seconds, or specify own time in ms */
                                return false;
                            },
                            saveFilter: function () {
                                /* Should the status filters be saved on changes */
                                return false;
                            }
                        }
                    })
            }
        ]);
})(USS$ || $);
