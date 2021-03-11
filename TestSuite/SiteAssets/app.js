"use strict";
(function ($) {
    app 
        // theming, formatting, routing and app logic
        .config([
            "$routeProvider", "$mdThemingProvider", "$mdDateLocaleProvider", function ($routeProvider, $mdThemingProvider, $mdDateLocaleProvider) {
				// theming
                $mdThemingProvider.theme('default').primaryPalette('red', {
                    'default': '900',
                    'hue-1': '800',
                    'hue-2': '800'
                }).accentPalette('grey');

				// UK dates
				$mdDateLocaleProvider.formatDate = function(date) {
					return date == null ? null : (('0' + date.getDate()).slice(-2) + '/' + ('0' + (date.getMonth() + 1)).slice(-2) + '/' + date.getFullYear());
				};

                // Routing for home page
                // To support SP modals all routes end in a / so have an /:ignore here as SP modals will append IsDlg=1
                // Can use otherwise for home page so the url doesnt get appended #/ or #!/
                $routeProvider
					.when('/Test/:Id', {
                        controller: "Forms",
                        templateUrl: window.appsubfolder + "Views/Test.html",
                        resolve: {
                            listName: function () {
                                return 'null';
                            },
                            spData: function () {
                                return null;
                            },
                            additFuncs: function () {
                                return function (scope, webSvc, listSvc) {
                                    scope.people = function () {
                                        return JSON.stringify(scope.People);
                                    }
                                    scope.People = [];
                                    scope.cal = {
                                        plugins: [ 'dayGrid' ],
                                        height: 450,
                                        editable: true,
                                        eventLimit: true,
                                        defaultDate: moment(),
                                        timeFormat: 'h(:mm)a',
                                        header:{
                                          left: 'title',
                                          center: '',
                                          right: 'today prev,next'
                                        },
                                        dayClick: function (e) {
                                        },
                                        eventClick: function (e) {
                                        }
                                    };
                                    webSvc.user().then(function (u) {
                                        var dept = u.d.UserProfileProperties.results.filter(function (x) {return x.Key == "Department";});
                                        scope.form.Department = dept.length > 0 ? dept[0].Value : '';
                                        scope.$apply();
                                    });
                                    scope.f = [[]]
                                };
                            },
                            additFields: function () {
                                return [
                                    {field: 'KId', value: _spPageContextInfo.userId},
                                    {field: 'LId', value: {results:[_spPageContextInfo.userId]}},
                                    {field: 'M', value: 12345.6768},
                                    {field: 'Title', value: _spPageContextInfo.userDisplayName},
                                    {field: 'People', value: _spPageContextInfo.userId},
                                    {field: 'Start', value:'<h1>Tested</h1>'}
                                ];
                            },
                            additSave: function () {
                                return function (f, allowInvalid, viewedStage, scope) {
                                    delete f.D;
                                    delete f.MMS;
                                    delete f.KId;
                                    delete f.LId;
                                    delete f.M;
                                    return;
                                }
                            }
                        }
					})
                    .otherwise({
                        redirectTo: '/Test/1'
                    });
            }
        ]);
})(APP$ || $);