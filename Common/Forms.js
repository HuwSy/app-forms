"use strict";

(function ($) {
	var appMod = angular.module('Forms', ['appforms']);
    
    appMod.controller("Forms", function ($scope, listSvc, webSvc, fileSvc, $routeParams, $uibModal, $timeout, $location, $q, $rootScope,
		listName, spData, additFuncs, additFields, additSave) {
		// stop reloading home.js loop
		if (typeof app.interval != "undefined")
			clearInterval(app.interval);
		$scope.go = function (g) {
			try {
				$rootScope.$apply(function() {
					$location.path(g);
				});
			} catch (e) {
				$location.path(g);
			}
		}
		// permissions
		$scope.permissions = Override.Permissions;
		$scope.userId = _spPageContextInfo.userId;
		$scope.web = _spPageContextInfo.webServerRelativeUrl;
		// start of day
		$scope.today = new Date();
		$scope.today.setHours(0,0,0,0);
		// current editing id
		$scope.id = $routeParams && $routeParams.Id > 0 ? $routeParams.Id : null;
		if ($scope.id)
			$scope.today = null;
		// get route in case there are other parts wanted for html
		$scope.routeParams = $routeParams || {};
		if (!$scope.routeParams.Stage)
			$scope.routeParams.Stage = 0;
		// form submitted for validation css
		$scope.submitted = false;
		// field details from list
		$scope.Choices = {};
		$scope.Descriptions = {};
		$scope.Requireds = {};
		$scope.Titles = {};
		$scope.TypeAsString = {};

		// loop through json object ensuring dates are dates
		$scope.parseLoop = function (i) {
			try {
				if (typeof i == "string") {
					if (i.match(/20[0-9]{2}\-[01][0-9]\-[0-3][0-9]/) != null) {
						return new Date(i);
					}
				} else if (typeof i == "object") {
					try {
						for (var a in i)
							i[a] = $scope.parseLoop(i[a]);
					} catch (e) {}
				} 
			} catch (e) {}
			return i;
		}

		// parse list object ensuring usable data types for angular
		$scope.parseListData = function(d) {
			for (var x in d) {
				if (!d[x])
					continue;
				
				// remove deferred
				if (d[x].__deferred) {
					delete d[x];
					continue;
				}

				// dont change metadata
				if (x == "__metadata")
					continue;

				// parse out json from field where possible, do not store json into html/rich text fields
				try {
					if (d[x].toString().trim().substring(0,1) == '{' || d[x].toString().trim().substring(0,1) == '[') {
						d[x] = JSON.parse(d[x]);
						d[x] = $scope.parseLoop(d[x]);
						continue;
					}
				} catch (e) {}

				// convert date text to date object
				if (d[x].toString().match(/[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]/) != null) {
					d[x] = new Date(d[x]);
					continue;
				}
			}

			return d;
		}

		// show history modal
		$scope.History = function (h, t, l) {
			h = $scope.parseListData(h);
			$uibModal.open({
				animation: true,
				size: $scope.modalSize || 'sm',
				templateUrl: window.appsubfolder + "Views/" + t + ".html",
				controller: 'Forms',
				resolve: {
					listName: function () {
						return l || listName;
					},
					spData: function () {
						return null;
					},
					additFuncs: function () {
						return additFuncs;
					},
					additFields: function () {
						return null;
					},
					additSave: function () {
						return null;
					},
					$routeParams: function () {
						return {
							Id: null,
							Stage: -2,
							Form: h
						}
					}
				}
			})
		}

		// load list data, history and attachments
		$scope.load = function () {
			listSvc.getItems(listName, 'Id eq ' + $scope.id).then(function (d) {
				if (d.d.results.length > 0) {
					// convert sp data to required formats
					$scope.form = $scope.parseListData(d.d.results[0]);
					$scope.guid = $scope.form.__metadata.uri.match(/guid'[^']*/)[0].substring(5);

					// save loaded set
					$scope.form2 = JSON.parse(JSON.stringify($scope.form));

					// overwrite specifics from launching module
					if (additFields)
						for (var x in additFields) {
							$scope.form[additFields[x].field] = additFields[x].value;
						}
					
					$scope.$apply();

					listSvc.history(listName, $scope.id).then(function (d) {
						var h = [];
						for (var i = 0; i < d.d.results.length; i++) {
							// undo double encode
							for (var g in d.d.results[i])
								if (~g.indexOf('_x005f_'))
									d.d.results[i][g.replace(/_x005f_/g,'_')] = d.d.results[i][g];
							// remove reminders
							if (d.d.results[i].History == "Reminder issued" || d.d.results[i].History == "Reminders Issued")
								continue;
							// remove same minute same author, i.e. attachments
							if (i > 0
								&& d.d.results[i].Editor.LookupValue == d.d.results[i-1].Editor.LookupValue
								&& d.d.results[i].Modified.substring(0,16) == d.d.results[i-1].Modified.substring(0,16))
								continue;
							h.push(d.d.results[i]);
						}
						for (var i = 0; i < h.length - 1; i++) {
							// remove duplicate history so updates without history changes don't mislead
							if (h[i+1].History == h[i].History)
								h[i].History = '';
						}
						$scope.history = h;
						$scope.$apply();
					});

					listSvc.getAttachments(listName, $scope.id).then(function (d) {
						$scope.Files = d.d.results;
						$scope.$apply();
					});
				}
			});
		}

		$scope.fields = function (c, r) {
			var n = (r || '') + (c.InternalName.indexOf('_') == 0 ? 'OData_' : '') + c.InternalName;
			if (c.Choices && c.Choices.results && !$scope.Choices[n])
				$scope.Choices[n] = c.Choices.results;
			if (c.Description && !$scope.Descriptions[n])
				$scope.Descriptions[n] = c.Description;
			if (c.Required != null && !$scope.Requireds[n])
				$scope.Requireds[n] = c.Required;
			if (c.Title && !$scope.Titles[n])
				$scope.Titles[n] = c.Title;
			if (c.TypeAsString == "DateTime" && c.DisplayFormat == 0 && !$scope.TypeAsString[n])
				$scope.TypeAsString[n] = "Date";
			if (c.TypeAsString == "Note" && !c.RichText && !$scope.TypeAsString[n])
				$scope.TypeAsString[n] = "Multiple lines of text";
			if (c.TypeAsString != null && !$scope.TypeAsString[n])
				$scope.TypeAsString[n] = c.TypeAsString;
		}
		
		// load list fields
		if (listName)
			listSvc.field(listName, null, null, 'InternalName,Choices,Description,Required,Title,TypeAsString,DisplayFormat,RichText', "InternalName")
				.then(function (d) {
					// choice field options from list
					d.results.forEach(function (d) {
						$scope.fields(d);
					});
					
					// needs data loading
					if ($scope.routeParams.Form) {
						$scope.form = $scope.routeParams.Form;
						// ensure its broke and cant save
						delete $scope.form.Id;
						delete $scope.form.__metadata;
						$scope.$apply();
					} else if ($scope.id) {
						$scope.load();
					} else {
						// overwrite specifics from launching module
						$timeout(function () {
							if (additFields)
								for (var x in additFields) {
									$scope.form[additFields[x].field] = additFields[x].value;
								}
						}, 250);
					}
				});

		// define form to save
		$scope.form = {
			__metadata: {
				type: 'SP.Data.' + (spData || listName || '').replace(/ /g, '_x0020_') + 'ListItem'
			}
		};

		// textarea options
		$scope.tinymceOptions = {
			resize: false,
			selector: "textarea",
			height: 200,
			menubar: false,
			plugins: "textcolor lists table link paste",
			toolbar: "forecolor | bold italic underline | bullist numlist outdent indent | table | link",
			statusbar: false,
			debounce: false,
			paste_data_images: true
		};

		// close method for saved or close
		$scope.close = function(){
			if (typeof $scope.$close == "function") {
				if ($scope.$$prevSibling && typeof $scope.$$prevSibling.load == "function")
					$scope.$$prevSibling.load();
				return $scope.$close();
			}
			parent.postMessage('OK', '*');
			$scope.go('/Home');
		}

		// save methods
		$scope.draft = function (dontClose) {
			if ($scope.form.Status != null && $scope.form.Status.indexOf('raft') < 0)
				$scope.form.Status = 'Draft';
			return $scope.save(true, dontClose);
		}
		$scope.reject = function (allowInvalid, dontClose, field, initial) {
			var scope = $scope;
			scope.submitted = true;
			if (field && scope.form[field] && scope.form[field].trim() != '') {
				return scope.save(allowInvalid, dontClose);
			}
			$uibModal.open({
				size: 'sm',
				template: ' <div class="modal-header">\
								<h4 class="modal-title">Reason</h4>\
							</div>\
							<div class="modal-body">\
								<textarea ng-model="Reason" style="height: 50px; width: 100%; box-sizing: border-box;" placeholder="Reason to email back to the initiator." maxlength="255"></textarea>\
							</div>\
							<div class="modal-footer">\
								<button class="btn" ng-click="$close()" ng-disabled="submitting">Cancel</button>\
								<button class="btn btn-primary" ng-click="ok()" ng-disabled="submitting">OK</button>\
							</div>',
				controller: function ($scope) {
					$scope.Reason = initial || '';
					$scope.ok = function () {
						scope.form[field || 'Rejection'] = $scope.Reason;
						$scope.$close();
						return scope.save(allowInvalid, dontClose);
					}
				}
			});
		}
		$scope.save = function (allowInvalid, dontClose, overrideForm, overrideList, whatIf) {
			var deferred = $q.defer();
			if (!$scope.form.valid && !allowInvalid)
				return $scope.submitted = true;
			
			// get and clean object not reference
			var f = JSON.parse(JSON.stringify(overrideForm || $scope.form))

			// delete specifics from save
			delete f["OData__UIVersionString"];
			delete f["Created"];
			delete f["Modified"];
			delete f["AuthorId"];
			delete f["EditorId"];
			delete f["valid"];
			delete f["$$hashKey"];
					
			// clean up people fields now bring in object, id and string
			var found = 0, r;
			for (var t in f) {
				try {
					if (~t.indexOf("StringId") && t.indexOf("StringId") == t.length - 8) {
						found = 0;
						for (r in f) {
							if (r == t.substring(0, t.indexOf("StringId"))) {
								found++;
							}
							if (r == t.substring(0, t.indexOf("StringId")) + "Id") {
								found++;
							}
						}
						if (found >= 1) {
							delete f[t];
							continue;
						}
					}
					if (~t.indexOf("Id") && t.indexOf("Id") == t.length - 2) {
						found = 0;
						for (r in f) {
							if (r == t.substring(0, t.indexOf("Id"))) {
								found++;
							}
						}
						if (found == 1) {
							delete f[t.substring(0, t.indexOf("Id"))];
							continue;
						}
					}
				} catch (e) {}
			}

			// run additional on save commands
			if (additSave)
				additSave(f, allowInvalid, $scope.routeParams.Stage, $scope);
			// ensure form is still valid
			if (!$scope.form.valid && !allowInvalid)
				return $scope.submitted = true;
			
			// fix data types
			for (var x in f) {
				try {
					// delete any that are uncahnged
					if (x != "Id" && $scope.form2 && $scope.form2[x] == f[x]) {
						delete f[x];
						continue;
					}

					// if setting field null then keep
					if (!f[x])
						continue;
					
					// ensure no saving deferred
					if (f[x].__deferred) {
						delete f[x];
						continue;
					}
					
					// dates back to string
					if (f[x].toJSON)
						f[x] = f[x].toJSON();
					
					// people back to array of Id or single Id
					if (~x.indexOf('Id') && typeof f[x] == "object" && f[x].results)
						f[x] = {
							__metadata: {type: "Collection(Edm.Int32)"},
							results: f[x].results.length > 0 && f[x].results[0].Id ? f[x].results.map(function(x) {return x.Id || x;})
								: f[x].results
						};
					else if (~x.indexOf('Id') && typeof f[x] == "object" && f[x].length > 0)
						f[x] = {
							__metadata: {type: "Collection(Edm.Int32)"},
							results: f[x].length > 0 && f[x][0].Id ? f[x].map(function(x) {return x.Id || x;})
								: f[x]
						};
					else if (~x.indexOf('Id') && typeof f[x] == "object")
						f[x] = f[x].Id || f[x];
					
					// not an sp object under results, i.e. not a multi field
					if (typeof f[x] == "object" && !f[x].results && x != "__metadata" && !f[x].__deferred)
						f[x] = JSON.stringify(f[x]);
				} catch (e) {}
			}
			
			// disable buttons and begin saving item then attachments then close
			if (whatIf)
				deferred.resolve('Simulated');
			else {
				$scope.processing = true;
				listSvc.postItem(overrideList || listName, f)
					.then (function (d) {
						if (d && d.d && d.d.Id)
							$scope.form.Id = d.d.Id;
						if (!$scope.Files || $scope.Files.length < 1) {
							$scope.submitted = 'Saved';
							$scope.processing = false;
							if (!dontClose)
								$scope.close();
							else if (dontClose === true)
								$scope.$apply();
							return deferred.resolve('Saved');
						}

						listSvc.upload(overrideList || listName, $scope.form.Id, $scope.Files)
							.then(function () {
								$scope.submitted = 'Saved';
								$scope.processing = false;
								if (!dontClose)
									$scope.close();
								else if (dontClose === true)
									$scope.$apply();
								return deferred.resolve('Saved');
							}, function (a, b) {
								$scope.submitted = b || 'Attachments failed';
								$scope.processing = false;
								return deferred.reject('Attachments failed');
							});
					}, function (a, b) {
						$scope.submitted = b || 'Save failed';
						$scope.processing = false;
						return deferred.reject('Save failed');
					});
			}
			
			return deferred.promise;
		}
		
		// custom app specifics here
		if (additFuncs)
			additFuncs($scope, webSvc, listSvc, $uibModal, $timeout, $q, fileSvc);
	})
})(APP$ || $);
