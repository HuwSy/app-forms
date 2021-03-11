(function () {
	var appMod = angular.module('people', []);
	
    appMod.directive('people',
        function () {
            /// <summary>Renders a SharePoint people picker compatible with ng-model</summary>
            /// NOTE: mdl name might need to be .results if its an array and not handeled elsewhere
            return {
                restrict: 'E',
				require: 'ngModel',
                template: '<div>This control has not loaded</div>',
				link: function (scope, element, attrs, controller) {
					// step to parent scope, hack for some modal tech only
					if (attrs['parent'] == 'true')
						scope = scope.$parent;
					
					// change the output format to return non array, if array is false multi will be false. default of true
					var array = !(attrs['array'] == 'false');
                    // allow multi, can not be changed after init, default is multi="true". Note: Even in single user mode this returns an array unless overridden below
					var multi = !array ? false : !(attrs['multi'] == 'false');
					
					// include search text in output results, can not be changed after init
					var searches = attrs['searches'] == 'true';
					
                    // allow what type of users, can not be changed after init
				    var avail = attrs['groups'] == 'all' ? 'User,DL,SecGroup,SPGroup' // all possible users
						: attrs['groups'] == 'security' ? 'User,SecGroup,SPGroup' // securable set of objects
						: attrs['groups'] == 'email' ? 'User,DL,SPGroup' // email enabled set of objects
						: attrs['groups'] == 'sharepoint' ? 'User,SPGroup' // SP queriable objects
						: attrs['groups'] != null && parseInt(attrs['groups']).toString() == "NaN" ? attrs['groups'] // allow custom combinations
						: 'User'; // default to users
					var grp = parseInt(attrs['groups']).toString() != "NaN" ? parseInt(attrs['groups']) : null;
						
					// users to this instance
					var users = [];
					
				    // read the model item/path and create a reference to the array ready
				    // use attrs and eval not controller as it doesnt interfear with sp client people picker
					var mdl = attrs['ngModel'] || attrs['dataNgModel'];
					var ensureMdl = function () {
						var mdlSplit = mdl.split('.');
						if (mdlSplit.length > 1) {
							var soFar = 'scope';
							for (var h = 0; h < mdlSplit.length - 1; h++) {
								soFar += "." + mdlSplit[h]
								if (eval(soFar) == null)
									eval(soFar + ' = {};')
							}
						}
					}

                    // is disabled
					var dis = attrs['disabled'] == "disabled" || attrs['disabled']
						|| (attrs['disabled'] || attrs['ngDisabled'] || attrs['dataNgDisabled']) == "true";
                    showHide();

                    // is disabled off scope object
                    if (!dis && (attrs['ngDisabled'] || attrs['dataNgDisabled']) != null) {
						dis = true;
                        scope.$watch(attrs['ngDisabled'] || attrs['dataNgDisabled'],
                            function () {
								var r = (attrs['ngDisabled'] || attrs['dataNgDisabled']);
								// issues, will not work with mixed objects
								// i.e. form.Test && admin, but would work with form.Test && form.Admin
								if (~r.indexOf('.')) {
									dis = eval(r.replace(/([A-Za-z][A-Za-z0-9]+\.)/g, 'scope.$1'));
								} else {
									dis = eval(r.replace(/([A-Za-z][A-Za-z0-9]+)/g, 'scope.$1'));
								}
                                showHide();
                            });
                    }

                    // shows or hides elements based on disabled, read only div if disabled
                    function showHide() {
                        if (dis) {
                            var styles = angular.element(element).attr('style');
                            angular.element(element).hide();
                            var display = angular.element(element).next();
                            if (display.length == 0 || display[0].id != 'people-display-mode') {
                                angular.element(element).after('<span id="people-display-mode" style="min-height: 20px;'+styles+'"></span>');
                                display = angular.element(element).next();
                            }
                            if (display.length >= 1 && display[0].id == 'people-display-mode') {
                                var html = '';
								// prefer model for display but fall back to array
								var displayArray = [];
								try {
									displayArray = eval('scope.' + mdl) || [];
								} catch (e) {}
								if (typeof(displayArray) != "object" || displayArray.length == null)
									displayArray = [displayArray];
								if (displayArray.length == 0 || typeof(displayArray[0]) != "object")
									displayArray = users;
								// show users at that time
                                displayArray.forEach(function (u) {
                                    html += (u.DisplayText || u.Title || u) + '; ';
                                });
                                display.html(html.substring(0, html.lastIndexOf(';')));
                            }
                        } else {
                            angular.element(element).show();
                            var display = angular.element(element).next();
                            if (display.length >= 1 && display[0].id == 'people-display-mode') {
                                display.remove();
                            }
                        }
                    }

                    // is required off html required
					var req = attrs['required'] == "required" || attrs['required']
						|| (attrs['required'] || attrs['ngRequired'] || attrs['dataNgRequired']) == "true";

				    // is required off scope
				    if (!req && (attrs['ngRequired'] || attrs['dataNgRequired']) != null) {
						req = true;
				        scope.$watch(attrs['ngRequired'] || attrs['dataNgRequired'],
                            function () {
								var r = (attrs['ngRequired'] || attrs['ngRequired']);
								// issues, will not work with mixed objects
								// i.e. form.Test && admin, but would work with form.Test && form.Admin
								if (~r.indexOf('.')) {
									req = eval(r.replace(/([A-Za-z][A-Za-z0-9]+\.)/g, 'scope.$1'));
								} else {
									req = eval(r.replace(/([A-Za-z][A-Za-z0-9]+)/g, 'scope.$1'));
								}
								setValidation();
				            });
				    }

                    // set failed or not based on users on scope and required
					function setValidation () {
						var mdlUsers = [];
						try {
							mdlUsers = eval('scope.' + mdl) || [];
						} catch (e) {}
						if (typeof(mdlUsers) != "object" || mdlUsers.length == null)
							mdlUsers = [mdlUsers];
						var invalid = (mdlUsers.length == 0 && req && !searches)
							|| (mdlUsers.length == 0 && req && searches && (eval('scope.' + mdl + '_EditorInput') || '') == '');
						controller.$setValidity('required', !invalid);
					}

				    // ensure the element always has a unique id and name as required by the sp client control and required
					if (attrs.id == null) {
						attrs.id = 'PeoplePicker_' + mdl.replace(/\./g, '_') + "_" + Math.random().toString().split('.')[1];
						element.attr('id', attrs.id);
						element.attr('name', attrs.id);
					}
					
					// ensure the scripts are registered
					if (LoadSodByKey('clientpeoplepicker.js', null) == Sods.missing) {
						RegisterSod('clientpeoplepicker.js', '/_layouts/15/clientpeoplepicker.js');
					}
					if (LoadSodByKey('clientforms.js', null) == Sods.missing) {
						RegisterSod('clientforms.js', '/_layouts/15/clientforms.js');
					}
					if (LoadSodByKey('clienttemplates.js', null) == Sods.missing) {
						RegisterSod('clienttemplates.js', '/_layouts/15/clienttemplates.js');
					}
					if (LoadSodByKey('autofill.js', null) == Sods.missing) {
						RegisterSod('autofill.js', '/_layouts/15/autofill.js');
					}
					
					// ensure the scripts are loaded
					SP.SOD.loadMultiple(["sp.js", "sp.runtime.js", "clienttemplates.js", "clientforms.js", "clientpeoplepicker.js", "autofill.js"], function () {
						// setup options
						var schema = {
							PrincipalAccountType: avail,
							SearchPrincipalSource: 15,
							ResolvePrincipalSource: 15,
							AllowMultipleValues: multi,
							MaximumEntitySuggestions: 2
						};
						if (grp)
							schema['SharePointGroupID'] = grp;
						
						// on scope change
						scope.$watch(mdl, function() {
							setTimeout(function () {
								if (document.getElementById(attrs.id + '_TopSpan_EditorInput') == null)
									return;
	
								setValidation();
	
								var picker = SPClientPeoplePicker.SPClientPeoplePickerDict[attrs.id + '_TopSpan'];
								var mdlUsers = [];
								try {
									mdlUsers = eval('scope.' + mdl) || [];
								} catch (e) {}
								if (typeof(mdlUsers) != "object" || mdlUsers.length == null)
									mdlUsers = [mdlUsers];
	
								// removes
								picker.IterateEachProcessedUser(function (index, usr) {
									if (!~mdlUsers.map(function (x) {return x.Id || x}).indexOf(usr.UserInfo.Id))
										picker.DeleteProcessedUser(document.getElementById(usr.UserContainerElementId));
								});
								
								// adds
								users = [];
								mdlUsers.forEach(function (usr) {
									if (usr.Name) {
										if (!~picker.GetAllUserInfo().map(function (x) {return x.Key}).indexOf(usr.Name)) {
											document.getElementById(attrs.id + '_TopSpan_EditorInput').value = usr.Name;
											picker.AddUnresolvedUserFromEditor(true);
											users.push(usr.Name);
											showHide();
										}
									} else {
										if (!~picker.GetAllUserInfo().map(function (x) {return x.Id}).indexOf(usr.Id || usr)) {
											var user = context.get_web().getUserById(usr.Id || usr);
											context.load(user);
											context.executeQueryAsync(function () {
												document.getElementById(attrs.id + '_TopSpan_EditorInput').value = user.get_loginName();
												picker.AddUnresolvedUserFromEditor(true);
												users.push(user.get_title());
												showHide();
											});
										}
									}
								});

								setTimeout(showHide, 250);
							},10);
						});
						
						// load users from model and convert to resolved usable array for init
						var mdlUsers = [];
						try {
							mdlUsers = eval('scope.' + mdl) || [];
						} catch (e) {}
						if (typeof(mdlUsers) != "object" || mdlUsers.length == null)
							mdlUsers = [mdlUsers];
							
						var context = new SP.ClientContext.get_current();
						var remaining = mdlUsers.length;
						
						function updateUsr(usr) {
							if (usr != null)
								users.push(usr);
							remaining--;
                            if (remaining <= 0)
								setTimeout(init, 10);
						}
						
						if (remaining == 0)
							return updateUsr();

						mdlUsers.forEach(function (usr) {
							if (typeof(usr) == 'object') {
								if (usr.Title)
									usr = {
										AutoFillDisplayText: usr.Title,
										AutoFillKey: usr.Name,
										AutoFillSubDisplayText: "",
										AutoFillTitleText: usr.Title,
										Description: usr.Title,
										DisplayText: usr.Title,
										EntityType: "User",
										IsResolved: true,
										Key: usr.Name,
										ProviderDisplayName: "Tenant",
										ProviderName: "Tenant",
										Resolved: true,
										Id: usr.Id
									}
								updateUsr(usr);
							}
							else {
								var user = context.get_web().getUserById(usr);
								context.load(user);
								context.executeQueryAsync(function () {
									usr = {
										AutoFillDisplayText: user.get_title(),
										AutoFillKey: user.get_loginName(),
										AutoFillSubDisplayText: "",
										AutoFillTitleText: user.get_title(),
										Description: user.get_title(),
										DisplayText: user.get_title(),
										EntityType: "User",
										IsResolved: true,
										Key: user.get_loginName(),
										ProviderDisplayName: "Tenant",
										ProviderName: "Tenant",
										Resolved: true,
										Id: user.get_id()
									}
									updateUsr(usr);
								}, function (s, e) {
									updateUsr();
								});
							}
						});

						// init the field
						function init() {
							// add the field with the resolved users
							SPClientPeoplePicker_InitStandaloneControlWrapper(attrs.id, users.length > 0 ? users : null, schema);
							
                            // mark as touched and dirt if clicked
							element.on('click', function() {
								controller.$setTouched();
								controller.$setDirty();
							});
							
							// retrigger the disabled version view-only html version if required
							showHide();
                            
                            // on blur update editor input model relative to actual model, occasionally used for person not found
							setTimeout(function () {
								if (document.getElementById(attrs.id + '_TopSpan_EditorInput') != null) {
									document.getElementById(attrs.id + '_TopSpan_EditorInput').setAttribute('autocomplete', 'off-' + attrs.id);
									document.getElementById(attrs.id + '_TopSpan_EditorInput').onblur = function () {
										if (searches) {
											ensureMdl();
											eval('scope.' + mdl + '_EditorInput = this.value');
											setValidation();
										}
									};
								}
							}, 100);
							
							// on field change update the model
							var delaySPLoop;
							SPClientPeoplePicker.SPClientPeoplePickerDict[attrs.id + '_TopSpan']
								.OnUserResolvedClientScript = function (peoplePickerId, selectedUsersInfo) {
									var notAllFound = false;
									var remaining = selectedUsersInfo.length;
									function updateId () {
										remaining--;
										if (remaining <= 0 && !notAllFound)	{
											clearTimeout(delaySPLoop);
											delaySPLoop = setTimeout(function () {
												ensureMdl();
												if (array)
													eval('scope.' + mdl + ' = selectedUsersInfo');
												else
													eval('scope.' + mdl + ' = selectedUsersInfo[0]');
												scope.$apply();
											}, 500);
										}
									}

									if (remaining == 0)
									    updateId();

									selectedUsersInfo.forEach(function (u) {
										if (true) { //u.Id == null // always search as u.Id can become corrupted
											var user = context.get_web().ensureUser(u.Key);
											context.load(user);
											context.executeQueryAsync(function () {
												u.Id = user.get_id();
												if (!u.Resolved && !searches)
													notAllFound = true;
												updateId();
											}, function (s, e) {
												if (!searches)
													notAllFound = true;
												updateId();
											});
										} else {
											updateId();
										}
									});
								};
						}
					});
                }
            }
        });
})();