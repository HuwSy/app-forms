'use strict';
// start loading immediately
if (typeof SP != "undefined")
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', null);

// sp pnp js
// https://pnp.github.io/pnpjs/sp/

// all list operations
{
    Lists = Lists || {};
    {
        Lists.InvokeRibbon = function () {
            /// <summary>Invokes the library ribbon controls when webparts have been adjusted on a library page</summary>
            // now present in override if its in use
            var elem = $('div[id^="MSOZoneCell_WebPart"]:not(.s4-wpcell-plain)')[0];
            if (elem != null) {
                try {
                    var dummyevent = {
                        target: elem,
                        srcElement: elem
                    }
                    WpClick(dummyevent);
                    _ribbonStartInit("Ribbon.Browse", true);
                    var tabl = $(elem).find('table[onmouseover]')[0];
                    EnsureSelectionHandler(dummyevent, tabl, tabl.getAttribute('onmouseover').match(/[0-9]+/));
                } catch (e) {
                }
            }
        }

        Lists.OverrideDelete = function (callback) {
            /// <summary>overrides the delete key on libraries or lists to specific callback function</summary>
            setInterval(function () {
                $('tr').on('keydown', function (e) {
                    if (e.which === 46) {
                        e.preventDefault();
                        e.stopImmediatePropagation();
                        callback();
                    }
                });
            }, 500);
        }

        Lists.UpdateDigest = function (suppress, force) {
            /// <summary>Updates the reqest digest on the page for any data post/put</summary>
            /// <returns type="String">Deferred object returning a copy of the digest value</returns>

            var def = $.Deferred();

            try {
                if (!force) {
                    // dont spam digest
                    var d = new Date(document.getElementById("__REQUESTDIGEST").value.split(',')[1]);
                    d.setMinutes(d.getMinutes()+1);
                    if (d > new Date()) {
                        def.resolve(document.getElementById("__REQUESTDIGEST").value);
                        return def.promise();
                    }
                }
            } catch (e) {}

            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl + "/_api/contextinfo",
                method: "POST",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose"
                },
                success: function (result) {
                    var i = document.getElementById("__REQUESTDIGEST");
                    if (i == null) {
                        // make this on modern pages as its useful to have less requests to digest
                        i = document.createElement('input');
                        i.type = "hidden";
                        i.id = "__REQUESTDIGEST";
                        document.body.appendChild(i);
                    }
                    i.value = result.d.GetContextWebInformation.FormDigestValue;
                    return def.resolve(result.d.GetContextWebInformation.FormDigestValue);
                },
                error: function (result, status) {
                    if (!suppress)
                        alert('Error (digest) with status: ' + status);
                    return def.reject(status);
                }
            });

            return def.promise();
        };

        Lists.RemoveListItem = function (listName, id, isSite) {
            /// <summary>Removes a list item</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="id" type="Number">The list item .Id (not guid)</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Number">Deferred object</returns>

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_api/web/lists/getbytitle('" +
                listName +
                "')/items(" +
                id +
                ")/recycle()";

            return Lists.UpdateDigest()
                .then(function (d) {
                    return $.ajax({
                        url: endPoint,
                        type: "POST",
                        headers: {
                            "ACCEPT": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": d,
                            "IF-MATCH": "*",
                            "X-HTTP-Method": "DELETE"
                        }
                    });
                });
        }

        Lists.SetListItem = function (listName, items, id, isSite) {
            /// <summary>Sets a list item, new or update</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="items" type="Object">The list item(s) object required for rest, i.e. must include _meta etc. If array passed do all before resolve, if adding item with Folder.Path this will use _vti_bin which formats input differently but should be avoided</param>
            /// <param name="id" type="Number">Optional, the list item .Id (not guid) to update, or null to create new</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Number">Deferred object</returns>

            var def = $.Deferred();

            // null item then resolve
            if (!items || items.length == 0) {
                def.resolve();
                return def.promise();
            }

            // not an array make an array so all processed together
            if (typeof items.length == "undefined") {
                items = [items];
            }

            var promises = [];
            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_api/web/lists/getbytitle('" +
                listName +
                "')/items";
            
            var vtiEndPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_vti_bin/listdata.svc/" + listName.replace(/ /g,'');

            Lists.UpdateDigest().then(function (d) {
                items.forEach(function (item) {
                    var url = endPoint;
                    if (item.Folder && item.Folder.Path) {
                        item.Path = item.Folder.Path;
                        delete item.Folder;
                        delete item.__metadata;
                        url = vtiEndPoint;
                    } else if (!item.__metadata) {
                        item.__metadata = {type: 'SP.Data.' + listName.replace(/ /g, '_x0020_') + 'ListItem'};
                    }

                    promises.push($.ajax({
                        url: url + ((id || item.Id) > 0 ? '(' + (id || item.Id) + ')' : ''),
                        type: "POST",
                        data: JSON.stringify(item),
                        processData: false,
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": d
                        },
                        beforeSend: function (request) {
                            if (id || item.Id) {
                                request.setRequestHeader("IF-MATCH", "*");
                                request.setRequestHeader("X-HTTP-Method", "MERGE");
                            }
                        }, success: function (d) {
                            if (items.length == 1)
                                return def.resolve(d);
                        }, fail:  function (a, b) {
                            return def.reject(a, b);
                        }
                    }));
                });

                if (items.length > 1) {
                    $.when(promises).then(function () {
                        return def.resolve();
                    },function (a, b) {
                        return def.reject(a, b);
                    });
                }
            },function (a, b) {
                return def.reject(a, b);
            });

            return def.promise();
        };

        Lists.SetView = function (url, list, fields, useModernLists) {
            /// <summary>Sets the view field order, can duplicate fields so use clear correctly</summary>
            /// <param name="view" type="String">view to update fields on, or null to not update a view</param>
            /// <param name="fields" type="String">comma delimeted fields in order, docicon should always be first</param>
            /// <param name="clear" type="String">comma delimeted fields to remove before adding, docicon should always be first</param>
            /// <returns type="Number">Deferred object only if ctx is null</returns>
            var def = $.Deferred();
            var ctx = new SP.ClientContext(url || _spPageContextInfo.webServerRelativeUrl);

            var l = ctx.get_web().get_lists().getByTitle(list);
            var v = l.get_defaultView();
            var f = v.get_viewFields();
            var p = l.get_fields();
            ctx.load(f);
            ctx.load(p);

            if (useModernLists === true || useModernLists === false) {
                l.set_listExperienceOptions(useModernLists ? SP.ListExperience.newExperience : SP.ListExperience.classicExperience);
                l.update();
            }
            
            var setViewAdd = function () {
                if (fields) {
                    var cur = [];
                    var q = p.getEnumerator();
                    while (q.moveNext()) {
                        cur.push(q.get_current().get_staticName());
                    }
                    fields.replace('DocIcon,','').split(',').forEach(function (x) {
                        if (x != '' && ~cur.indexOf(x))
                            f.add(x);
                    });
                    v.update();
                }

                ctx.executeQueryAsync(function () {
                    return def.resolve();
                }, function (a, b) {
                    return def.reject(list, a, b, fields, cur);
                });
            }

            var setViewDelete = function () {
                if (fields) {
                    var clear = [];
                    var e = f.getEnumerator();
                    while (e.moveNext()) {
                        clear.push(e.get_current());
                    }
                    clear.forEach(function (x) {
                        if (x != '' && x != 'DocIcon')
                            f.remove(x);
                    });
                    v.update();
                }
                
                ctx.executeQueryAsync(setViewAdd, setViewAdd);
            }

            ctx.executeQueryAsync(setViewDelete, setViewDelete);

            return def.promise();
        }

        Lists.SetFieldDefault = function (listName, fieldName, isSite, defaultValue) {
            if (!listName)
                isSite = true;

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl)
                .replace(/\/$/, "") +
                "/_api/web/" +
                (listName ? "lists/getbytitle('" + listName + "')/" : "") +
                "Fields/GetByTitle('" +
                fieldName +
                "')";

            var def = $.Deferred();

            if (!fieldName || fieldName == "null" || typeof defaultValue == "undefined")
                def.resolve();
            else
                Lists.GetField(listName, fieldName, isSite, 'TypeAsString')
                    .then(function (f) {
                        Lists.UpdateDigest()
                            .then(function (d) {
                                $.ajax({
                                    url: endPoint,
                                    type: "POST",
                                    data: "{ '__metadata': { 'type':'SP." + (f.TypeAsString == "TaxonomyField" || f.TypeAsString == "TaxonomyFieldType" ? ('Taxonomy.' + 'TaxonomyField') : ('Field' + f.TypeAsString)) + "' }, 'DefaultValue': '" + Terms.ToString(defaultValue) + "' }",
                                    headers: {
                                        "Accept": "application/json;odata=verbose",
                                        "X-HTTP-Method": "MERGE",
                                        "content-type": "application/json;odata=verbose",
                                        "X-RequestDigest": d,
                                        "If-Match": "*"
                                    },
                                    success: function () {
                                        return def.resolve();
                                    },
                                    error: function () {
                                        return def.reject();
                                    }
                                });
                            });
                    });

            return def.promise();
        }

        Lists.GetField = function (listName, fieldName, isSite, selects, orderBy) {
            /// <summary>Get field details</summary>
            /// <param name="listName" type="String">List title</param>
            /// <param name="fieldName" type="String">Field title</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <param name="selects" type="String">Filter elements</param>
            /// <returns type="Object">Deferred returning {FieldTitle, + any requested i.e. SspId, TermSetId}</returns>

            if (!listName)
                isSite = true;

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl)
                .replace(/\/$/, "") +
                "/_api/web/" +
                (listName ? "lists/getbytitle('" + listName + "')/" : "") +
                "Fields" +
                (fieldName ? "/GetByTitle('" + fieldName + "')" : '') +
                '?' +
                (selects ? '$select=' + selects : '') +
                '&' +
                (orderBy ? '$orderby=' + orderBy : '');

            var def = $.Deferred();

            $.ajax({
                url: endPoint,
                type: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                },
                success: function (d) {
                    d.d.FieldTitle = fieldName;
                    return def.resolve(d.d);
                },
                error: function () {
                    return def.reject();
                }
            });

            return def.promise();
        }

        Lists.GetCaml = function (listName, caml, isSite) {
            /// <summary>Get a set of list items</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="caml" type="String">CAML query</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred object returning rest formatted  as {return}.d.results</returns>

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl)
                .replace(/\/$/, "") +
                "/_api/web/lists/getbytitle('" +
                listName +
                "')/GetItems(query=@v1)?@v1={\"ViewXml\":\"" +
                caml +
                "\"}";

            return Lists.UpdateDigest()
                .then(function (d) {
                    return $.ajax({
                        url: endPoint,
                        type: "POST",
                        headers: {
                            "accept": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": d
                        }
                    });
                });
        }

        Lists.GetItemHistory = function (listName, id) {
            var endPoint = _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "") +
                "/_api/web/lists/getbytitle('" + listName + "')/items(" + id + ")/versions";

            return $.ajax({
                url: endPoint,
                type: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                }
            });
        };

        Lists.GetListItem = function (listName, filter, select, isSite, expand, orderby, top) {
            /// <summary>Get a set of list items</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="filter" type="String">Rest formatted filter parameters, if filtering on Folder/Path this will use _vti_bin which formats output differently but should be avoided</param>
            /// <param name="select" type="String">Rest formatted select parameters, may be overridden with &$orderby etc</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred object returning rest formatted as {return}.d.results</returns>

            if (select != null && (expand == null || expand == '') && ~select.indexOf('/')) {
                var x = [];
                select.split(',').forEach(function (s) {
                    var e = s.trim().split('/');
                    if (e.length == 2 && e[0] != "" && !~x.indexOf(e[0]))
                        x.push(e[0])
                });
                expand = x.join(',');
            }

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_api/web/lists" +
                (listName ? "/getbytitle('" + listName + "')/items" : "") +
                "?" +
                (filter ? '&$filter=' + filter : '') +
                (select ? '&$select=' + select : '') +
                (expand ? '&$expand=' + expand : '') +
                (orderby ? '&$orderby=' + orderby : '') +
                (top ? '&$top=' + top : '&$top=5000');

            var vtiEndPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_vti_bin/listdata.svc" +
                (listName ? "/" + listName.replace(/ /g,'') : "") +
                "?" +
                (filter ? '&$filter=' + filter.replace(/Folder\/Path/g,'Path') : '') +
                (select ? '&$select=' + select : '') +
                (expand ? '&$expand=' + expand : '') +
                (orderby ? '&$orderby=' + orderby : '') +
                (top ? '&$top=' + top : '&$top=5000');

            return $.ajax({
                url: filter && ~filter.indexOf("Folder/Path") ? vtiEndPoint : endPoint,
                type: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                }
            });
        };

        Lists.UploadAttachments = function(listName, itemId, Files) {
            var def = $.Deferred();
			
            if (Files == null) {
                def.resolve();
                return def.promise();
            }

            var dl = function (i) {
                if (i >= Files.length)
                    return def.resolve();
                var a = Files[i];
                if (!a.Deleted && a.ServerRelativeUrl == null && a.Data)
                    Lists.AddListAttachments(listName, itemId, a.FileName, a.Data).then(function (d) {
                        a.ServerRelativeUrl = d.d.ServerRelativeUrl;
                        delete a.Data;
                        dl(i+1);
                    },function () {
                        dl(i+1);
                    });
                else
                    dl(i+1);
            }

            var ad = function (i) {
                if (i >= Files.length)
                    return dl(0);
                var a = Files[i];
                if (a.Deleted && a.ServerRelativeUrl != null)
                    Lists.DelListAttachments(listName, itemId, a.FileName).then(function () {
                        a.ServerRelativeUrl = null;
                        ad(i+1);
                    },function () {
                        ad(i+1);
                    });
                else
                ad(i+1);
            }

            ad(0);

			return def.promise();
        }

        Lists.GetListAttachments = function (listName, id, isSite) {
            /// <summary>Get list item attachments to an items</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="id" type="String">List item Id</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred object returning rest formatted as {return}.d.results</returns>

            var endPoint = (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") +
                "/_api/web/lists/getbytitle('" +
                listName +
                "')/items(" +
                id +
                ")/AttachmentFiles/";

            return $.ajax({
                url: endPoint,
                type: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                }
            });
        };

		Lists.Permissions = function (folder, limitLevel) {
			var def = $.Deferred();
			
			Lists.UpdateDigest().then(function (d) {
				$.ajax({
					method: 'POST',
					url: _spPageContextInfo.webAbsoluteUrl +
						"/_api/Web/GetFolderByServerRelativeUrl('" +
						folder +
						"')?$expand=ListItemAllFields/RoleAssignments/Member,ListItemAllFields/RoleAssignments/RoleDefinitionBindings,ListItemAllFields/RoleAssignments/Member/Users",
					headers: {
						"accept": "application/json;odata=verbose",
						"X-RequestDigest": d,
						"content-Type": "application/json;odata=verbose"
                    },
                    success: function(roles) {
                        if (roles.d.ListItemAllFields.RoleAssignments == null)
                            return def.resolve([]);
                        return def.resolve(
                            _.filter(
                                roles.d.ListItemAllFields.RoleAssignments.results,
                                function (role) {
                                    return role.Member.LoginName.indexOf('i:') === 0
                                        && (limitLevel == null || role.RoleDefinitionBindings.results.map(function (x) { return x.Name; }).indexOf(limitLevel) >= 0);
                                })
                        );
                    },
                    error: function() {
                        return def.reject();
                    }
				});
			});

			return def.promise();
		}
		
		Lists.AddPermissions = function (ctx, list, id, loginNames, level) {
			var def = $.Deferred();
			if (ctx == null || list == null || id == null || loginNames == null) {
				return def.resolve();
			}
			
			if (level == null)
				level = 'Contribute';
			
			if (typeof(loginNames) == "string")
				loginNames = [loginNames];
			
            var role = new SP.RoleDefinitionBindingCollection(ctx);
			var level = ctx.get_web().get_roleDefinitions().getByName(level);
			var lists = ctx.get_web().get_lists().getByTitle(list);
			var item = lists.getById(id);

			ctx.load(role);
			ctx.load(level);
			ctx.load(lists);
			ctx.load(item, 'HasUniqueRoleAssignments');

			ctx.executeQueryAsync(function() {
				if (!item.get_hasUniqueRoleAssignments())
					item.breakRoleInheritance(false, false);
				
				ctx.load(item);
				ctx.executeQueryAsync(function() {
					role.add(level);
					for (var i = 0; i < loginNames.length; i++)
						item.get_roleAssignments().add(ctx.get_web().ensureUser(loginNames[i]), role);
					//item.update();
					ctx.executeQueryAsync(function() {
							return def.resolve();
						},
						function() {
							return def.reject();
						});
				},
				function() {
					return def.reject();
				});
			},
			function() {
				return def.reject();
			});

			return def.promise();
		}

        Lists.DelPermissions = function(ctx, list, id, loginNames) {
			var def = $.Deferred();
			if (ctx == null || list == null || id == null || loginNames == null) {
				return def.resolve();
			}
			
			if (typeof(loginNames) == "string")
				loginNames = [loginNames];
			
			var lists = ctx.get_web().get_lists().getByTitle(list);
			var item = lists.getById(id);

			ctx.load(lists);
			ctx.load(item, 'HasUniqueRoleAssignments');

			ctx.executeQueryAsync(function() {
				if (!item.get_hasUniqueRoleAssignments())
					return def.reject();
				
				ctx.load(item);
				ctx.executeQueryAsync(function () {
					item.get_roleAssignments().getByPrincipal(ctx.get_web().ensureUser(loginName)).deleteObject();
					ctx.executeQueryAsync(function () {
							return def.resolve();
						},
						function () {
							return def.reject();
						});
				},
				function () {
					return def.reject();
				});
			},
			function() {
				return def.reject();
			});

			return def.promise();
		}
		
		Lists.UrlExists = function (url) {
            var def = $.Deferred();
            
            Lists.UrlContent(url).then(function (s) {
                def.resolve(s);
            }, function (s) {
                def.reject(s);
            })

			return def.promise();
		}
		
		Lists.UrlContent = function (url) {
			var def = $.Deferred();
			var xhttp = new XMLHttpRequest();
			xhttp.onreadystatechange = function () {
				if (xhttp.readyState === 4) {
					if (xhttp.status === 401 || ~(xhttp.responseURL || '').indexOf('/15/AccessDenied.aspx') || ~(xhttp.response || '').indexOf('Access required'))
						return def.reject(true);
					if (xhttp.status === 404)
						return def.resolve(false);
					if (xhttp.status === 200)
						return def.resolve(true, xhttp.response);
				}
			};
			xhttp.open("GET", url, true);
			if (url == null || url == '' || url.match(/^undefined$/i) != null)
				def.reject(false);
			else
				xhttp.send();

			return def.promise();
		}

        Lists.GetCurrentFolder = function (ctx, root, invFld, typFld, forceItem) {
            /// <summary>Get current SPFolder (and if the current item is a folder in a list SPListItem, SPList)</summary>
            /// <param name="ctx" type="Object">Client context, or current context if null</param>
            /// <param name="root" type="String">Server Relative of Absolute URL to get, or current if null</param>
            /// <param name="invFld" type="String">Additional field to ensure load from list item, i.e Primary metadata field</param>
            /// <param name="typFld" type="String">Additional field to ensure load from list item</param>
            /// <param name="forceItem" type="Bool">Force load of list item details, will error if root is not a list item</param>
            /// <returns type="Object">Deferred object returning SPFolder, SPLIstItem, SPLIst</returns>

            if (ctx == null) ctx = new SP.ClientContext.get_current();

            // push in random ootb fields so that te function returns the list fine
            if (invFld == null) invFld = 'Title';
            if (typFld == null) typFld = 'Created';

            var def = $.Deferred();

            var list;
            if (root == null && Override.ParameterByName('RootFolder', window.location.href) != null) {
                root = Override.ParameterByName('RootFolder', window.location.href);
                list = _spPageContextInfo.listTitle ? ctx.get_web().get_lists().getByTitle(_spPageContextInfo.listTitle) : null;
            }
            if (root == null && _spPageContextInfo.listUrl != null && _spPageContextInfo.listTitle != "Pages") {
                root = _spPageContextInfo.listUrl;
                list = _spPageContextInfo.listTitle ? ctx.get_web().get_lists().getByTitle(_spPageContextInfo.listTitle) : null;
            }
            if (root == null && !_spPageContextInfo.serverRequestPath.toLowerCase().split('/forms/')[0].includes("pages")) {
                root = _spPageContextInfo.serverRequestPath.toLowerCase().split('/forms/')[0];
            }
            if (root == null) {
                root = ctx.get_url();
            }

            var currentFolder = ctx.get_web().getFolderByServerRelativeUrl(root);
            ctx.load(currentFolder);

            var currentFolderItem;
            if (forceItem || Override.ParameterByName('RootFolder', window.location.href) == root) {
                currentFolderItem = currentFolder.get_listItemAllFields();
                ctx.load(currentFolderItem, 'ContentTypeId', invFld || 'Title', typFld || 'Title', 'FileRef', 'FileLeafRef', 'ServerUrl', 'FileDirRef', 'Modified', 'ContentType', 'HasUniqueRoleAssignments');
            }

            if (currentFolderItem != null && list == null)
                list = currentFolderItem.get_parentList();

            if (list != null) {
                ctx.load(list);
                ctx.load(list.get_rootFolder());
            }

            ctx.executeQueryAsync(function () {
                if (list != null || root.toLowerCase() == ctx.get_url().toLowerCase())
                    return def.resolve(currentFolder, currentFolderItem, list);

                Lists.GetListObject(ctx, root)
                    .then(function (list) {
                        return def.resolve(currentFolder, currentFolderItem, list);
                    });
            },
                function () {
                    return def.reject();
                });

            return def.promise();
        }

        Lists.GetSiteContentType = function (name) {
            /// <summary>Get site collection content type (with a new context object so must be used with care)</summary>
            /// <param name="name" type="String">Content type name to locate, case sensitive</param>
            /// <returns type="Object">Deferred object returning content type object</returns>

            var def = $.Deferred();
            if (name == null || name == '' || name.indexOf('0x') == 0) {
                def.resolve(null);
                return def.promise();
            }

            var ctx = new SP.ClientContext(_spPageContextInfo.siteServerRelativeUrl);
            var web = ctx.get_web();
            var contentTypes = web.get_contentTypes();
            ctx.load(contentTypes);

            ctx.executeQueryAsync(function () {
                    var contentTypeArray = contentTypes.get_data();
                    contentTypeArray.forEach(function (ct) {
                        if (ct.get_name() === name)
                            return def.resolve(ct);
                        },
                        this);

                    return def.reject();
                },
                function (request, args) {
                    return def.reject();
                });

            return def.promise();
        }

        Lists.GetInFolder = function (web, folderURL, type) {
            /// <summary>Gets all files in a folder, non recursive</summary>
            /// <param name="web" type="String">Web server relative url, null to use current web</param>
            /// <param name="folderURL" type="String">Server relative URL to folder</param>
            /// <param name="type" type="String">Additional fields to expand in rest format</param>
            /// <returns type="Object">Deferred object returning rest formatted  as {return}.results</returns>

            var def = $.Deferred();
            if (web == null || web === "") {
                web = _spPageContextInfo.webServerRelativeUrl;
            }
            if (folderURL == null || folderURL === "") {
                return def.resolve([]);
            }

            var endPoint = web.replace(/\/$/, "") +
                "/_api/web/GetFolderByServerRelativeUrl('" +
                folderURL +
                "')?$expand=" + type;

            $.ajax({
                url: endPoint,
                type: "GET",
                headers: {
                    "accept": "application/json;odata=verbose"
                },
                success: function (documents) {
                    return def.resolve(documents.d);
                },
                error: function () {
                    return def.reject();
                }
            });

            return def.promise();
        };

        Lists.CreateFolders = function (ctx, list, fldCt, invFld, typFld, folders, blankMoveCopy, documentTypes, termName, termValue, mandateValue, locationURL, noDt) {
            /// <summary>Creates a collection of folders, and folders within</summary>
            /// <param name="ctx" type="Object">Ignored variable, do not use</param>
            /// <param name="list" type="Object">SPList object, context is loaded off this</param>
            /// <param name="fldCt" type="String">Folder content type name to use for all folders</param>
            /// <param name="invFld" type="String">Primary metadata field name to setup on folders</param>
            /// <param name="typFld" type="String">Secondary metadata field name to setup on folders</param>
            /// <param name="folders" type="Object">Object of folders to create as per GetSubFolders function, will include source URLs when blankMoveCopy is Move or Copy</param>
            /// <param name="blankMoveCopy" type="Number">0 = Only create folders, 1 = Move folders within folders parameter, 2 = Copy folders in parameter</param>
            /// <param name="documentTypes" type="Object">Tree version of the Document Type termset via Terms.GetTermSetAsTree</param>
            /// <param name="termName" type="String">Used for building the folder path if the folders object doesnt contain a toLocation</param>
            /// <param name="termValue" type="Object">Default primary term value unless overridden by folders parameter</param>
            /// <param name="mandateValue" type="Object">Default secondary term value unless overridden by folders parameter</param>
            /// <param name="locationURL" type="String">Used for building the folder path if the folders object doesnt contain a toLocation</param>
            /// <returns type="Object">Deferred object</returns>

            /*
            basic usage: pass list object, folders as below, noDt as true
            folders:[
                {"fileLeafRef" : "Admin", "folderRelativeURL" :""},
                {"fileLeafRef" : "Approval Committees", "folderRelativeURL" :""},
                {"fileLeafRef" : "AC", "folderRelativeURL" :"/Approval Committees"},
                {"fileLeafRef" : "PMIC", "folderRelativeURL" :"/Approval Committees"},
                {"fileLeafRef" : "Legal", "folderRelativeURL" :""},
                {"fileLeafRef" : "Transaction Docs-LPA-Side letters", "folderRelativeURL" :"/Legal"}
            ]
            */

            ctx = list.get_context();
            var def = $.Deferred();
            if (folders.length === 0) {
                return def.resolve();
            }

            if (!locationURL || locationURL == '')
                locationURL = list.get_rootFolder().get_serverRelativeUrl();

            if (folders[0].toLocation == null &&
                locationURL != null &&
                termName != null &&
                folders[0].folderRelativeURL != null) {
                folders[0].toLocation = locationURL + '/' + termName + folders[0].folderRelativeURL;
            }

            if (folders[0].toLocation == null &&
                locationURL != null &&
                folders[0].folderRelativeURL != null) {
                folders[0].toLocation = locationURL + folders[0].folderRelativeURL;
            }

            Lists.GetSubFolders(ctx, null, null, null, null, null, list, folders[0].toLocation, noDt)
                .done(function (destinationConflicts) {
                    var destinations = [];
                    for (var i = 0; i < destinationConflicts.length; i++) {
                        destinations.push(decodeURIComponent(destinationConflicts[i].ServerUrl.substring(destinationConflicts[i].ServerUrl.replace(/\/$/, '').lastIndexOf('/'))).toLowerCase());
                        destinations.push(decodeURIComponent(destinationConflicts[i].ServerUrl.replace(/\/$/, '')).toLowerCase());
                    }

                    Lists.GetSiteContentType(fldCt || 'Folder')
                        .done(function (contentType) {
                            for (var i = 0; i < folders.length; i++) {
                                var currentTermValue = termValue;
                                if (folders[i].investments != null) {
                                    currentTermValue = folders[i].investments;
                                }
                                var currentMandateValue = mandateValue;
                                if (folders[i].mandateType != null) {
                                    currentMandateValue = folders[i].mandateType;
                                }

                                if (folders[i].toLocation == null &&
                                    locationURL != null &&
                                    termName != null &&
                                    folders[i].folderRelativeURL != null) {
                                    folders[i].toLocation = locationURL + '/' + termName + folders[i].folderRelativeURL;
                                }

                                if (folders[i].toLocation == null &&
                                    locationURL != null &&
                                    folders[i].folderRelativeURL != null) {
                                    folders[i].toLocation = locationURL + folders[i].folderRelativeURL;
                                }

                                // folder already exists then skip creation
                                if (folders[i].toLocation
                                    && destinations.indexOf(decodeURIComponent(folders[i].toLocation.replace(/\/$/, '') +
                                        '/' +
                                        folders[i].fileLeafRef).toLowerCase()) >=
                                    0)
                                    continue;

                                var folderInfo = new SP.ListItemCreationInformation();
                                folderInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
                                folderInfo.set_leafName(folders[i].fileLeafRef);
                                folderInfo.set_folderUrl(folders[i].toLocation);

                                folders[i].DocumentType = folders[i].DocumentType || folders[i].docType;

                                var folder = list.addItem(folderInfo);
                                folder.set_item('ContentTypeId', folders[i].contentTypeId || contentType.get_id());
                                if (currentTermValue != null)
                                    folder.set_item(invFld, Terms.ToString(currentTermValue));
                                if (currentMandateValue != null)
                                    folder.set_item(typFld, Terms.ToString(currentMandateValue));
                                if (folders[i].DocumentType != null) {
                                    var dt = Terms.Find(documentTypes, folders[i].DocumentType);
                                    if (dt != null)
                                        folder.set_item('DocumentType', Terms.ToString(dt.term));
                                }

                                folder.update();
                            }

                            ctx.executeQueryAsync(function () {
                                Lists.IntoFolders(folders, blankMoveCopy || 0)
                                    .done(function () {
                                        return def.resolve();
                                    })
                                    .fail(function () {
                                        return def.reject();
                                    });
                            },
                                function () {
                                    return def.reject();
                                });
                        })
                        .fail(function () {
                            return def.reject();
                        });
                })
                .fail(function () {
                    return def.reject();
                });

            return def.promise();
        };

        Lists.CreateFolder = function (ctx, list, continueIfExists, folderName, desc, type, invFld, typFld, locationURL, termValue, mandateValue) {
            /// <summary>Creates a folders</summary>
            /// <param name="ctx" type="Object">Ignored variable, do not use</param>
            /// <param name="list" type="Object">SPList object, context is loaded off this</param>
            /// <param name="continueIfExists" type="Boolean">Continue if the folder already exists, else throw deferred.reject()</param>
            /// <param name="folderName" type="String">Folder name to create</param>
            /// <param name="desc" type="String">Not used</param>
            /// <param name="type" type="String">Folder content type name to use for folder</param>
            /// <param name="invFld" type="String">Primary metadata field name to setup on folder</param>
            /// <param name="typFld" type="String">Secondary metadata field name to setup on folder</param>
            /// <param name="locationURL" type="String">Parent folder server relative URL</param>
            /// <param name="termValue" type="Object">Primary term value</param>
            /// <param name="mandateValue" type="Object">Secondary term value</param>
            /// <returns type="Object">Deferred object returning new SPFolder</returns>

            ctx = list.get_context();

            var def = $.Deferred();
            var ext = $.Deferred();

            if (locationURL) {
                var currentFolder = ctx.get_web().getFolderByServerRelativeUrl(locationURL.replace(/\/$/, '') + '/' + folderName);
                ctx.load(currentFolder);
                ctx.executeQueryAsync(function () {
                    try {
                        if (currentFolder.get_exists())
                            return continueIfExists ? def.resolve(currentFolder) : def.reject();
                    } catch (e) {
                    }

                    return ext.resolve();
                },
                function () {
                    return ext.resolve();
                });
            } else
                ext.resolve();

            $.when(ext)
                .then(function () {
                    Lists.GetSiteContentType(type)
                        .done(function (contentType) {
                            var folderInfo = new SP.ListItemCreationInformation();
                            folderInfo.set_underlyingObjectType(SP.FileSystemObjectType.folder);
                            folderInfo.set_leafName(folderName);
                            if (locationURL)
                                folderInfo.set_folderUrl(locationURL);

                            var item = list.addItem(folderInfo);
                            if (contentType && contentType.get_id && contentType.get_id())
                                item.set_item('ContentTypeId', contentType.get_id());
                            else if (type && type != '' && type.indexOf('0x') == 0)
                                item.set_item('ContentTypeId', type);
                            if (desc)
                                item.set_item('DocumentSetDescription', desc);
                            if (invFld && termValue != null)
                                item.set_item(invFld, Terms.ToString(termValue));
                            if (typFld && mandateValue != null)
                                item.set_item(typFld, Terms.ToString(mandateValue));
                            item.update();

                            var folder = item.get_folder();
                            ctx.load(folder);
                            ctx.executeQueryAsync(function () {
                                return def.resolve(folder);
                            },
                                function () {
                                    return def.reject();
                                });
                        })
                        .fail(function () {
                            return def.reject();
                        });
                });

            return def.promise();
        };

        Lists.Listtitle = function (termName) {
            /// <summary>Calculates the A-Z based list title based on the term name</summary>
            /// <param name="termName" type="String">Term name</param>
            /// <returns type="Object">name = list title, path = list path part</returns>

            if (termName.charAt(0) == '/')
                termName = termName.substring(1);

            termName = termName.split('/')[0];

            if (termName.match(/^[A-Z]/i) == null)
                return {
                    name: "0-9#",
                    path: '09'
                };

            return {
                name: termName.toUpperCase().charAt(0),
                path: termName.toUpperCase().charAt(0)
            }
        }

        Lists.GetRecursiveFiles = function (ctx, url, lib, invFld, typFld, trnFld) {
            /// <summary>Gets all files in the specified URL and below, over 5000 items under specified URL will error</summary>
            /// <param name="ctx" type="Object">Client context, or current context if null</param>
            /// <param name="url" type="String">Server Relative of Absolute URL to get</param>
            /// <param name="lib" type="String">Library title to load from</param>
            /// <param name="invFld" type="String">Additional field to ensure load from list item, i.e Primary metadata field</param>
            /// <param name="typFld" type="String">Additional field to ensure load from list item</param>
            /// <param name="trnFld" type="String">Additional bool field to ensure load from list item and used as selector for isSelected</param>
            /// <returns type="Object">Deferred returning documents array</returns>

            if (ctx == null) ctx = new SP.ClientContext.get_current();
            var def = $.Deferred();
            var oLibDocs = ctx.get_web().get_lists().getByTitle(lib);
            var caml = SP.CamlQuery.createAllItemsQuery();
            caml.set_folderServerRelativeUrl(url);

            // where and order by must be indexed fields
            caml.set_viewXml('<View Scope="RecursiveAll">\
                            <Query>\
                                <OrderBy Override="TRUE">\
                                    <FieldRef Name="FileDirRef" Ascending="True" />\
                                    <FieldRef Name="FileLeafRef" Ascending="True" />\
                                </OrderBy>\
                            </Query>\
                            <RowLimit>1000</RowLimit>\
                        </View>');

            var allDocumentsCol = oLibDocs.getItems(caml);
            ctx.load(allDocumentsCol, "Include(FileSystemObjectType, Id, FileLeafRef, ServerUrl, Modified, ContentType, File_x0020_Size, " + (invFld ? invFld + ", " : '') + (typFld ? typFld + ", " : '') + "FileDirRef" + (trnFld ? ", " + trnFld : '') + ")");

            ctx.executeQueryAsync(function () {
                var documents = [];
                var listEnumerator = allDocumentsCol.getEnumerator();
                while (listEnumerator.moveNext()) {
                    if (listEnumerator.get_current().get_fileSystemObjectType() != 0)
                        continue;

                    var d = {
                        id: listEnumerator.get_current().get_id(),
                        FileName: listEnumerator.get_current().get_fieldValues()['FileLeafRef'],
                        ServerUrl: listEnumerator.get_current().get_fieldValues()['ServerUrl'],
                        Location: ("/" + listEnumerator.get_current().get_fieldValues()['FileDirRef'].substring(url.length + 1) + '/').replace('//', '/'),
                        Modified: listEnumerator.get_current().get_fieldValues()['Modified'],
                        ContentType: listEnumerator.get_current().get_fieldValues()['ContentType'],
                        FileSize: Lists.BytesToDisplay(listEnumerator.get_current().get_fieldValues()['File_x0020_Size']),
                        FileBytes: listEnumerator.get_current().get_fieldValues()['File_x0020_Size'],
                        investments: !invFld ? null : listEnumerator.get_current().get_fieldValues()[invFld],
                        mandateType: !typFld ? null : listEnumerator.get_current().get_fieldValues()[typFld],
                        Closing_x0020_Bible: trnFld != null && listEnumerator.get_current().get_fieldValues()[trnFld] || false,
                        isSelected: trnFld != null && listEnumerator.get_current().get_fieldValues()[trnFld] || false
                    };

                    d.FileType = d.FileName.substring(d.FileName.lastIndexOf('.') + 1);

                    documents.push(d);
                }

                return def.resolve(documents);
            },
                function () {
                    return def.reject();
                });

            return def.promise();
        };

        Lists.BytesToDisplay = function (b) {
            if (b > 1048576)
                return Math.round(b / 1048576) + " MB";
            else if (b > 1024)
                return Math.round(b / 1024) + " KB";
            else
                return b + " B";
        }

        Lists.GetSubFolders = function (ctx, folder, invFld, typFld, toWeb, toLocation, oLibDocs, url, noDt) {
            /// <summary>Gets all folders in the specified URL and below, over 5000 items under specified URL will error</summary>
            /// <param name="ctx" type="Object">Client context, or current context if null</param>
            /// <param name="folder" type="Object">SPFolder object</param>
            /// <param name="invFld" type="String">Additional field to ensure load from list item, i.e Primary metadata field</param>
            /// <param name="typFld" type="String">Additional field to ensure load from list item</param>
            /// <param name="toWeb" type="String">Destination web server relative url to add to the toLocation on the folder object for use with CreateFolders</param>
            /// <param name="toLocation" type="String">Destination folder url to add to the toLocation on the folder object for use with CreateFolders</param>
            /// <param name="oLibDocs">Optional use this list from SPList</param>
            /// <param name="url">Optional search under this server relative url</param>
            /// <returns type="Object">Deferred returning folders array</returns>

            var def = $.Deferred();
            if (!ctx)
                ctx = (oLibDocs || folder).get_context();
            if (!oLibDocs)
                oLibDocs = folder.get_listItemAllFields().get_parentList();
            var caml = SP.CamlQuery.createAllItemsQuery();
            caml.set_folderServerRelativeUrl(url || (folder ? folder.get_serverRelativeUrl() : ''));

            // where and order by must be indexed fields
            caml.set_viewXml('<View Scope="RecursiveAll">\
                            <Query>\
                                <OrderBy Override="TRUE">\
                                    <FieldRef Name="FileDirRef" Ascending="True" />\
                                    <FieldRef Name="FileLeafRef" Ascending="True" />\
                                </OrderBy>\
                            </Query>\
                            <RowLimit>1000</RowLimit>\
                        </View>');

            var allDocumentsCol = oLibDocs.getItems(caml);
            ctx.load(allDocumentsCol, "Include(FileSystemObjectType, ServerUrl, ContentTypeId, " + (invFld ? invFld + ", " : '') + (typFld ? typFld + ", " : '') + " " + (!noDt ? "DocumentType, " : "") + "FileLeafRef)");

            ctx.executeQueryAsync(function () {
                var folders = [];
                if (toLocation && (url || folder))
                    folders.push({
                        SourceWeb: ctx.get_url(),
                        DestWeb: toWeb || ctx.get_url(),
                        fileLeafRef: toLocation.substring(toLocation.lastIndexOf('/') + 1),
                        ServerUrl: url || folder.get_serverRelativeUrl(),
                        contentTypeId: folder.get_listItemAllFields().get_fieldValues()['ContentTypeId'],
                        investments: invFld ? folder.get_listItemAllFields().get_fieldValues()[invFld] : null,
                        mandateType: typFld ? folder.get_listItemAllFields().get_fieldValues()[typFld] : null,
                        DocumentType: !noDt ? folder.get_listItemAllFields().get_fieldValues()['DocumentType'] : null,
                        toLocation: (toWeb != null ? toWeb + '/' : '') + toLocation.substring(0, toLocation.lastIndexOf('/'))
                    });

                var listEnumerator = allDocumentsCol.getEnumerator();
                while (listEnumerator.moveNext()) {
                    if (listEnumerator.get_current().get_fileSystemObjectType() != 1)
                        continue;

                    var f = {
                        SourceWeb: ctx.get_url(),
                        DestWeb: toWeb || ctx.get_url(),
                        fileLeafRef: listEnumerator.get_current().get_fieldValues()['FileLeafRef'],
                        ServerUrl: listEnumerator.get_current().get_fieldValues()['ServerUrl'],
                        contentTypeId: listEnumerator.get_current().get_fieldValues()['ContentTypeId'],
                        investments: invFld ? listEnumerator.get_current().get_fieldValues()[invFld] : null,
                        mandateType: typFld ? listEnumerator.get_current().get_fieldValues()[typFld] : null,
                        DocumentType: !noDt ? listEnumerator.get_current().get_fieldValues()['DocumentType'] : null,
                        toLocation: (toWeb != null ? toWeb + '/' : '') + toLocation
                    };

                    f.toLocation += f.ServerUrl.substring((url || (folder ? folder.get_serverRelativeUrl() : '')).length);
                    f.toLocation = f.toLocation.substring(0, f.toLocation.lastIndexOf('/'));

                    folders.push(f);
                }

                return def.resolve(folders);
            },
                function () {
                    return def.reject();
                });

            return def.promise();
        };

        Lists.GetListObject = function (ctx, listNameOrUrl) {
            /// <summary>Gets SPList object</summary>
            /// <param name="ctx" type="Object">Client context, or current context if null</param>
            /// <returns type="Object">Deferred returning SPList, Root folder server relative URL</returns>

            if (ctx == null) ctx = new SP.ClientContext.get_current();
            var def = $.Deferred();
            var list;

            // only name passed in
            if (listNameOrUrl.indexOf('/') < 0) {
                list = ctx.get_web().get_lists().getByTitle(listNameOrUrl);
                ctx.load(list);
                ctx.load(list.get_rootFolder());
                ctx.executeQueryAsync(function () {
                    return def.resolve(list, list.get_rootFolder().get_serverRelativeUrl());
                },
                    function () {
                        return def.reject();
                    });
            }

            // try to calculate list title as its more reliable than url for root addresses
            if (listNameOrUrl.indexOf('/') > -1) {
                var p = listNameOrUrl.substring(ctx.get_url().replace(/\/$/, '').length + 1).split('/')[0];
                list = ctx.get_web().get_lists().getByTitle(p == '09' ? '0-9#' : p);
                ctx.load(list);
                ctx.load(list.get_rootFolder());
                ctx.executeQueryAsync(function () {
                    return def.resolve(list, list.get_rootFolder().get_serverRelativeUrl());
                },
                    function () {
                        // if fail try to get specified url then up through to list
                        var currentFolder = ctx.get_web().getFolderByServerRelativeUrl(listNameOrUrl);
                        ctx.load(currentFolder);
                        var currentFolderItem = currentFolder.get_listItemAllFields();
                        ctx.load(currentFolderItem, 'ServerUrl');
                        list = currentFolderItem.get_parentList();
                        ctx.load(list);
                        ctx.load(list.get_rootFolder());
                        ctx.executeQueryAsync(function () {
                            return def.resolve(list, list.get_rootFolder().get_serverRelativeUrl());
                        },
                            function () {
                                return def.reject();
                            });
                    });
            }

            return def.promise();
        };

        Lists.IntoFolders = function (folders, blankMoveCopy) {
            /// <summary>Moves or copies a set of folders in parralel to an existing structure of folders, generally called by createfolders</summary>
            /// <param name="folders" type="Object">Object of folders to transfer as per GetSubFolders function</param>
            /// <param name="blankMoveCopy" type="Number">0 = Only create folders, 1 = Move folders within folders parameter, 2 = Copy folders in parameter</param>
            /// <returns type="Object">Deferred object</returns>

            var def = $.Deferred();
            var flength = folders.length;
            if (flength === 0 || blankMoveCopy === 0) {
                return def.resolve();
            }

            var counter = 0;
            folders.forEach(function (folder) {
                Lists.GetInFolder(folder.SourceWeb, folder.ServerUrl, 'Files')
                    .done(function (documents) {
                        Lists.MoveCopyDoc(folder.SourceWeb,
                            folder.DestWeb,
                            folder.toLocation.replace(/\/$/, '') + "/" + folder.fileLeafRef,
                            documents.Files.results,
                            blankMoveCopy)
                            .done(function () {
                                counter += 1;
                                if (counter === flength) {
                                    return def.resolve();
                                }
                            })
                            .fail(function () {
                                return def.reject();
                            });
                    })
                    .fail(function () {
                        return def.reject();
                    });
            },
                this);

            return def.promise();
        }

        Lists.MoveCopyDoc = function (src, dst, url, documents, blankMoveCopy) {
            /// <summary>Moves or copies a set of documents in parralel to an existing folder, generally called by intofolders</summary>
            /// <param name="src" type="String">URL of src web</param>
            /// <param name="dst" type="String">URL of dst web</param>
            /// <param name="url" type="String">URL of dst folder</param>
            /// <param name="documents" type="Object">Object of documents to transfer as per GetInFolder function</param>
            /// <param name="blankMoveCopy" type="Number">0 = Only create folders, 1 = Move folders within folders parameter, 2 = Copy folders in parameter</param>
            /// <returns type="Object">Deferred object</returns>

            var def = $.Deferred();
            var dlength = documents.length;
            if (dlength === 0 || blankMoveCopy === 0) {
                return def.resolve();
            }

            var counter = 0;
            documents.forEach(function (doc) {
                if (!doc.File) {
                    doc.File = {}
                }
                doc.File.ServerRelativeUrl = doc.File.ServerRelativeUrl || doc.ServerRelativeUrl || doc.ServerUrl;
                doc.File.Name = doc.File.Name || doc.Name || doc.FileName || doc.Title || doc.FileLeafRef;

                // 1st attempt
                Lists.MoveCopyTo(src, dst, doc.File.ServerRelativeUrl, url.replace(/\/$/, '') + '/' + doc.File.Name, blankMoveCopy)
                    .done(function () {
                        counter += 1;
                        if (counter === dlength) {
                            return def.resolve();
                        }
                    })
                    .fail(function () {
                        console.log('1st fail ' + doc.File.Name);
                        // 2nd attempt
                        Lists.MoveCopyTo(src, dst, doc.File.ServerRelativeUrl, url.replace(/\/$/, '') + '/' + doc.File.Name, blankMoveCopy)
                            .done(function () {
                                counter += 1;
                                if (counter === dlength) {
                                    return def.resolve();
                                }
                            })
                            .fail(function () {
                                console.log('2nd fail ' + doc.File.Name);
                                // 3rd attempt
                                Lists.MoveCopyTo(src, dst, doc.File.ServerRelativeUrl, url.replace(/\/$/, '') + '/' + doc.File.Name, blankMoveCopy)
                                    .done(function () {
                                        counter += 1;
                                        if (counter === dlength) {
                                            return def.resolve();
                                        }
                                    })
                                    .fail(function () {
                                        console.log('3rd fail ' + doc.File.Name);
                                        return def.reject();
                                    });
                            });
                    });
            },
                this);

            return def.promise();
        };

        Lists.MoveCopyTo = function (src, dst, sourceFileUrl, targetFileUrl, blankMoveCopy) {
            var def = $.Deferred();

            var endpointUrl = _spPageContextInfo.webAbsoluteUrl.replace(/\/$/,'') + "/_api/SP.MoveCopyUtil." + 
                (blankMoveCopy == 1 ? "MoveFileByPath()" : "CopyFileByPath()");

            Lists.UpdateDigest()
                .then(function (d) {
                    $.ajax({
                        url: endpointUrl,
                        method: "POST",
                        contentType: "application/json;odata=verbose",
                        headers: {
                            "X-RequestDigest": d
                        },
                        data: JSON.stringify({
                            srcPath: {
                                __metadata: {type: "SP.ResourcePath"},
                                DecodedUrl: document.location.origin + sourceFileUrl
                            },
                            destPath: {
                                __metadata: {type: "SP.ResourcePath"},
                                DecodedUrl: document.location.origin + targetFileUrl
                            },
                            options: {
                                ResetAuthorAndCreatedOnCopy: false,
                                RetainEditorAndModifiedOnMove: true,
                                ShouldBypassSharedLocks: true,
                                __metadata: {type: "SP.MoveCopyOptions"}
                            }
                        }),
                        success: function () {
                            return def.resolve();
                        },
                        error: function () {
                            return def.reject();
                        }
                    });
                });

            return def.promise();
        }

        // fix for IE and data binaries
        Lists.RequestExecutor = function () {
            /// <summary>Loads and overrides request executor to properly support binary files within internet explorer</summary>
        
            var def = $.Deferred();
            if (typeof SP.RequestExecutor != "undefined" && typeof SP.RequestExecutorInternalSharedUtility.BinaryDecode != "undefined") {
                def.resolve();
                return def.promise();
            }

            $.getScript(_spPageContextInfo.webAbsoluteUrl + '/_layouts/15/SP.RequestExecutor.js',
                function () {
                    SP.RequestExecutorInternalSharedUtility.BinaryDecode = function 
                        SP_RequestExecutorInternalSharedUtility$BinaryDecode(data) {
                        var ret = '';

                        if (data) {
                            var byteArray = new Uint8Array(data);

                            for (var i = 0; i < data.byteLength; i++) {
                                ret = ret + String.fromCharCode(byteArray[i]);
                            }
                        };
                        return ret;
                    };
    
                    SP.RequestExecutorUtility.IsDefined = function SP_RequestExecutorUtility$$1(data) {
                        var nullValue = null;
    
                        return data === nullValue || typeof data === 'undefined' || !data.length;
                    };
    
                    SP.RequestExecutor.ParseHeaders = function SP_RequestExecutor$ParseHeaders(headers) {
                        if (SP.RequestExecutorUtility.IsDefined(headers)) {
                            return null;
                        }
                        var result = {};
                        var reSplit = new RegExp('\r?\n');
                        var headerArray = headers.split(reSplit);
    
                        for (var i = 0; i < headerArray.length; i++) {
                            var currentHeader = headerArray[i];
    
                            if (!SP.RequestExecutorUtility.IsDefined(currentHeader)) {
                                var splitPos = currentHeader.indexOf(':');
    
                                if (splitPos > 0) {
                                    var key = currentHeader.substr(0, splitPos);
                                    var value = currentHeader.substr(splitPos + 1);
    
                                    key = SP.RequestExecutorNative.trim(key);
                                    value = SP.RequestExecutorNative.trim(value);
                                    result[key.toUpperCase()] = value;
                                }
                            }
                        }
                        return result;
                    };
    
                    SP.RequestExecutor.internalProcessXMLHttpRequestOnreadystatechange = function 
                        SP_RequestExecutor$internalProcessXMLHttpRequestOnreadystatechange(xhr, requestInfo,timeoutId) {
                        if (xhr.readyState === 4) {
                            if (timeoutId) {
                                window.clearTimeout(timeoutId);
                            }
                            xhr.onreadystatechange = SP.RequestExecutorNative.emptyCallback;
                            var responseInfo = new SP.ResponseInfo();

                            responseInfo.state = requestInfo.state;
                            responseInfo.responseAvailable = true;
                            if (requestInfo.binaryStringResponseBody) {
                                responseInfo.body = SP.RequestExecutorInternalSharedUtility
                                    .BinaryDecode(xhr.response);
                            } else {
                                responseInfo.body = xhr.responseText;
                            }
                            responseInfo.statusCode = xhr.status;
                            responseInfo.statusText = xhr.statusText;
                            responseInfo.contentType = xhr.getResponseHeader('content-type');
                            responseInfo.allResponseHeaders = xhr.getAllResponseHeaders();
                            responseInfo.headers = SP.RequestExecutor
                                .ParseHeaders(responseInfo.allResponseHeaders);
                            if (xhr.status >= 200 && xhr.status < 300 || xhr.status === 1223) {
                                if (requestInfo.success) {
                                    requestInfo.success(responseInfo);
                                }
                            } else {
                                var error = SP.RequestExecutorErrors.httpError;
                                var statusText = xhr.statusText;

                                if (requestInfo.error) {
                                    requestInfo.error(responseInfo, error, statusText);
                                }
                            }
                        }
                    };

                    return def.resolve();
                });

            return def.promise();
        }

        Lists.Fix = function (name) {
            return encodeURIComponent(name) // encode url for browser
                    .replace(/'/g,'%2527') // double encode ' as they are typically within ('...')
                    .replace(/%3A/g,':') // revert : as only ever exists in https://
                    .replace(/%2F/g,'/'); // revert / as its only ever folder dividers
        }

        Lists.GetFile = function (src, name) {
            var def = $.Deferred();

            Lists.UrlContent(Lists.Fix(name).replace(/%2527/g, '%27')).then(function (exists, response) {
                if (exists)
                    return def.resolve({body: response});
                return def.reject();
            },function () {
                return def.reject();
            })

            return def.promise();
        };

        // del with ' in file name will del as %27
        Lists.DelListAttachments = function (listName, id, name, isSite) {
            /// <summary>Delete list item attachments to an items</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="id" type="String">List item Id</param>
            /// <param name="name" type="String">File name</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred object returning rest formatted as {return}.d.results</returns>

            var def = $.Deferred();
        
            Lists.UpdateDigest()
                .then(function (d) {
                    $.ajax({
                        url: (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") + "/_api/web/lists/getbytitle('" + listName + "')/items(" + id + ")/AttachmentFiles/getByFileName('" +  Lists.Fix(name) + "')",
                        method: "DELETE",
                        headers: {
                            "ACCEPT": "application/json;odata=verbose",
                            "X-RequestDigest": d
                        },
                        contentType: "application/json;odata=verbose",
                        success: function (d) {
                            return def.resolve(d);
                        },
                        fail: function (a,b) {
                            return def.reject(a,b);
                        }
                    });
                });
                
            return def.promise();
        };

        // add with ' in file name will save as %27, # will break _vti_bin
        Lists.AddListAttachments = function (listName, id, name, data, isSite) {
            /// <summary>Get list item attachments to an items</summary>
            /// <param name="listName" type="String">The list title</param>
            /// <param name="id" type="String">List item Id</param>
            /// <param name="name" type="String">File name</param>
            /// <param name="data" type="String">File data</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred object returning rest formatted as {return}.d.results</returns>

            var def = $.Deferred();
            
            Lists.UpdateDigest()
                .then(function (d) {
                    Lists.RequestExecutor().then(function () {
                        var dest = new SP.RequestExecutor(typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl);
                        var info = {
                            url: "_api/web/lists/GetByTitle('" + listName + "')/items(" + id + ")/AttachmentFiles/add(FileName='" + Lists.Fix(name) + "')",
                            method: "POST",
                            headers: {
                                "Accept": "application/json; odata=verbose",
                                "X-RequestDigest": d
                            },
                            contentType: "application/json;odata=verbose",
                            binaryStringRequestBody: true,
                            body: data,
                            success: function (d) {
                                return def.resolve(d.body ? JSON.parse(d.body) : d);
                            },
                            fail: function (a,b) {
                                return def.reject(a,b);
                            }
                        }
                        dest.executeAsync(info);
                    });
                });
                
			return def.promise();
        };
		
        // del with % ' # in file name will error
        Lists.DelFile = function (isSite, name) {
            /// <summary>Deleted a file</summary>
            /// <param name="isSite" type="String">URL of dst web</param>
            /// <param name="name" type="String">URL of dst file</param>
            /// <returns type="Object">Deferred object</returns>

            var def = $.Deferred();

            Lists.UpdateDigest()
                .then(function (d) {
                    $.ajax({
                        url: (typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl).replace(/\/$/, "") + "/_api/web/GetFileByServerRelativeUrl('" + Lists.Fix(name) + "')",
                        method: "DELETE",
                        headers: {
                            "Accept": "application/json; odata=verbose",
                            "X-RequestDigest": d
                        },
                        contentType: "application/json;odata=verbose",
                        success: function (d) {
                            return def.resolve(d);
                        },
                        error: function () {
                            return def.reject();
                        }
                    });
                });

            return def.promise();
        };

        // add with % in file name will error
        Lists.AddFile = function (isSite, targetFileUrl, data, plainText) {
            /// <summary>Adds a file to a location</summary>
            /// <param name="dst" type="String">URL of dst web</param>
            /// <param name="targetFileUrl" type="String">URL of dst file</param>
            /// <param name="data" type="Object">Binary data</param>
            /// <returns type="Object">Deferred object</returns>

            var def = $.Deferred();

            Lists.UpdateDigest()
                .then(function (d) {
                    Lists.RequestExecutor().then(function () {
                        var dest = new SP.RequestExecutor(typeof (isSite) == "string" ? isSite : isSite ? _spPageContextInfo.siteServerRelativeUrl : _spPageContextInfo.webServerRelativeUrl);
                        var info = {
                            url: "_api/web/GetFolderByServerRelativeUrl('" + Lists.Fix(targetFileUrl.substring(0, targetFileUrl.lastIndexOf('/'))) + "')/Files/Add(url='" + Lists.Fix(targetFileUrl.substring(targetFileUrl.lastIndexOf('/') + 1)) + "', overwrite=true)",
                            method: "POST",
                            headers: {
                                "Accept": "application/json; odata=verbose",
                                "X-RequestDigest": d
                            },
                            contentType: "application/json;odata=verbose",
                            binaryStringRequestBody: plainText != true,
                            body: data,
                            success: function (d) {
                                //return def.resolve(d);
                                var b = d.body ? JSON.parse(d.body) : d;
                                $.ajax({
                                    url: b.d.ListItemAllFields.__deferred.uri,
                                    type: "GET",
                                    headers: {
                                        "accept": "application/json;odata=verbose"
                                    },
                                    success: function (d) {
                                        return def.resolve(d);
                                    },
                                    error: function () {
                                        return def.resolve(null);
                                    }
                                });

                            },
                            error: function () {
                                return def.reject();
                            }
                        };
                        dest.executeAsync(info);
                    });
                });

            return def.promise();
        };

        Lists.LookupFix = function (src, dst, fld) {
            var def = $.Deferred();

            var context = SP.ClientContext.get_current();
            var web = context.get_web();
            
            var source = web.get_lists().getByTitle(src);
            var dest = web.get_lists().getByTitle(dst);
            
            context.load(source, "Id");
            
            var field = dest.get_fields().getByInternalNameOrTitle(fld);
            
            context.load(field, "SchemaXml");
            
            context.executeQueryAsync(function () {
                field.set_schemaXml(field.get_schemaXml().replace(/List[^ ]*/, 'List="{' + source.get_id().toString().toUpperCase() + '}"'));
                field.update();
                
                setTimeout(function () {
                    context.executeQueryAsync(function () {
                        return def.resolve();
                    }, function () {
                        return def.reject();
                    });
                },100);
            }, function () {
                return def.reject();
            }); 

            return def.promise();           
        }
    }
}
