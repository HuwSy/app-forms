'use strict';
// start loading immediately
if (typeof SP != "undefined")
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', null);
    
// sp pnp js
// https://pnp.github.io/pnpjs/sp/

// all web operations
{
    Webs = Webs || {};
    {
		Webs.Email = function (emailBody, subject, to, cc, bc, from) {
			if (typeof(to) == 'string')
				to = to.split(/[,;]/);
			if (cc == null)
				cc = [''];
			if (typeof(cc) == 'string')
				cc = cc.split(/[,;]/);
			if (bc == null)
				bc = [''];
			if (typeof(bc) == 'string')
				bc = bc.split(/[,;]/);
			
			var def = $.Deferred();
			Lists.UpdateDigest()
				.then(function (d) {
					$.ajax({
						contentType: 'application/json',
						url: _spPageContextInfo.webAbsoluteUrl + "/_api/SP.Utilities.Utility.SendEmail",
						type: "POST",
						data: JSON.stringify({
							'properties': {
								'__metadata': { 'type': 'SP.Utilities.EmailProperties' },
								'From': (from || _spPageContextInfo.userEmail),
								'To': { 'results': to },
								'CC': { 'results': cc },
								'BCC': { 'results': bc },
								'Body': emailBody,
								'Subject': subject
							}
						}),
						headers: {
							"Accept": "application/json;odata=verbose",
							"content-type": "application/json;odata=verbose",
							"X-RequestDigest": d
						},
						success: function () {
							return def.resolve();
						},
						error: function () {
							return def.reject();
						}
					});
                },function () {
                    return def.reject();
                });
			
            return def.promise();
		}

        Webs.IsEditing = function() {
            /// <summary>Is the page in edit mode</summary>
            try {
                if (document.forms[MSOWebPartPageFormName].MSOLayout_InDesignMode.value == "1")
                    return true;
            } catch (e) {
            }
            try {
                if (document.forms[MSOWebPartPageFormName]._wikiPageMode.value == "Edit")
                    return true;
            } catch (e) {
            }

            return false;
        }

        var LoopProperties = function (up, key) {
            if (up != null && up.results != null && up.results.length > 0) {
                var prop = up.results;
                for (var i = 0; i < prop.length; i++) {
                    if (prop[i].Key === key) {
                        return prop[i].Value === null ? "" : prop[i].Value;
                    }
                }
            }
            return "";
        }

        Webs.GetFlatSearch = function (data, useHref, isWeb, docKeywords) {
            var nodes = [];
            var results = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results;
            for (var i = 0; i < results.length; i++) {
                try {
                    var result = results[i].Cells;
                    var p = LoopProperties(result, "Path").replace(/\/$/, "");

                    var item = {
                        desc: LoopProperties(result, "Description"),
                        modified: LoopProperties(result, "LastModifiedTime"),
                        site: LoopProperties(result, "SiteTitle"),
                        split: "",
                        search: 1,
                        path: p.toLowerCase(),
                        web: isWeb
                    };

                    item.title = LoopProperties(result, "Title").replace(LoopProperties(result, "SiteTitle") + " - ", "");
                    // if its policy template and not been renamed by user use file name
                    if (item.title == "Policy Template") {
                        item.title = p.substring(p.lastIndexOf('/') + 1);
                        if (item.title.lastIndexOf(".") > 0)
                            item.title = item.title.substring(0, item.title.lastIndexOf("."));
                    }
                    // if its a web
                    if (item.title == "") {
                        item.title = p.substring(p.lastIndexOf('/') + 1);
                        item.path = p.substring(0, p.lastIndexOf('/')).toLowerCase();
                    }

                    if (docKeywords)
                        item.keywords = LoopProperties(result, docKeywords);

                    if (item.path.indexOf('/forms/') > 0) {
                        item.path = item.path.substring(0, item.path.indexOf('/forms/'));
                    }

                    item.split = item.path.substring(_spPageContextInfo.webAbsoluteUrl.length + 1);
                    item.parent = item.split.substring(0, item.split.lastIndexOf('/'));

                    if (!useHref) {
                        item.path = null;
                    }
					
                    nodes.push(item);
                } catch (e) {
                }
            }

            return nodes;
        }

        Webs.GetFlatData = function (data, useHref, isWeb, parentUrl) {
            var nodes = [];
            var results = data.d.results != null ? data.d.results : data.d;
            for (var i = 0; i < results.length; i++) {
                try {
                    var result = results[i];
                    if ((result.NoCrawl == null || result.NoCrawl == false) &&
                        (parentUrl || result.Url != null || result.DefaultView.ServerRelativeUrl != null)) {

                        var item = {
                            title: result.Title || result.name,
                            desc: result.Description || result.name,
                            modified: result.LastItemModifiedDate || result.term.get_lastModifiedDate(),
                            site: null,
                            split: "",
                            search: 0,
                            path: parentUrl
                                ? (parentUrl + Terms.CleanToUrl(result.term).replace(/\/$/, ""))
                                : result.Url != null
                                ? result.Url.toLowerCase().replace(/\/$/, "")
                                : (window.location.origin.replace(/\/$/, "") + result.DefaultView.ServerRelativeUrl).toLowerCase().replace(/\/$/, ""),
                            web: isWeb,
                            code: parentUrl != null
                        };

                        if (item.path.indexOf('/forms/') > 0) {
                            item.path = item.path.substring(0, item.path.indexOf('/forms/'));
                        }

                        item.split = item.path.substring((parentUrl ? parentUrl : _spPageContextInfo.webAbsoluteUrl).replace(/\/$/, "").length + 1).replace(/\/$/, '');
                        item.parent = item.split.substring(0, item.split.lastIndexOf('/'));

                        if (!useHref) {
                            item.path = null;
                        }

                        nodes.push(item);

                        var child = null;
                        if (result.children && result.children.length > 0)
                            child = Webs.GetFlatData({ d: result.children}, true, false, item.path + '/');
                        if (child)
                            nodes = nodes.concat(child);
                    }
                } catch (e) {}
            }

            return nodes;
        }

        Webs.GetUser = function() {
            /// <summary>Get current user</summary>
            var def = $.Deferred();

            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") +
                    '/_api/SP.UserProfiles.PeopleManager/GetMyProperties',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(data);
                },
                error: function() {
                    return def.resolve();
                }
            });

            return def.promise();
        }

        Webs.EnsureUser = function(u) {
            /// <summary>Gets specified user</summary>
            var def = $.Deferred();

            var ctx = new SP.ClientContext.get_current();
            var user = ctx.get_web().ensureUser(u);
            ctx.load(user);
            ctx.executeQueryAsync(function () {
                    return def.resolve(user);
                }, function () {
                    return def.reject();
                }
            );

            return def.promise();
        }

        Webs.GetWebPermMasks = function (p, web) {
            var def = $.Deferred();

            var check = function (data) {
                if (data.d)
                    window.EffectiveBasePermissions = JSON.stringify(data.d.EffectiveBasePermissions);
                var perm = new SP.BasePermissions();
                perm.initPropertiesFromJson(JSON.parse(window.EffectiveBasePermissions));
                return def.resolve(perm.has(p));
            }

            if (window.EffectiveBasePermissions) {
                check(window.EffectiveBasePermissions);
                return def.promise();
            }

            $.ajax({
                url: (web || _spPageContextInfo.webAbsoluteUrl).replace(/\/$/, "") +
                    '/_api/web/EffectiveBasePermissions',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: check,
                error: function() {
                    return def.resolve(false);
                }
            });

            return def.promise();
        }

        Webs.GetWebs = function(web) {
            /// <summary>Get webs under current web</summary>
            var def = $.Deferred();

            Webs.GetWebPermMasks(SP.PermissionKind.useRemoteAPIs, web).then(function (p) {
                if (!p)
                    return def.resolve();

                $.ajax({
                    url: (web || _spPageContextInfo.webAbsoluteUrl).replace(/\/$/, "") +
                        '/_api/web/webs/?$filter=effectivebasepermissions/high%20gt%2048',
                    type: "GET",
                    headers: { "accept": "application/json;odata=verbose" },
                    success: function(data) {
                        return def.resolve(Webs.GetFlatData(data, true, true));
                    },
                    error: function() {
                        return def.resolve();
                    }
                });
            });

            return def.promise();
        }

        Webs.GetWebsSearch = function(web) {
            /// <summary>Get webs under current web via search</summary>
            var def = $.Deferred();

            $.ajax({
                // run search against root site collection as read only users dont have access to this api on current web
                url: window.location.origin +
                    '/_api/search/query?querytext=%27Path:' +
                    (web || _spPageContextInfo.webAbsoluteUrl) +
                    ' -Path=' +
                    (web || _spPageContextInfo.webAbsoluteUrl) +
                    ' ContentClass=STS_Web%27&selectproperties=%27Title,Description,LastModifiedTime,Path,SiteTitle%27&rowlimit=500&trimduplicates=false&sortlist=%27LastModifiedTime:descending%27',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(Webs.GetFlatSearch(data, false, true));
                },
                error: function() {
                    return def.resolve();
                }
            });

            return def.promise();
        }

        Webs.GetLibs = function(web) {
            /// <summary>Get libs under current web</summary>
            var def = $.Deferred();

            Webs.GetWebPermMasks(SP.PermissionKind.useRemoteAPIs, web).then(function (p) {
                if (!p)
                    return def.resolve();

                $.ajax({
                    url: (web || _spPageContextInfo.webAbsoluteUrl).replace(/\/$/, "") +
                        '/_api/web/lists?$expand=DefaultView&$filter=Title ne %27MicroFeed%27 and NoCrawl eq false and Hidden eq false',
                    type: "GET",
                    headers: { "accept": "application/json;odata=verbose" },
                    success: function(data) {
                        return def.resolve(Webs.GetFlatData(data, true, false));
                    },
                    error: function() {
                        return def.resolve();
                    }
                });
        
            });

            return def.promise();
        }

        Webs.GetLibsSearch = function(web) {
            /// <summary>Get libs under current web by search</summary>
            var def = $.Deferred();

            $.ajax({
                // run search against root site collection as read only users dont have access to this api on current web
                url: window.location.origin +
                    '/_api/search/query?querytext=%27Path:' +
                    (web || _spPageContextInfo.webAbsoluteUrl) +
                    ' ContentClass=STS_List_DocumentLibrary%27&selectproperties=%27Title,Description,LastModifiedTime,Path,SiteTitle%27&rowlimit=500&trimduplicates=false&sortlist=%27LastModifiedTime:descending%27',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(Webs.GetFlatSearch(data, true, false));
                },
                error: function() {
                    return def.resolve();
                }
            });

            return def.promise();
        }

        Webs.GetFilesSearch = function(web) {
            /// <summary>Get libs under current web by searching for files within that the user has access to</summary>
            var def = $.Deferred();

            $.ajax({
                // run search against root site collection as read only users dont have access to this api on current web
                url: window.location.origin +
                    '/_api/search/query?querytext=%27Path:' +
                    (web || _spPageContextInfo.webAbsoluteUrl) +
                    ' ContentClass=STS_ListItem_DocumentLibrary%27&selectproperties=%27LastModifiedTime,Path,SiteTitle%27&rowlimit=500&trimduplicates=false&sortlist=%27LastModifiedTime:descending%27',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(Webs.GetFlatSearch(data, true, false));
                },
                error: function() {
                    return def.resolve();
                }
            });

            return def.promise();
        }

        Webs.GetFilesUnderPath = function(docKeywords, web) {
            /// <summary>Get files beneath web</summary>
            var def = $.Deferred();

            $.ajax({
                // run search against root site collection as read only users dont have access to this api on current web
                url: window.location.origin +
                    '/_api/search/query?querytext=%27Path:' +
                    (web || _spPageContextInfo.webAbsoluteUrl) +
                    ' ContentClass:STS_ListItem_%27&selectproperties=%27Title,Description,LastModifiedTime,Path,SiteTitle' +
                    (docKeywords ? ',' + docKeywords : '') +
                    '%27&rowlimit=500&trimduplicates=false&sortlist=%27LastModifiedTime:descending%27',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(Webs.GetFlatSearch(data, true, false, docKeywords));
                },
                error: function () {
                    return def.resolve();
                }
            });

            return def.promise();
        }

        Webs.CurrentWeb = function () {
            var def = $.Deferred();

            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "") + '/_api/web',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(data.d);
                },
                error: function() {
                    return def.reject();
                }
            });

            return def.promise();
        }

        Webs.CreateWeb = function (address, template, unique, retries) {
            var def = $.Deferred();
			
            var title, url;
            if (typeof address == "object" && address.length == 2) {
                title = address[1];
                url = address[0];
            } else {
                title = address;
                url = address.toLowerCase().replace(/\W+/g, "-").replace(/--/g, "-");
            }

            if (title.length <= 0 || url.length <= 0)
                return def.reject();

            SP.SOD.registerSod('SP.Publishing.js', '\u002f_layouts\u002f15\u002fSP.Publishing.js');
            SP.SOD.executeFunc('sp.publishing.js',
                'SP.Publishing.Navigation',
                function () {
                    var clientContext = new SP.ClientContext.get_current();
                    var web = clientContext.get_web();
                    clientContext.load(web);

                    // create web
                    var webCreationInformation = new SP.WebCreationInformation();
                    webCreationInformation.set_title(title);
                    webCreationInformation.set_description(title);
                    webCreationInformation.set_language(1033);
                    webCreationInformation.set_url(url);
                    webCreationInformation.set_webTemplate(template);
                    webCreationInformation.set_useSamePermissionsAsParentSite(!unique);

                    var nweb = web.get_webs().add(webCreationInformation);
                    clientContext.load(nweb);
                    clientContext.executeQueryAsync(function () {
                            return def.resolve(nweb, web);
                        },
                        function () {
                            var nweb = web.get_webs().add(webCreationInformation);
                            clientContext.load(nweb);
                            clientContext.executeQueryAsync(function () {
                                    return def.resolve(nweb, web);
                                },
                                function () {
                                    var nweb = web.get_webs().add(webCreationInformation);
                                    clientContext.load(nweb);
                                    clientContext.executeQueryAsync(function () {
                                            return def.resolve(nweb, web);
                                        },
                                        function () {
                                            var nweb = web.get_webs().add(webCreationInformation);
                                            clientContext.load(nweb);
                                            clientContext.executeQueryAsync(function () {
                                                    return def.resolve(nweb, web);
                                                },
                                                function () {
                                                    var nweb = web.get_webs().add(webCreationInformation);
                                                    clientContext.load(nweb);
                                                    clientContext.executeQueryAsync(function () {
                                                            return def.resolve(nweb, web);
                                                        },
                                                        function () {
                                                            return def.reject();
                                                        });
                                                });
                                        });
                                });
                        });
                });

            return def.promise();
        }

        Webs.SetupWeb = function (url, master, custom, features, inclibs) {
            var def = $.Deferred();

            if (!url)
                url = _spPageContextInfo.webServerRelativeUrl;
            url = url.replace(/\/$/, '');

            SP.SOD.registerSod('sp.publishing.js', '\u002f_layouts\u002f15\u002fSP.Publishing.js');
            SP.SOD.loadMultiple(['sp.publishing.js'], function () {
                Lists.GetListItem(null, null, 'Title', url)
                    .then(function (d) {
                        var allLists = d.d.results.map(function (l) { return l.Title });

                        var nweb = new SP.ClientContext(url).get_web();
                        var cont = nweb.get_context();

                        // set master pages
                        {
                            if (master)
                                nweb.set_masterUrl(master);
                            if (custom)
                                nweb.set_customMasterUrl(custom);
                            nweb.update();
                        }

                        // set navigation to structured current and siblings
                        {
                            var webNavSettings = new SP.Publishing.Navigation.WebNavigationSettings(cont, nweb);
                            var navigation = webNavSettings.get_currentNavigation();
                            navigation.set_source(1);
                            webNavSettings.update();

                            var prop = nweb.get_allProperties();
                            prop.set_item('__NavigationShowSiblings', "False");
                            prop.set_item('__InheritCurrentNavigation', "False");
                            prop.set_item('__IncludeSubSitesInNavigation', "False");
                            nweb.update();
                        }

                        // set always required features
                        {
                            // publishing, required
                            nweb.get_features().add(new SP.Guid('{94c94ca6-b32f-4da9-a9e3-1f3d343d7ecb}'), true, SP.FeatureDefinitionScope.farm);
                            // nintex workflow
                            if (!_spPageContextInfo.isSPO)
                                nweb.get_features().add(new SP.Guid('{9bf7bf98-5660-498a-9399-bc656a61ed5d}'), true, SP.FeatureDefinitionScope.farm);
                            // home page, required
                            nweb.get_features().add(new SP.Guid('{27ac1171-70a3-4603-8862-65f849174038}'), true, SP.FeatureDefinitionScope.site);
                            // modern sites
                            //https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages?tabs=pnpposh
                            //add web part manually
                        }

                        // no crawl
                        {
                            var docs = null;
                            if (~allLists.indexOf('Documents')) {
                                docs = nweb.get_lists().getByTitle('Documents');
                                cont.load(docs.get_rootFolder());
                            }

                            var shared = null;
                            if (~allLists.indexOf('Shared Documents')) {
                                shared = nweb.get_lists().getByTitle('Shared Documents');
                            }

                            var emails = null;
                            if (~allLists.indexOf('Emails')) {
                                emails = nweb.get_lists().getByTitle('Emails');
                            }

                            var lists = [];
                            if (~allLists.indexOf('Pages')) {
                                lists.push(nweb.get_lists().getByTitle('Pages'));
                                lists[lists.length - 1].set_noCrawl(false);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Site Pages')) {
                                lists.push(nweb.get_lists().getByTitle('Site Pages'));
                                lists[lists.length - 1].set_noCrawl(false);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Site Assets')) {
                                lists.push(nweb.get_lists().getByTitle('Site Assets'));
                                lists[lists.length - 1].set_noCrawl(true);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Images')) {
                                lists.push(nweb.get_lists().getByTitle('Images'));
                                lists[lists.length - 1].set_noCrawl(true);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Form Templates')) {
                                lists.push(nweb.get_lists().getByTitle('Form Templates'));
                                lists[lists.length - 1].set_noCrawl(true);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Site Collection Images')) {
                                lists.push(nweb.get_lists().getByTitle('Site Collection Images'));
                                lists[lists.length - 1].set_noCrawl(true);
                                lists[lists.length - 1].update();
                            }
                            if (~allLists.indexOf('Site Collection Documents')) {
                                lists.push(nweb.get_lists().getByTitle('Site Collection Documents'));
                                lists[lists.length - 1].set_noCrawl(true);
                                lists[lists.length - 1].update();
                            }
                        }

                        // execute
                        cont.executeQueryAsync(function () {
                            // if documents lib is shared docs path then rename, MS bug fixing
                            if (docs && docs.get_rootFolder().get_name() == "Shared Documents") {
                                docs.set_title("Shared Documents");
                                docs.update();
                                shared = docs;
                                docs = null;
                            }

                            cont.executeQueryAsync(function () {
                                // make libs
                                var dci = null;
                                if (!docs) {
                                    dci = new SP.ListCreationInformation();
                                    dci.set_title("Documents");
                                    dci.set_templateType(SP.ListTemplateType.documentLibrary);
                                    docs = nweb.get_lists().add(dci);
                                }

                                var sci = null;
                                if (!shared) {
                                    sci = new SP.ListCreationInformation();
                                    sci.set_title("Shared Documents");
                                    sci.set_templateType(SP.ListTemplateType.documentLibrary);
                                    shared = nweb.get_lists().add(sci);
                                }

                                var eci = null, to = null;
                                if (!emails) {
                                    eci = new SP.ListCreationInformation();
                                    eci.set_title("Emails");
                                    eci.set_templateType(SP.ListTemplateType.documentLibrary);
                                    emails = nweb.get_lists().add(eci);
                                } else {
                                    to = emails.get_fields();
                                    cont.load(to);
                                }

                                // addit features, may add cts
                                if (features) {
                                    for (var f in features) {
                                        nweb.get_features().add(features[f][0], features[f][1], features[f][2]);
                                    }
                                }

                                cont.executeQueryAsync(function () {
                                    // ensure crawl, versioning etc
                                    {
                                        docs.set_noCrawl(!inclibs);
                                        docs.set_enableVersioning(true);
                                        docs.set_enableMinorVersions(true);
                                        docs.set_forceCheckout(true);
                                        docs.set_onQuickLaunch(!!inclibs);
                                        docs.update();

                                        shared.set_noCrawl(!inclibs);
                                        shared.set_enableVersioning(true);
                                        shared.set_enableMinorVersions(true);
                                        shared.set_forceCheckout(true);
                                        shared.set_onQuickLaunch(!!inclibs);
                                        shared.update();

                                        emails.set_noCrawl(!inclibs);
                                        emails.set_enableVersioning(false);
                                        emails.set_enableMinorVersions(false);
                                        emails.set_forceCheckout(false);
                                        emails.set_onQuickLaunch(!!inclibs);
                                        if (eci || to.get_count() < 128) {
                                            emails.get_fields().addFieldAsXml('<Field ID="{54bda702-0f88-486c-9dd9-0584fa9ad520}" Type="Note" DisplayName="Cc" Description="The identity of the secondary recipients of the message." Required="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" NumLines="6" RichText="FALSE" Sortable="FALSE" StaticName="MailCc" Name="MailCc" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{7ca5d3cd-16b4-412e-ad55-2345bb906cfc}" Type="DateTime" DisplayName="Date" Description="The date and time when the message was sent." Required="FALSE" Format="DateTime" StaticName="MailDate" Name="MailDate" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{c8e8ffa2-caa6-4065-8e64-9830cafa729f}" Type="Text" DisplayName="From" Description="The identity of the person who sent the message." Required="FALSE" MaxLength="255" StaticName="MailFrom" Name="MailFrom" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{c86d98a8-3c2b-4be6-a91b-a1f659ab2a74}" Type="Boolean" DisplayName="Attachments" Description="Indicates if the e-mail message contains one or more attachments." Required="FALSE" MaxLength="255" StaticName="MailAttachments" Name="MailAttachments" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}"><Default>0</Default></Field>');
                                            emails.get_fields().addFieldAsXml('<Field ID="{d912b86a-5a1e-48d2-8aba-8a7abd731f16}" Type="Text" DisplayName="In-Reply-To" Description="The contents of this field identify previous correspondence that this message answers." Required="FALSE" MaxLength="255" StaticName="MailIn-Reply-To" Name="MailIn_x002d_Reply_x002d_To" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{5840efa9-d228-4d83-9dd7-790e160a2a70}" Type="Text" DisplayName="OriginalSubject" Description="A summary of the message." Required="FALSE" MaxLength="255" StaticName="MailOriginalSubject" Name="MailOriginalSubject" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{bb8eccf1-cfcf-4bdc-9cd4-747008eb0156}" Type="Text" DisplayName="References" Description="The contents of this field identify other correspondence that this message answers." Required="FALSE" MaxLength="255" StaticName="MailReferences" Name="MailReferences" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{0bcfa246-cd44-4a1b-a798-f5240bc39fb0}" Type="Text" DisplayName="Reply-To" Description="Indicates any mailbox(es) to which responses are to be sent." Required="FALSE" MaxLength="255" StaticName="MailReply-To" Name="MailReply_x002d_To" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{c9ac09c7-685d-4b60-a95a-4f8630b134e7}" Type="Text" DisplayName="Subject" Description="A summary of the message." Required="FALSE" MaxLength="255" StaticName="MailSubject" Name="MailSubject" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().addFieldAsXml('<Field ID="{e4306523-3254-4e5a-96e3-3470a7a4d0b2}" Type="Note" DisplayName="To" Description="The identity of the primary recipients of the message." Required="FALSE" UnlimitedLengthInDocumentLibrary="FALSE" NumLines="6" RichText="FALSE" Sortable="FALSE" StaticName="MailTo" Name="MailTo" Group="E-mail Columns" Customization="" SourceID="{e396cdcc-f931-4978-844e-80207b53bfb4}" />');
                                            emails.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms'));
                                            emails.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms_TaxHTField'));
                                        }
                                        emails.update();
                                    }

                                    // load cts
                                    var cts = cont.get_site().get_rootWeb().get_contentTypes();
                                    cont.load(cts);

                                    // get nodes
                                    var ql = nweb.get_navigation().get_quickLaunch();
                                    cont.load(ql);
                                    cont.load(nweb, 'Title','ServerRelativeUrl','HasUniqueRoleAssignments');

                                    cont.executeQueryAsync(function () {
                                        Lists.SetView(nweb.get_serverRelativeUrl(), 'Emails','DocIcon,MailAttachments,LinkFilename,MailFrom,MailTo,MailCc,MailDate,Modified,Editor');

                                        // add cts to doc libs
                                        {
                                            var ctEnum = cts.getEnumerator();
                                            var expected = [
                                                "Document",
                                                "Excel",
                                                "PowerPoint",
                                                "Letter",
                                                "BoardPaper",
                                                "Meeting",
                                                "Minutes"
                                            ];

                                            var hasCt = false;
                                            while (ctEnum.moveNext()) {
                                                var ct = ctEnum.get_current();
                                                if (~expected.indexOf(ct.get_name())) {
                                                    hasCt = true;

                                                    docs.set_contentTypesEnabled(true);
                                                    shared.set_contentTypesEnabled(true);
                                                    docs.get_contentTypes().addExistingContentType(ct);
                                                    shared.get_contentTypes().addExistingContentType(ct);
                                                }
                                            }

                                            docs.update();
                                            shared.update();
                                        }

                                        // loop navigation to see if home and back are needed and remove some common unwanted
                                        {
                                            var linkEnum = ql.getEnumerator();
                                            var needsHome = true, needsBack = url != _spPageContextInfo.siteServerRelativeUrl.replace(/\/$/, ''), nnci, deletes = [], i = 0;

                                            while (linkEnum.moveNext()) {
                                                var res = linkEnum.get_current();
                                                if (res.get_title() == nweb.get_title())
                                                    needsHome = false;
                                                if (res.get_title() == 'Back')
                                                    needsBack = false;
                                                if (res.get_title() == 'Site Contents'
                                                    || res.get_title() == 'Libraries'
                                                    || res.get_title() == 'Lists'
                                                    || res.get_title() == 'Recent'
                                                    || res.get_title() == 'Home')
                                                    deletes.push(i);
                                                i++;
                                            }

                                            deletes.reverse().forEach(function (d) {
                                                ql.getItemAtIndex(d).deleteObject();
                                            });

                                            if (needsHome) {
                                                nnci = new SP.NavigationNodeCreationInformation();
                                                nnci.set_title(nweb.get_title());
                                                nnci.set_url(url);
                                                ql.add(nnci);
                                            }

                                            if (needsBack) {
                                                nnci = new SP.NavigationNodeCreationInformation();
                                                nnci.set_title('Back');
                                                nnci.set_url(url.substring(0,url.lastIndexOf('/')));
                                                ql.add(nnci);
                                            }
                                        }

                                        var full = nweb.get_roleDefinitions().getByName('Full Control');
                                        cont.load(full, 'Id');
                                        cont.executeQueryAsync(function () {
                                            Lists.GetListItem('Pages', null, null, nweb.get_serverRelativeUrl(), 'RoleAssignments,File')
                                                .then(function (p) {
                                                    // attempt fields add
                                                    docs.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('KeyDocument'));
                                                    docs.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('DocumentAuthor'));
                                                    docs.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('DocumentType'));
                                                    docs.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms'));
                                                    docs.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms_TaxHTField'));
    
                                                    shared.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('KeyDocument'));
                                                    shared.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('DocumentAuthor'));
                                                    shared.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('DocumentType'));
                                                    shared.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms'));
                                                    shared.get_fields().add(cont.get_site().get_rootWeb().get_fields().getByInternalNameOrTitle('SiteTerms_TaxHTField'));
                                                
                                                    docs.update();
                                                    shared.update();

                                                    cont.executeQueryAsync(testMakeFields, testMakeFields);
                                                    
                                                    var restricted;
                                                    function testMakeFields () {
                                                        // get roles
                                                        restricted = nweb.get_roleDefinitions().getByName('Restricted Web Viewers');
                                                        cont.load(restricted, 'Id');
                                                        cont.load(nweb, 'HasUniqueRoleAssignments', 'ServerRelativeUrl');
                                                        
                                                        cont.executeQueryAsync(testMakePerm, testMakePerm);
                                                    }

                                                    function testMakePerm() {
                                                        try {
                                                            // restricted exists, else make it
                                                            restricted.get_id();
                                                        } catch (e) {
                                                            // Set up permissions.
                                                            var permissions = new SP.BasePermissions();
                                                            permissions.set(SP.PermissionKind.viewPages);
                                                            permissions.set(SP.PermissionKind.useRemoteAPIs);
                                                            permissions.set(SP.PermissionKind.useClientIntegration);
                                                            permissions.set(SP.PermissionKind.open);

                                                            // Create a new role definition.
                                                            var roleDefinitionCreationInfo = new SP.RoleDefinitionCreationInformation();
                                                            roleDefinitionCreationInfo.set_name('Restricted Web Viewers');
                                                            roleDefinitionCreationInfo.set_description('Restricted to access only content explicity shared');
                                                            roleDefinitionCreationInfo.set_basePermissions(permissions);
                                                            restricted = cont.get_site().get_rootWeb().get_roleDefinitions().add(roleDefinitionCreationInfo);
                                                        }

                                                        // subsite with unique permissions
                                                        if (nweb.get_serverRelativeUrl() != _spPageContextInfo.siteServerRelativeUrl
                                                            && nweb.get_hasUniqueRoleAssignments()) {
                                                            // ensure internal users have restricted web for redirect and root page access
                                                            var restrictedWeb = SP.RoleDefinitionBindingCollection.newObject(cont);
                                                            restrictedWeb.add(restricted);
                                                            nweb.get_roleAssignments().add(nweb.ensureUser('c:0-.f|rolemanager|spo-grid-all-users/' + _spPageContextInfo.siteSubscriptionId), restrictedWeb);

                                                            // ensure site owners group has access, may get removed by user later but initially should be here
                                                            var fullAdmin = SP.RoleDefinitionBindingCollection.newObject(cont);
                                                            fullAdmin.add(full);
                                                            nweb.get_roleAssignments().add(cont.get_site().get_rootWeb().get_associatedOwnerGroup(), fullAdmin);
                                                        }

                                                        cont.executeQueryAsync(function () {
                                                            var hasPages = 0, tiles = '', home = null;
                                                            p.d.results.forEach(function (page) {
                                                                // modern sites
                                                                // https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/single-part-app-pages?tabs=pnpposh

                                                                // ensure layout
                                                                if (page.File.Name == "Tiles.aspx" || !page.PublishingPageLayout || !page.PublishingPageLayout.Url || ~page.PublishingPageLayout.Url.toLowerCase().indexOf('~sitecollection')) {
                                                                    hasPages++;
                                                                    tiles = page.File.ServerRelativeUrl;
                                                                    var pageItem = cont.get_web().get_lists().getByTitle("Pages").getItemById(page ? page.Id : 2);
                                                                    pageItem.get_file().checkOut();
                                                                    var val = new SP.FieldUrlValue();
                                                                    val.set_url(!page || !page.PublishingPageLayout || !page.PublishingPageLayout.Url
                                                                        ? _spPageContextInfo.siteAbsoluteUrl.replace(/\/$/, '') + '/_catalogs/masterpage/BlankWebPartPage.aspx'
                                                                        : page.PublishingPageLayout.Url.indexOf('_') === 0
                                                                        ? _spPageContextInfo.siteAbsoluteUrl.replace(/\/$/, '') + '/' + page.PublishingPageLayout.Url
                                                                        : page.PublishingPageLayout.Url.toLowerCase().replace('~sitecollection', _spPageContextInfo.siteAbsoluteUrl.replace(/\/$/, '')));
                                                                    val.set_description(!page || !page.PublishingPageLayout || !page.PublishingPageLayout.Description
                                                                        ? 'Blank Web Part page'
                                                                        : page.PublishingPageLayout.Description);
                                                                    pageItem.set_item("PublishingPageLayout", val);
                                                                    pageItem.update();
                                                                    pageItem.get_file().checkIn('', 1);
                                                                    cont.executeQueryAsync(function () {
                                                                        hasPages--;
                                                                        if (hasPages == 0) {
                                                                            // get home page
                                                                            if (tiles != '') {
                                                                                home = nweb.getFileByServerRelativeUrl(tiles).get_listItemAllFields();
                                                                                cont.load(home, 'HasUniqueRoleAssignments');
                                                                                // set as welcome page
                                                                                var folder = nweb.get_rootFolder();
                                                                                folder.set_welcomePage("Pages/Tiles.aspx");
                                                                                folder.update();
                                                                            }
                                                                            // get defs
                                                                            cont.executeQueryAsync(function () {
                                                                                // remove old home file
                                                                                nweb.getFileByServerRelativeUrl(nweb.get_serverRelativeUrl().replace(/\/$/,'') + '/Pages/default.aspx').recycle();
                                                                                // home page needs permissions
                                                                                if (home && !home.get_hasUniqueRoleAssignments()) {
                                                                                    home.breakRoleInheritance(true);
                                                                                    var restrictedHome = SP.RoleDefinitionBindingCollection.newObject(cont);
                                                                                    restrictedHome.add(restricted);
                                                                                    home.get_roleAssignments().add(nweb.ensureUser('c:0-.f|rolemanager|spo-grid-all-users/' + _spPageContextInfo.siteSubscriptionId), restrictedHome)
                                                                                }
                        
                                                                                // execute
                                                                                cont.executeQueryAsync(function () {
                                                                                    return def.resolve();
                                                                                }, function () {
                                                                                    return def.reject();
                                                                                });
                                                                            }, function () {
                                                                                return def.reject();
                                                                            });
                                                                        }
                                                                    }, function () {
                                                                        return def.reject();
                                                                    });
                                                                }
                                                            });

                                                            setTimeout(function () {
                                                                if (hasPages == 0)
                                                                    return def.resolve();
                                                            },1);
                                                        }, function () {
                                                            return def.reject();
                                                        });
                                                    }
                                                }, function () {
                                                    return def.reject();
                                                });
                                            },
                                            function (sender, args) {
                                                return def.reject();
                                            });
                                        },
                                        function (sender, args) {
                                            return def.reject();
                                        });
                                    },
                                    function (sender, args) {
                                        return def.reject();
                                    });
                                },
                                function (sender, args) {
                                    return def.reject();
                                });
                            },
                            function (sender, args) {
                                return def.reject();
                            });
                        },
                        function (sender, args) {
                            return def.reject();
                        });
                    });
            
            return def.promise();
        }
        
        Webs.CreateGroup = function (group, owner, edit) {
            var def = $.Deferred();

			Lists.UpdateDigest()
                .then(function (d) {
                    var item = {
                        "__metadata": {
                            "type": "SP.Group"
                        },
                        Title: group,
                        Description: group,
                        OnlyAllowMembersViewMembership: false
                    };
                    if (owner)
                        item.Owner = {
                            __metadata: { type: 'SP.Principal' },
                            LoginName: owner
                        };
                    $.ajax({
                        url:  _spPageContextInfo.webAbsoluteUrl.replace(/\/$/,'') +
                            '/_api/web/sitegroups',
                        type: 'POST',
                        headers: { 
                            "Accept": "application/json;odata=verbose",
                            "content-type": "application/json;odata=verbose",
                            "X-RequestDigest": d
                        },
                        data: JSON.stringify(item),
                        success: function (data) {
                            var ctx = new SP.ClientContext.get_current();
                            var ngroup = ctx.get_web().get_siteGroups().getById(data.d.Id);
                            var pgroup = ctx.get_web().get_siteGroups().getByName(owner);
                            if (data.d.OwnerTitle != owner)
                                ngroup.set_owner(pgroup);
                            ngroup.set_allowMembersEditMembership(edit == true);
                            ngroup.set_onlyAllowMembersViewMembership(false);
                            ngroup.update();

                            ctx.executeQueryAsync(function () {
                                return def.resolve(data.d.Id, ngroup, ctx);
                            }, function () {
                                return def.reject();
                            });
                        },
                        error: function () {
                            return def.reject();
                        }
                    });
                });
                
            return def.promise();
        }

        Webs.GetGroups = function (select, expand, filter) {
            var def = $.Deferred();
            
            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/,'') +
                    '/_api/web/sitegroups?$top=5000&$expand=' + (expand || '') + '&$select=' + (select || '') + '&$filter=' + (filter || '' ) + '',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    return def.resolve(data.d.results);
                },
                error: function () {
                    return def.reject();
                }
            });
            
            return def.promise();
        }
        
        // wont always create with owner
        Webs.EnsureGroup = function (group, owner) {
            var def = $.Deferred();
            
            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/,'') +
                    '/_api/web/sitegroups?$filter=LoginName eq %27' +
                    group +
                    '%27',
                type: "GET",
                headers: { "accept": "application/json;odata=verbose" },
                success: function(data) {
                    if (data.d.results.length > 0)
                        return def.resolve(data.d.results[0].Id);

                    Webs.CreateGroup(group, owner)
                        .then(function (id, ngroup, ctx) {
                            return def.resolve(id, ngroup, ctx);
                        }, function () {
                            Webs.CreateGroup(group)
                                .then(function (id, ngroup, ctx) {
                                    return def.resolve(id, ngroup, ctx);
                                }, function () {
                                    return def.reject();
                                });
                        });
                },
                error: function () {
                    return def.reject();
                }
            });
            
            return def.promise();
        }

        Webs.AddUsers = function (gid, uids, clear) {
            var def = $.Deferred();
            
            var ctx = new SP.ClientContext.get_current();
            var collGroup = ctx.get_web().get_siteGroups();
            var oGroup = collGroup.getById(gid);
            var users = oGroup.get_users();
            
            var addUsers = function () {
                var added = [];
                for (var i in uids) {
                    if (added.indexOf(uids[i].Id ? uids[i].Id : uids[i]) >= 0)
                        continue;
                    added.push(uids[i].Id ? uids[i].Id : uids[i]);
                    if (uids[i].Id) {
                        users.addUser(ctx.get_web().getUserById(uids[i].Id));
                    } else if (uids[i].Key) {
                        var userCreationInfo = new SP.UserCreationInformation();
                        userCreationInfo.set_loginName(uids[i].Key);
                        userCreationInfo.set_title(uids[i].DisplayText);
                        users.add(userCreationInfo);
                    } else {
                        users.addUser(ctx.get_web().getUserById(uids[i]));
                    }
                }
                ctx.executeQueryAsync(function () {
                    return def.resolve();
                }, function () {
                    return def.reject();
                });
            }

            ctx.load(oGroup);
            ctx.load(users);
            ctx.executeQueryAsync(function () {
                if (clear) {
                    var userEnumerator = users.getEnumerator();
                    while (userEnumerator.moveNext()) {
                        var oUser = userEnumerator.get_current();
                        users.removeById(oUser.get_id());
                    }

                    ctx.executeQueryAsync(addUsers, function () {
                        return def.reject();
                    });
                } else {
                    addUsers();
                }
            }, function () {
                return def.reject();
            });
            
            return def.promise();
        }
    }
}
