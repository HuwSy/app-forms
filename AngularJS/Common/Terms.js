'use strict';
// start loading immediately
if (typeof SP != "undefined")
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', null);

// sp pnp js
// https://pnp.github.io/pnpjs/sp/
    
// all term operations
{
    Terms = Terms || {};
    {
        /// <field name='initalLoad' type='Object'>To ensure PMG code runs without a missing variable</field>  
        window.initalLoad = window.initalLoad || typeof($) != "undefined" ? $.Deferred() : null;

        Terms.UpdateSessionStorage = function (set, ctx, force) {
            /// <summary>Updates the navigtion session storage</summary>
            /// <param name="set" type="Object">String of term set name or object of SspId, TermSetId, FieldTitle</param>
            /// <param name="ctx" type="Object">Client context to use, or current if null</param>
            /// <param name="force" type="Boolean">Forces update of session even if it has not expired yet</param>
            /// <returns type="Object">window.initalLoad deferred object</returns>

            if (ctx == null) {
                ctx = new SP.ClientContext.get_current();
                force = true;
            }

            // hourly expiry on storage item even if session is still open
            var e = sessionStorage.getItem('Expiry.' + _spPageContextInfo.siteServerRelativeUrl);
            if (!force && (e == null || e < new Date())) {
                force = true;
            }

            if (force) {
                sessionStorage.removeItem('Expiry.' + _spPageContextInfo.siteServerRelativeUrl);
                sessionStorage.removeItem('Terms.' + _spPageContextInfo.siteServerRelativeUrl);
            }

            if (sessionStorage.getItem('Terms.' + _spPageContextInfo.siteServerRelativeUrl) == null)
                Terms.GetTermSet(ctx, set)
                    .done(function (terms) {
                        var tree = Terms.GetTermSetAsTree(terms);

                        sessionStorage.setItem('Terms.' +
                            _spPageContextInfo.siteServerRelativeUrl,
                            JSON.stringify(tree, Terms.Replacer));

                        var d = new Date();
                        d.setTime(d.getTime() + 3600 * 1000);
                        sessionStorage.setItem('Expiry.' + _spPageContextInfo.siteServerRelativeUrl, d);
						if (initalLoad != null)
							initalLoad.resolve(d.toString());
                    });
            else if (initalLoad != null)
                initalLoad.resolve(null);

			if (initalLoad != null)
				return initalLoad.promise();
        }

        Terms.Find = function (tree, search) {
            /// <summary>Find a term within 2 levels of the supplied tree from GetTermSetAsTree, will not return MMS .term member if it has been session stored</summary>
            /// <param name="tree" type="Object">Tree object of navigation produced by GetTermSetAsTree</param>
            /// <param name="force" type="String">Term name, or term path colon delimited or term guid to locate</param>
            /// <returns type="Object">Term object with .name, .guid, .term</returns>

            if (search == null || search == "" || search == ":" || tree == null || tree.children == null)
                return null;
            if (typeof (search) != "string")
                search = Terms.GetFieldValue(search).label;

            search = search.replace(/&/g, '\uFF06');
            for (var i in tree.children) {
                if (tree.children[i].name == search || tree.children[i].guid == search)
                    return tree.children[i];
                if (tree.children[i].children != null)
                    for (var s in tree.children[i].children) {
                        if (tree.children[i].name + ':' + tree.children[i].children[s].name == search || tree.children[i].children[s].guid == search)
                            return tree.children[i].children[s];
                    }
            }

            return null;
        }

        Terms.GetCurrentTerm = function (ctx, nameOfTermSets, properties) {
            /// <summary>Gets the first investment term associated to the current folder</summary>
            /// <param name="ctx" type="Object">Client context to use, or current if null</param>
            /// <param name="nameOfTermSets" type="Object">String of term set name or object of SspId, TermSetId, FieldTitle</param>
            /// <param name="properties" type="Object">List item of required object to locate term from</param>
            /// <returns type="Object">Deferred returning term, parent term</returns>

            if (properties != null) ctx = properties.get_context();
            if (ctx == null) ctx = new SP.ClientContext.get_current();
            if (typeof (nameOfTermSets) == "string" || nameOfTermSets.TermSetId) nameOfTermSets = [nameOfTermSets];
            nameOfTermSets = nameOfTermSets.filter(function (item, pos, self) {
                return self.indexOf(item) == pos;
            });

            var retterm = [];
            var retpar = [];
            var count = 0;
            var def = $.Deferred();

            nameOfTermSets.forEach(function (n) {
                var p = properties != null
                    ? Terms.GetFieldValue(properties.get_fieldValues()[n.FieldTitle || n.replace(/ /g, '')]).guid
                    : null;
                Terms.GetTermObjectByNameId(ctx, null, p, n)
                    .done(function (term, parent) {
                        if (nameOfTermSets.length == 1)
                            return def.resolve(term, parent);

                        retterm[n.FieldTitle || n] = term;
                        retpar[n.FieldTitle || n] = parent;
                        count++;
                        if (count == nameOfTermSets.length)
                            return def.resolve(retterm, retpar);
                    })
                    .fail(function () {
                        return def.reject();
                    });
            });

            return def.promise();
        }

        Terms.GetTermObjectByNameId = function (ctx, name, guid, nameOfTermSet) {
            /// <summary>Gets a specific investment term by name or guid, if neither specified returns the term set</summary>
            /// <param name="ctx" type="Object">Client context to use, or current if null</param>
            /// <param name="name" type="String">Term name to locate</param>
            /// <param name="guid" type="String">Guid to locate</param>
            /// <param name="nameOfTermSets" type="Object">String of term set name or object of SspId, TermSetId, FieldTitle</param>
            /// <returns type="Object">Deferred returning term, parent term. Reject returns termset</returns>

            var def = $.Deferred();
            if (ctx == null) ctx = new SP.ClientContext.get_current();

            Terms.GetTermSet(ctx, nameOfTermSet)
                .then(function (terms, termSet) {
                    var term = guid ? terms.getById(guid) : name ? terms.getByName(name) : termSet;
                    var parent = guid || name ? term.get_parent() : termSet;
                    ctx.load(term);
                    ctx.load(term.get_terms());
                    ctx.load(parent);

                    ctx.executeQueryAsync(function () {
                        return def.resolve(term,
                            typeof (term.get_isRoot) == "undefined" || term.get_isRoot() ? termSet : parent);
                    },
                        function () {
                            return def.reject(termSet);
                        });
                },
                function () {
                    return def.reject(null);
                });

            return def.promise();
        }

        Terms.CreateTermWithinContext = function (ctx, nameOfTermSet, newTermName, parentTerm, alias, classification, noCustomProps, extraProps) {
            /// <summary>Creates a new term under the prent term and associates it to a folder path and optionally an alias of term</summary>
            /// <param name="ctx" type="Object">Client context to use, or parent term context, or current if null</param>
            /// <param name="nameOfTermSets" type="Object">String of term set name or object of SspId, TermSetId, FieldTitle</param>
            /// <param name="newTermName" type="String">New term name to create</param>
            /// <param name="parentTerm" type="Object">Parent term or set for new term, SharePoint MMS term object</param>
            /// <param name="alias" type="Object">Term object to create as an alias of, or string to use in alias field</param>
            /// <returns type="Object">Deferred returning term, context</returns>

            if (parentTerm != null) ctx = parentTerm.get_context();
            if (ctx == null) ctx = new SP.ClientContext.get_current();

            var def = $.Deferred();
            var newGuid = new SP.Guid.newGuid();

            Terms.GetTermObjectByNameId(ctx, newTermName, null, nameOfTermSet)
                .fail(function(termSet) {
                    if (parentTerm == null) {
                        parentTerm = termSet;
                    }
                    if (parentTerm == null) {
                        return def.reject();
                    }

                    ctx.load(parentTerm);
                    ctx.executeQueryAsync(function() {
                            var newTerm = parentTerm.createTerm(newTermName, 1033, newGuid);

                            // parent url if exists, does not fail nicely or have any checking so within try catch
                            var url = '/' + newTermName,
                                web = null;

                            // load frob term
                            try {
                                if (typeof (alias) == "object") {
                                    url = alias
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Folder Path"];
                                    if (url == null || url == '')
                                        url = ('/' + alias.get_name());
                                    web = alias
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Web"] ||
                                        '';
                                    classification = classification ||
                                        alias
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Classification"] ||
                                        '';
                                    alias = alias
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Alias"];
                                    if (alias == null || alias == '')
                                        alias = Lists.Listtitle(url).path;
                                } else if (parentTerm != null && typeof (parentTerm.get_isRoot) != 'undefined') {
                                    var u = parentTerm
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Folder Path"];
                                    if (u == null || u == '')
                                        u = '/' + parentTerm.get_name();
                                    url = u + url;
                                    web = parentTerm
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Web"] ||
                                        '';
                                    classification = classification ||
                                        parentTerm
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Classification"] ||
                                        '';
                                    alias = parentTerm
                                        .get_objectData()
                                        .get_properties()["CustomProperties"]["Alias"];
                                    if (alias == null || alias == '')
                                        alias = Lists.Listtitle(url).path;
                                }
                            } catch (e) {
                                if (typeof (alias) == "object")
                                    alias = null;
                            }

                            if (!noCustomProps) {
                                newTerm.setCustomProperty('Folder Path', url);
                                newTerm.setCustomProperty('Alias', alias || Lists.Listtitle(newTermName).path);
                                newTerm.setCustomProperty('Web', web || ctx.get_url().substring(_spPageContextInfo.siteServerRelativeUrl.length));
                                newTerm.setCustomProperty('Classification', classification);
                            }

                            if (extraProps) {
                                for (var p in extraProps) {
                                    newTerm.setLocalCustomProperty(p, extraProps[p]);
                                }
                            }

                            ctx.load(newTerm);
                            ctx.executeQueryAsync(function() {
                                    return def.resolve(newTerm, ctx);
                                },
                                function(request, args) {
                                    return def.reject(newTerm, false);
                                });
                        },
                        function() {
                            return def.reject();
                        });
                })
                .done(function(term) {
                    return def.reject(term, true);
                });

            return def.promise();
        };

        Terms.MoveTerm = function (term, parent) {
            var ctx = term.get_context();
            var def = $.Deferred();
            term.move(parent);
            ctx.load(term);
            ctx.executeQueryAsync(function () {
                    return def.resolve();
                },
                function (request, args) {
                    return def.reject();
                });

            return def.promise();
        }

        Terms.GetFieldValue = function (f) {
            /// <summary>Converts a list field value into a term object regardless of current locale</summary>
            /// <param name="f" type="Object">Term object from mms or field or term object</param>
            /// <returns type="Object">Term guid, label, wssid</returns>

            // null object in
            if (f == null)
                return {
                    guid: null,
                    label: null,
                    wssid: null
                };

            // already a term object
            if (f.guid && f.label && f.wssid)
                return f;

            // mms term store object
            if (typeof (f.get_id) != "undefined")
                return {
                    guid: f.get_id()._m_guidString$p$0,
                    label: typeof (f.get_name) != "undefined" ? f.get_name() : '',
                    wssid: '-1'
                };

            // field value
            if (typeof(f.get_objectData) != "undefined"
                && typeof(f.get_objectData().get_$12_0) != "undefined"
                && f.get_objectData().get_$12_0() != null
                && f.get_objectData().get_$12_0().length > 0
                && typeof(f.get_objectData().get_$12_0()[0].get_termGuid) != "undefined")
                return {
                    guid: f.get_objectData().get_$12_0()[0].get_termGuid(),
                    label: f.get_objectData().get_$12_0()[0].get_label(),
                    wssid: f.get_objectData().get_$12_0()[0].get_wssId()
                };
                
            var r = f._Child_Items_ != null && f._Child_Items_.length > 0
                // multi choice column in hidden list
                ? f._Child_Items_[0]
                : f.$2_0 != null && f.$2_0.$1E_0 != null && f.$2_0.$1E_0.length > 0
                    // multi choice column not in hidden list yet
                    ? f.$2_0.$1E_0[0]
                    // single choice column
                    : f;

            // ones in and not in hidden list
            return {
                guid: r.TermGuid || r.$1_1 || null,
                label: r.Label || r.$0_1 || '',
                wssid: r.WssId || r.$2_1 || '-1'
            };
        }

        Terms.ToString = function (t) {
            /// <summary>Converts term to a string value for setting against list items regardless of current localle</summary>
            /// <param name="f" type="Object">Term object from mms or field or term object</param>
            /// <returns type="Object">Term tring ;# | format</returns>

            if (t == null) {
                return null;
            } else if (typeof (t) == "string") {
                return t;
            } else if (typeof (t) == "object") {
                var f = Terms.GetFieldValue(t);
                if (f.guid == null)
                    return null;
                return f.wssid +
                    ';#' +
                    f.label +
                    '|' +
                    f.guid;
            } else {
                return null;
            }
        }

        Terms.GetFieldTermIds = function (listName, fieldName, isSite) {
            /// <summary>Get field term sspid and set</summary>
            /// <param name="listName" type="String">List title</param>
            /// <param name="fieldName" type="String">Field title</param>
            /// <param name="isSite" type="Boolean">True to look in the site collection root instead of the current web</param>
            /// <returns type="Object">Deferred returning {SspId, TermSetId, FieldTitle}</returns>

            return Lists.GetField(listName, fieldName, isSite, 'SspId,TermSetId');
        }

        Terms.GetTermSet = function (ctx, name) {
            /// <summary>Gets a termset, based on name, then generally use GetTermSetAsTree</summary>
            /// <param name="ctx" type="Object">Client context to use, or current if null</param>
            /// <param name="name" type="Object">String of term set name or object of SspId, TermSetId, FieldTitle</param>
            /// <returns type="Object">Deferred returning term set</returns>

            if (ctx == null) ctx = new SP.ClientContext.get_current();
            var def = $.Deferred();

            // Make sure taxonomy library is registered
            SP.SOD.registerSod('sp.taxonomy.js', '\u002f_layouts\u002f15\u002fSP.Taxonomy.js');
            SP.SOD.loadMultiple(['sp.taxonomy.js'], function () {
                var grp = $.Deferred(),
                    site = ctx.get_site(),
                    taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx);

                if (typeof (name) != "string") {
                    var stores = taxonomySession.get_termStores(),
                        store = stores.getById(name.SspId || name.termStoreId),
                        termSet = store.getTermSet(name.TermSetId || name.termSetId);

                    ctx.executeQueryAsync(function () {
                        grp.resolve(termSet);
                    },
                        function () {
                            return def.reject();
                        });
                } else {
                    var termStore = taxonomySession.getDefaultSiteCollectionTermStore(),
                        termStoreGroup = termStore.getSiteCollectionGroup(site);

                    ctx.executeQueryAsync(function () {
                        var termSets = termStoreGroup.get_termSets(name, 1033);
                        return grp.resolve(termSets.getByName(name));
                    },
                        function () {
                            var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(ctx),
                                termSets = taxonomySession.getTermSetsByName(name, 1033);
                            return grp.resolve(termSets.getByName(name));
                        });
                }

                $.when(grp)
                    .then(function (termSet) {
                        var terms = termSet.getAllTerms();
                        ctx.load(terms);
                        ctx.load(termSet);

                        ctx.executeQueryAsync(Function.createDelegate(this,
                            function () {
                                return def.resolve(terms, termSet);
                            }),
                            function () {
                                return def.reject();
                            });
                    });
            });

            return def.promise();
        };

        Terms.GetTermSetAsTree = function (terms, dontbuild) {
            /// <summary>Builds as a tree</summary>
            /// <param name="terms" type="Object">Term set all terms object</param>
            /// <param name="dontbuild" type="Boolean">Dont mock/fill-in terms in term path automatically</param>
            /// <returns type="Object">Tree object containing MMS term, name, guid, level, folderpath, alias, web and children and repeats.</returns>

            var termsEnumerator = terms.getEnumerator(),
                tree = {
                    term: terms,
                    children: []
                };

            // Loop through each term
            while (termsEnumerator.moveNext()) {
                var currentTerm = termsEnumerator.get_current();
                if (!currentTerm.get_isAvailableForTagging()
                    || ~(currentTerm.get_objectData().get_properties().LocalCustomProperties._Sys_Nav_ExcludedProviders || '').indexOf('"CurrentNavigationTaxonomyProvider"'))
                    continue;

                var children = tree.children;
                var currentTermPath = dontbuild ? [currentTerm.get_pathOfTerm().split(';').pop()] : currentTerm.get_pathOfTerm().split(';');

                // Loop through each part of the path
                for (var i = 0; i < currentTermPath.length; i++) {
                    var foundNode = false;
                    for (var j = 0; j < children.length; j++) {
                        if (children[j].name === currentTermPath[i]) {
                            foundNode = true;
                            break;
                        }
                    }

                    // Select the node, otherwise create a new one
                    var term = foundNode ? children[j] : { name: currentTermPath[i], children: [] };
                    var termLevel = i + 1;

                    // If we're a child element, add the term properties
                    if (i === currentTermPath.length - 1) {
                        term.term = currentTerm;

                        // prefere current local display name
                        if (currentTerm.get_name() != null && currentTerm.get_name() != "")
                            term.name = currentTerm.get_name();

                        term.guid = currentTerm.get_id().toString();
                        term.level = termLevel;

                        // ensure '/' prefix on folder paths
                        term.folderPath = currentTerm.get_objectData().get_properties()["CustomProperties"]["Folder Path"] != undefined
                            && currentTerm.get_objectData().get_properties()["CustomProperties"]["Folder Path"] != ''
                                ? currentTerm.get_objectData().get_properties()["CustomProperties"]["Folder Path"]
                                : "/" + currentTerm.get_name();
                        if (term.folderPath != null && term.folderPath != "" && term.folderPath.charAt(0) != '/')
                            term.folderPath = '/' + term.folderPath;

                        // always include alias as the library letter to improve performance later
                        var termChar = term.folderPath.replace('/', '').charAt(0).toUpperCase();
                        if (!/[A-Z]/.test(termChar))
                            termChar = '09';
                        term.alias = currentTerm.get_objectData().get_properties()["CustomProperties"]["Alias"] != undefined
                            && currentTerm.get_objectData().get_properties()["CustomProperties"]["Alias"] != ''
                            ? currentTerm.get_objectData().get_properties()["CustomProperties"]["Alias"].toUpperCase()
                            : termChar;

                        // get sub web url
                        term.web = currentTerm.get_objectData().get_properties()["CustomProperties"]["Web"] != undefined
                            ? currentTerm.get_objectData().get_properties()["CustomProperties"]["Web"]
                            : '';

                        // get sub web url
                        term.Classification = currentTerm.get_objectData().get_properties()["CustomProperties"]["Classification"] != undefined
                            ? currentTerm.get_objectData().get_properties()["CustomProperties"]["Classification"]
                            : '';
                    }

                    // If the node did exist, let's look there next iteration
                    if (foundNode) {
                        children = term.children;
                    }
                    // If the segment of path does not exist, create it
                    else {
                        children.push(term);

                        // Reset the children pointer to add there next iteration
                        if (i !== currentTermPath.length - 1) {
                            children = term.children;
                        }
                    }
                }
            }

            return Terms.SortTermsFromTree(tree);
        };

        Terms.SortTermsFromTree = function (tree) {
            /// <summary>Sort children array of a term tree by a sort order</summary>
            /// <param name="terms" type="Object">Tree object</param>
            /// <returns type="Object">Tree object</returns>

            // Check to see if the get_customSortOrder function is defined. If the term is actually a term collection,
            // there is nothing to sort.
            if (tree.children.length) {
                var sortOrder = null;

                if (tree.term && tree.term.get_customSortOrder && tree.term.get_customSortOrder()) {
                    sortOrder = tree.term.get_customSortOrder();
                }

                // If not null, the custom sort order is a string of GUIDs, delimited by a :
                if (sortOrder) {
                    sortOrder = sortOrder.split(':');

                    tree.children.sort(function (a, b) {
                        var indexA = sortOrder.indexOf(a.guid);
                        var indexB = sortOrder.indexOf(b.guid);

                        if (indexA > indexB) {
                            return 1;
                        } else if (indexA < indexB) {
                            return -1;
                        }

                        return 0;
                    });
                }
                // If null, terms are just sorted alphabetically
                else {
                    tree.children.sort(function (a, b) {
                        if (a.name > b.name) {
                            return 1;
                        } else if (a.name < b.name) {
                            return -1;
                        }

                        return 0;
                    });
                }
            }

            for (var i = 0; i < tree.children.length; i++) {
                tree.children[i] = Terms.SortTermsFromTree(tree.children[i]);
            }

            return tree;
        };

        Terms.Replacer = function (key, value) {
            /// <summary>Passed as a parameter into JSON.stringify to remove circular refernces</summary>
            /// <param name="key" type="Object">Key</param>
            /// <param name="value" type="Object">Value</param>
            /// <returns type="Object">Value</returns>

            if (key == "term") return undefined;
            else if (key == "privateProperty2") return undefined;
            else return value;
        }

        Terms.CleanToUrl = function (t) {
            /// <summary>Gets the friendly URL for the term segment, must be patched to parent url to be useful</summary>
            try {
                t.get_isRoot();
                return (t.get_objectData().get_properties()["LocalCustomProperties"]["_Sys_Nav_FriendlyUrlSegment"] || t.get_name().trim().replace(/[ !`%&]/g, '-').replace(/"/g, '＂').replace(/--/g, '-').replace(/[^A-Za-z0-9¬\-＂£$-\(\)_+@ ]/g, '')) + '/';
            } catch (e) {
                return '';
            }
        }
    }
}
