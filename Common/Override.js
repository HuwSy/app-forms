'use strict';
// all overrides and scripts copied from elsewhere
if ((typeof ($) === "undefined" || typeof ($.fn) === "undefined") && typeof (APP$) !== "undefined")
    window.$ = APP$;
if (typeof (APP$) === "undefined" && typeof ($.fn) !== "undefined")
    window.APP$ = $;

// suppress white space in classic dialogs
if (~document.location.href.indexOf('IsDlg=1'))
    document.getElementById('globalNavBox').style.height = '40px';

// base page context, there are the min requirement for *.js
/*
if (!_spPageContextInfo) {
    let w = await pnp.sp.web();
    let s = await pnp.sp.site();
    let u = await pnp.sp.web.currentUser.get();
    _spPageContextInfo = {
        webAbsoluteUrl: window.origin + w.ServerRelativeUrl,
        webServerRelativeUrl: w.ServerRelativeUrl,
        siteAbsoluteUrl: window.origin + s.ServerRelativeUrl,
        siteServerRelativeUrl: s.ServerRelativeUrl,
        userId: u.Id,
        webTitle: w.Title,
        isSPO: true,
        
        // not possible... rethink
        listTitle: '',
        listUrl: '',
        serverRequestPath: ''
    }
}
*/

{
    Override = Override || {};
    {
		/// <summary>Load current user permission groups into object for quicker checking</summary>
		Override.Permissions = {};
        if (_spPageContextInfo && _spPageContextInfo.webAbsoluteUrl)
            $.ajax({
                url: _spPageContextInfo.webAbsoluteUrl.replace(/\/$/, '') + '/_api/web/currentuser/groups?$select=Title,Id',
                method: "GET",
                headers: {
                    "accept": "application/json;odata=verbose",
                    "content-Type": "application/json;odata=verbose"
                },
                success: function (d) {
                    Override.Permissions[null] = true;
                    Override.Permissions[_spPageContextInfo.userId] = true;
                    for (var g = 0; g < d.d.results.length; g++) {
                        Override.Permissions[d.d.results[g].Id] = true;
                        Override.Permissions[d.d.results[g].Title.replace(/ /g, '-')] = true;
                        if (d.d.results[g].Title.indexOf(_spPageContextInfo.webTitle + " ") == 0)
                            Override.Permissions[d.d.results[g].Title.replace(_spPageContextInfo.webTitle + " ",'').replace(/ /g, '-')] = true;
                    }
                }
            });

        Override.Lasta = function (ths) {
            /// <summary>Clicks the last active A on the same row for actions on dashboards</summary>

            event.preventDefault();
            document.location.href = $(ths || event.srcElement || event.target).parent().parent().find('a[href]').not('.inactive').not('.ng-hide').last().attr('href');
            return false;
        }

        Override.ParameterByName = function (name, url) {
            /// <summary>Gets specified parameter from query string</summary>
            /// <param name="name" type="String">Query string param</param>
            /// <param name="url" type="String">Optional search within this url, null use current address bar</param>
            /// <returns type="String">Param value</returns>
        
            if (!url) url = window.location.href;
            name = name.replace(/[\[\]]/g, "\\$&");
            var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
                results = regex.exec(url);
            if (!results) return null;
            if (!results[2]) return '';
            return decodeURIComponent(results[2].replace(/\+/g, " ")).toString();
        }
    }
}

// init origin for IE
if (!window.location.origin) {
    window.location.origin = window.location.protocol +
        "//" +
        window.location.hostname +
        (window.location.port ? ':' + window.location.port : '');
}

if (!String.prototype.includes) {
    String.prototype.includes = function (search, start) {
        /// <summary>Does the string include the search term case sensitive</summary>
        /// <param name="search" type="String">Search string</param>
        /// <param name="start" type="Number">Begining position</param>
        /// <returns type="Boolean">Returns true or false</returns>

        if (typeof start !== 'number') {
            start = 0;
        }

        if (start + search.length > this.length) {
            return false;
        } else {
            return this.indexOf(search, start) !== -1;
        }
    };
}

if (!String.prototype.endsWith) {
    String.prototype.endsWith = function (search) {
        /// <summary>Does the string end with the search term case sensitive</summary>
        /// <param name="search" type="String">Search string</param>
        /// <param name="start" type="Number">Begining position</param>
        /// <returns type="Boolean">Returns true or false</returns>
        
        return this.indexOf(search) === this.length - search.length;
    };
}

$.fn.setSpinner = function () {
    /// <summary>Creates loading spinner for spinner.js</summary>

    var opts = { color: 'red' };
    var spinner = new Spinner(opts).spin();
    var element = $(this);
    element.append(spinner.el);
    element.find('input, textarea, button, select').attr('disabled', true);
}

$.fn.removeSpinner = function () {
    /// <summary>Removes loading spinner for spinner.js</summary>

    var element = $(this);
    element.find('.spinner').remove();
    element.find('input, textarea, button, select').attr('disabled', false);
}

$(document).ready(function() {
    /// <summary>Ensures any webparts added to classic library pages dont effect the ribbon showing</summary>
    
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
		} catch (e) {}
	}
});
