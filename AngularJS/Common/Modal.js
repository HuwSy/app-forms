'use strict';
// all modal controllers

{
    Modal = Modal || {};
    {
        Modal.SPModal = function (e, o) {
            if (typeof o == "undefined" || o == null) {
                o = e;
                e = null;
            }
            if (e != null) {
                e.preventDefault ? e.preventDefault() : e.returnValue = false;
                e.stopPropagation ? e.stopPropagation() : false;
            }
            if (typeof (o) == "string")
                o = { url: options };
            if ($('#Modal').length < 1) {
                    $('body').first().append('<div class="Modals">\
                            <div id="Modal" class="modal" style="display: block;">\
                                    <div class="contents">\
                                            <div id="Header" class="header">\
                                                    <span class="close">×</span>\
                                                    <h2 style="color: white;margin-top: 10px;"></h2>\
                                            </div>\
                                            <div class="body">\
                                                    <iframe scrolling="no" seamless="seamless" frameborder="0" marginwidth="0" marginheight="0" allowfullscreen="" width="100%" height="99%" src="about:blank"></iframe>\
                                                    <div class="html" style="padding: 10px;"></div>\
                                            </div>\
                                    </div>\
                            </div>\
                    </div>\
                    <style type=text/css>\
                    .Modals .modal {\
                            display: none; /* Hidden by default */\
                            position: fixed; /* Stay in place */\
                            z-index: 1000001; /* Sit on top */\
                            left: 0;\
                            top: 0;\
                            width: 100vw; /* Full width */\
                            height: 100vh; /* Full height */\
                            overflow: auto; /* Enable scroll if needed */\
                            background-color: rgb(0,0,0); /* Fallback color */\
                            background-color: rgba(0,0,0,0.4); /* Black w/ opacity */\
                            margin-left: 0;\
                    }\
                    \
                    .Modals .contents {\
                            background-color: #fefefe;\
                            margin: 2vh auto auto;\
                            position: relative;\
                            padding: 0;\
                            border: 1px solid #888;\
                            width: ' + (o.width ? (o.width + (typeof o.width == "number" || o.width.match(/^[0-9]*$/) != null ? 'px' : '')) : '50vw') + ';\
                            height: ' + (o.height ? (o.height + (typeof o.height == "number" || o.height.match(/^[0-9]*$/) != null ? 'px' : '')) : '90vh') + ';\
                            box-shadow: 0 4px 8px 0 rgba(0,0,0,0.2),0 6px 20px 0 rgba(0,0,0,0.19);\
                            animation-name: animatetop;\
                            animation-duration: 0.4s;\
                            min-width: 400px;\
                    }\
                    \
                    .Modals .close {\
                            color: #fff;\
                            float: right;\
                            font-size: 28px;\
                            font-weight: bold;\
                    }\
                    \
                    .Modals .close:hover,\
                    .Modals .close:focus {\
                            color: black;\
                            text-decoration: none;\
                            cursor: pointer;\
                    }\
                    \
                    .Modals .header {\
                            padding: 2px 16px;\
                            background-color: #c52d2c;\
                            color: white;\
                            height: 70px;\
                    }\
                    \
                    .Modals .body {\
                            position: relative;\
                            height: calc(100% - 75px);\
                            overflow: auto;\
                            margin: 0 4px;\
                    }\
                    \
                    .Modals @keyframes animatetop {\
                            from {top: -300px; opacity: 0}\
                            to {top: 0; opacity: 1}\
                    }\
                    <style>');
        
                    var iframe = $('#Modal iframe').first();
                    var hframe = iframe.next();
                    var modal = iframe.parents('#Modal').first();
                    var header = modal.find('#Header')[0];
                    var title = modal.find('h2').first();
                    var span = modal.find('span').first();
                    
                    var hide = function (ret) {
                        modal[0].style.display = 'none';
                        document.documentElement.style.msContentZooming = "";
                        document.documentElement.style.touchAction = "";
                        iframe.attr("src", '');
                        hframe.html('');
                        if (o.dialogReturnValueCallback) {
                            window.SP = SP || {};
                            window.SP.UI = SP.UI || {};
                            window.SP.UI.DialogResult = SP.UI.DialogResult || {OK: "OK"};
                            o.dialogReturnValueCallback(ret || 'Cancel');
                        }
                    };
                    
                    span.on('click', hide);
                    
                    function listener(event) {
                        if (event.data == "Cancel" || event.data == "OK")
                            hide(event.data);
                    }
            
                    if (window.addEventListener) {
                        window.addEventListener("message", listener, false);
                    }
            } else {
                var iframe = $('#Modal iframe').first();
                var hframe = iframe.next();
                var modal = iframe.parents('#Modal').first();
                var header = modal.find('#Header')[0];
                var title = modal.find('h2').first();
                var span = modal.find('span').first();
            }
        
            header.style.backgroundColor = o.url && ~o.url.toLowerCase().indexOf('.sharepoint.com') ? null : '#3d505a';
            if (e != null)
                e.srcElement = e.srcElement || e.target;
            title.text((o.title || (e == null || e.srcElement == null ? '' : e.srcElement.parentElement.title || e.srcElement.title || e.srcElement.parentElement.getAttribute('tooltip') || '')).trim());
            document.documentElement.style.msContentZooming = "none";
            document.documentElement.style.touchAction = "none";
            iframe.hide();
            hframe.hide();
            modal[0].style.display = 'block';
        
            if (o.html)
                hframe.html(o.html).show();
            else
                setTimeout(function () {
                    iframe.attr("src", o.url).show();
                    iframe[0].onload = function () {
                        iframe.attr("height", "98%");
                        setTimeout(function () {
                            iframe.attr("height", "99%");
                        }, 10);
                    };
                }, 10);
        }

        Modal.Close = function (ok) {
            /// <summary>Closes a SharePoint modal with a status</summary>
            /// <param name="ok" type="Boolean">Close with status of ok or cancel or override to another status</param>

            var m = document.getElementById("Modal");
            if (typeof SP == "undefined" || typeof SP.UI == "undefined")
                return m != null ? m.remove() : null;
            
            if (typeof (ok) == "boolean" && ok)
                SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.OK);
            else if (typeof (ok) == "boolean" && !ok)
                SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.Cancel);
            else
                SP.UI.ModalDialog.commonModalDialogClose(ok);
        }

        Modal.URL = function (title, url, full, width, height, callback, newModal) {
            /// <summary>Displays a SharePoint modal of a url</summary>
            /// <param name="title" type="String">Title on modal</param>
            /// <param name="url" type="String">URL of modal</param>
            /// <param name="full" type="Boolean">Full screen</param>
            /// <param name="width" type="Number">Width in px</param>
            /// <param name="height" type="Number">Height in px</param>
            event.preventDefault ? event.preventDefault() : event.returnValue = false;
            
            var options = {
                url: url + (url.indexOf("?") > 0 ? "&" : "?") + (newModal || typeof SP == "undefined" ? '' : "IsDlg=1"),
                title: title,
                showMaximized: full,
                dialogReturnValueCallback: callback
            };

            if (width != null) {
                options.width = width;
                options.showMaximized = false;
            }

            if (height != null) {
                options.height = height;
                options.showMaximized = false;
            }

            if (newModal || typeof SP == "undefined")
                return Modal.SPModal(options);
            SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
        }

        Modal.Modal = function (title, html, full, width, height, callback, newModal) {
            /// <summary>Displays a SharePoint modal of html</summary>
            /// <param name="title" type="String">Title on modal</param>
            /// <param name="html" type="String">HTML to display</param>
            /// <param name="full" type="Boolean">Full screen</param>
            /// <param name="width" type="Number">Width in px</param>
            /// <param name="height" type="Number">Height in px</param>
            var def = $.Deferred();

            if (event)
                event.preventDefault ? event.preventDefault() : event.returnValue = false;

            var modal = document.createElement('div');
            if (typeof(html) === "string")
                modal.innerHTML = '\
                <div style="box-sizing: border-box;" id="modal-spinner">\
                    ' +
                    html +
                    '\
                    <p style="min-height: 50px; color: red;" id="seterr"></p>\
                </div>';
            else
                modal = $(html).clone()[0];

            var options = {
                html: modal,
                title: title,
                showMaximized: full,
                dialogReturnValueCallback: callback
            };

            if (width != null) {
                options.width = width;
                options.showMaximized = false;
            }

            if (height != null) {
                options.height = height;
                options.showMaximized = false;
            }

            if (newModal || typeof SP == "undefined")
                Modal.SPModal(options);
            else
                SP.SOD.execute('sp.ui.dialog.js', 'SP.UI.ModalDialog.showModalDialog', options);
            
            setTimeout(function () {
                if (document.getElementById('modal-input') != null)
                    document.getElementById('modal-input').focus();
                if (document.getElementById('dialogTitleSpan') != null || document.getElementById('Modal') != null)
                    return def.resolve();
                else
                    return def.reject();
            }, (newModal || typeof SP == "undefined") ? 10 : 1000);

            return def.promise();
        }

        Modal.Prompt = function (title, label, description, addit, call, full, width, height) {
            /// <summary>Displays a SharePoint modal prompting for info</summary>
            /// <param name="title" type="String">Title on modal</param>
            /// <param name="label" type="String">Label to display before input</param>
            /// <param name="description" type="String">Description after input</param>
            /// <param name="addit" type="String">Additional HTML to display</param>
            /// <param name="call" type="String">Call back function name on click</param>
            /// <param name="full" type="Boolean">Full screen</param>
            /// <param name="width" type="Number">Width in px</param>
            /// <param name="height" type="Number">Height in px</param>

            var html = '\
                <div class="ms-TextField"> \
                    <label class="ms-Label">' +
                label +
                '</label>\
                    <input id="modal-input" class="ms-TextField-field" onkeyup="javascript:{Modal.Alpha(this);} " maxlength="100"> \
                    <span class="ms-TextField-description">' +
                description +
                '</span> \
                </div> \
                ' +
                addit +
                '\
                <input id="setsub" type="button" value="Create" onclick="' +
                call +
                '" style="float: right;" />';

            Modal.Modal(title, html, full, width, height);
        }

        Modal.YESNO = function (title, label, call, width, height, close, addit) {
            /// <summary>Displays a SharePoint modal prompting for yes no</summary>
            /// <param name="title" type="String">Title on modal</param>
            /// <param name="label" type="String">Label to display before input</param>
            /// <param name="call" type="String">Call back function name on click yes</param>
            /// <param name="width" type="Number">Width in px</param>
            /// <param name="height" type="Number">Height in px</param>
            /// <param name="close" type="String">Call back function on click no, null just close</param>
            /// <param name="addit" type="String">Additional HTML to display</param>

            if (addit == null) addit = '';
            var html = '\
                <div class="row" style="margin: 0;padding: 0 overflow:hidden;"> \
                    <div class="col-sm-1" style="padding-right: 0;"><i class="ms-font-su ms-fontColor-yellow ms-Icon ms-Icon--alertOutline" aria-hidden="true"></i></div> \
                    <div class="col-sm-10"><p style="padding-left: 25px" class="ms-u-slideRightIn10 ms-fontSize-xl">' +
                        label +
                    '</p>\
                    ' +
                    addit +
                    '\
                </div>\
                </div> \
                <br/> \
                <div style="padding-left: 400px;" class="row"> \
                    <button class="col-sm-2 ms-Button" onclick="' +
                call +
                '" > \
                        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>\
                        <span class="ms-Button-label">Yes</span>\
                    </button>\
                    <button class="col-sm-2 ms-Button ms-Button--primary" onclick="' + (close != null ? close : typeof SP != "undefined" ? 'SP.UI.ModalDialog.commonModalDialogClose(1, {})' : 'javascript:{document.getElementById("Modal").remove()}') + '">\
                        <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--x"></i></span>\
                        <span class="ms-Button-label">No</span>\
                    </button>\
                </div>';

            Modal.Modal(title, html, false, width, height);
        }

        Modal.Alpha = function (ths) {
            /// <summary>Strips non mms term non folder path safe characters during input in Prompt</summary>
            /// <param name="ths" type="Object">DOM object</param>

            var v = ths.value.replace(/[^a-zA-Z0-9\(\)"£$_\+\@\ \']/g, '-')
                .replace(/- /g, ' ')
                .replace(/ -/g, ' ')
                .replace(/--/g, '-');

            if (v != ths.value) {
                ths.value = v;
                ths.style.borderColor = "#f99";
            } else {
                ths.style.borderColor = "#000";
            }
        }

        Modal.Error = function (r, e, m) {
            /// <summary>Updates the current modal with error details or opens a new with error</summary>
            /// <param name="r" type="Object">When coming from execQuery</param>
            /// <param name="e" type="Object">When coming from execQuery</param>
            /// <param name="m" type="String">Message to display if not from execQuery</param>

            console.log(r);
            console.log(e);

            try {
                $("#s4-workspace").removeSpinner();
                $("#modal-spinner").removeSpinner();
                $("#bible-workflow").removeSpinner();
            } catch (e) {
                // spinner may not be used if this code is reused elsewhere
            }

            if (document.getElementById('seterr') == null)
                Modal.Modal('Error', '');

            document.getElementById('seterr').innerHTML =
                (m != null
                    ? m
                    : "") +
                (e != null
                    ? "<br>Error: " + (typeof e.get_message != "undefined" ? e.get_message() : r.responseText)
                    : "");

            if (document.getElementById('setsub') != null) {
                document.getElementById('setsub').value = 'Create';
                document.getElementById('setsub').disabled = null;
            }
        }

        Modal.Title = function () {
            /// <summary>Get the text from Prompt</summary>
            /// <returns type="String">Entered text</returns>

            if (document.getElementById('seterr') != null)
                document.getElementById('seterr').innerText = '';
            if (document.getElementById('setsub') != null) {
                document.getElementById('setsub').disabled = 'disabled';
                document.getElementById('setsub').value = 'Processing...';
            }

            var ttl = document.getElementById('modal-input').value.trim();
            if (ttl === "") {
                return null;
            }

            return ttl;
        }

        Modal.Fields = function (message) {
            /// <summary>Displays an error message and re-enables the create button</summary>
            /// <param name="message" type="String">Error to display</param>

            if (document.getElementById('setsub') != null) {
                document.getElementById('setsub').value = 'Create';
                document.getElementById('setsub').disabled = null;
            }

            Modal.Error(null, null, message);
        }

        Modal.Redirect = function (url) {
            /// <summary>Redirects to a new URL once the destination exists, returns status 200</summary>
            /// <param name="url" type="String">Destination</param>

            if (document.getElementById('setsub') != null)
                document.getElementById('setsub').value = 'Created';
            if (document.getElementById('seterr') == null)
                Modal.Modal('Redirecting', '');
            if (document.getElementById('seterr').innerText.indexOf("Redirecting shortly") < 0) {
                document.getElementById('seterr').innerText = "Redirecting shortly";
            } else {
                document.getElementById('seterr').innerText += ".";
            }

            // ping destination until response is 200 before loading page to account for mms term store delays
            setTimeout(function () {
                //if ($.active === 0)
                //    Modal.Redirect(url);
                var xhttp = new XMLHttpRequest();
                xhttp.onreadystatechange = function () {
                    if (xhttp.readyState == 4) {
                        if (xhttp.status == 200)
                            window.top.location = url;
                        else
                            Modal.Redirect(url);
                    }
                };
                xhttp.open("GET", url, true);
                xhttp.send();
            }, 2500);
        }
    }
}