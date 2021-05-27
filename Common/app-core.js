"use strict";
// no conflict jQuery
if (typeof APP$ == "undefined")
    window.APP$ = $.noConflict();
// hide ribbon in modals
//if (~document.location.href.toLowerCase().indexOf('isdlg=1'))
//    document.getElementById('s4-ribbonrow').remove();
// suppress white space in classic dialogs
if (~document.location.href.indexOf('IsDlg=1') && document.getElementById('globalNavBox') != null)
    document.getElementById('globalNavBox').style.height = '40px';
// force redirect if from old syntax ie email that hasnt been updated
if (document.location.hash.indexOf('#/') == 0)
    document.location.hash = document.location.hash.replace('#/','#!/');