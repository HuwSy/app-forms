"use strict";
// if no app module make the module, else bootstrap
if (typeof angular != "undefined") {
    if (typeof app == "undefined")
        window.app = angular
                        .module("App", ['ngRoute','ngMaterial','ui.calendar','ui.bootstrap','ui.tinymce'])
                        // filter to allow html to be written from within angular variable
                        .filter('to_trusted', ['$sce', function ($sce) {
                            return function (text) {
                                return $sce.trustAsHtml(text || '');
                            };
                        }]);
    else if (document.getElementById('App') != null && (document.getElementById('App').getAttribute('ng-app') || '') == '')
        setTimeout(function () {
            if (typeof InEdit == "undefined" || !InEdit()) {
                var mods = ['choice','enter','error','maxlength','people','scroll','tooltip','upload','appforms','Forms','Home','App'];
                angular.bootstrap(document.getElementById('App'), mods.filter(function (m) {try {angular.module(m); return true} catch (e) {return false}}));
            }
        },10);
}
