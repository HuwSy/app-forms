"use strict";
(function () {
    var appMod = angular.module('choice', []);

    appMod.directive('choice',
        function () {
            /// <summary>Pre made field formats, doesn't support rich text fields or people fields</summary>
            return {
                restrict: 'E',
                scope: {
                    object: '=', // form containing field
                    field: '@', // field name on object,
                    relative: '@', // type, desc, title etc are prefixed with this

                    override: '@', // override any type, desc, title for this field. sent as string as passing object kills large form performance for the few tiles this is useful

                    // choice only options
                    none: '@', // none option text in addition or instead of null
                    radio: '=', // override choice to radio buttons, doesn't support not fill-in below
                    hides: '=', // hide the Other fill-in option
                    other: '@', // override Other fill-in option text to something else, will override hides

                    ngShow: '=',
                    ngHide: '=',
                    ngDisabled: '='
                },
                template: `
<div style="margin: 12px 0 0 0;">
    <label ng-if="get('TypeAsString') == 'Boolean'">
        <input flex ng-disabled="ngDisabled" type="checkbox" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_Y" style="height: 10px !important;min-height: 0 !important;">
        {{get('Titles')}}
    </label>

    <span class="cTitle" ng-if="get('TypeAsString') != 'Boolean'">{{get('Titles')}}</span>
    <span ng-if="get('Requireds')" class=required>*</span>
    <i ng-if="t" class="fa fa-info-circle" aria-hidden="true" tooltip="{{t}}"></i>
    
    <div layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Text'">
        <input flex ng-disabled="ngDisabled" type="text" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_T" maxlength="250">
    </div>
    
    <div ng-hide="ngDisabled" layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Number' || get('TypeAsString') == 'Integer'">
        <input flex ng-disabled="ngDisabled" type="number" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_I">
    </div>
    
    <div ng-show="ngDisabled" layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Number' || get('TypeAsString') == 'Integer'">
        <div class="disabled">{{object[field] | number:(((object[field] || 0).toString() + '.').split('.')[1].length)}}</div>
    </div>
    
    <div layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Date'">
        <input flex ng-disabled="ngDisabled" type="date" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_D">
    </div>
    
    <div layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'DateTime'">
        <input flex ng-disabled="ngDisabled" type="datetime-local" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_D">
    </div>
    
    <div layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Multiple lines of text'">
        <textarea rows="5" flex ng-disabled="ngDisabled" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_M"></textarea>
    </div>
    
    <div ng-hide="ngDisabled" layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Note'">
        <textarea ui-tinymce="tinymceOptions" rows="5" flex ng-disabled="ngDisabled" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_N"></textarea>
    </div>
    
    <div ng-show="ngDisabled" layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == 'Note'">
         <textarea ui-tinymce="tinymceROOptions" rows="5" flex ng-disabled="ngDisabled" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_N"></textarea>
    </div>
    
    <div layout="column" style="padding-top: 6px;" ng-if="get('TypeAsString') == null || get('TypeAsString') == 'User' || get('TypeAsString') == 'UserMulti' || get('TypeAsString') == 'Lookup' || get('TypeAsString') == 'LookupMulti' || get('TypeAsString') == 'TaxonomyFieldType' || get('TypeAsString') == 'TaxonomyFieldTypeMulti'">
        <div class="disabled">Error, not yet available...</div>
    </div>
    
    <div layout="row" style="padding-top: 6px;" ng-show="get('TypeAsString') == 'Choice' && radio">
        <label ng-if="none" style="display:block; margin-right: 15px">
            <input value="-" ng-disabled="ngDisabled" type="radio" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_R" style="height: 10px !important;min-height: 0 !important;">
            {{none}}
        </label>
        <label ng-repeat="i in get('Choices')" style="display:block; margin-right: 15px">
            <input ng-value="i" ng-disabled="ngDisabled" type="radio" ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_R" style="height: 10px !important;min-height: 0 !important;">
            {{i}}
        </label>
    </div>
    
    <div layout="row" style="padding-top: 6px;" ng-show="get('TypeAsString') == 'Choice' && !radio">
        <select ng-disabled="ngDisabled" ng-change="selChangeS()" ng-model="c" ng-required="r" name="{{(relative || '') + field}}_0" flex>
            <option ng-if="none" value="-">{{none}}</option>
            <option ng-repeat="i in get('Choices')" ng-value="i" ng-if="!other || other != i">{{i}}</option>
            <option ng-if="!hides || other" ng-value="v">{{other || 'Other'}}</option>
        </select>
        <input ng-model-options="{debounce:2500}" ng-disabled="ngDisabled" type="text" ng-if="(!hides || other) && get('TypeAsString') == 'Choice' && !radio" ng-show="s" flex ng-model="object[field]" ng-required="r" name="{{(relative || '') + field}}_1" maxlength="250" onclick="this.select();">
    </div>
    
    <div layout="column" style="padding-top: 6px;" ng-show="get('TypeAsString') == 'MultiChoice'">
        <select flex multiple="multiple" style="height: 80px !important" ng-disabled="ngDisabled" ng-click="selChangeM($event)" ng-model="c" ng-required="r" name="{{(relative || '') + field}}_S">
            <option ng-repeat="i in get('Choices')" ng-value="i">{{i}}</option>
        </select>
    </div>
    
    <span style="display:none !important" class="fields" id="{{(relative || '') + field}}">{{object[field]}}</span>
</div>`,
                link: function(scope, element) {
                    element.addClass('layout-column');
                    // IE only hack
                    if (element.attr('flex') == null) {
                        element.attr('style',(element.attr('style') || '') + ';display: block');
                    }
                },
                controller: function ($scope, $timeout, $element) {
                    // define the default other value
                    $scope.v = ($scope.other || 'Other') + ', please specify...';

                    // rich text field
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
                    
                    // readonly/disabled rich text field
                    $scope.tinymceROOptions = {
                        selector: "textarea",
                        height: 200,
                        menubar: false,
                        toolbar: false,
                        statusbar: false
                    };

                    // gets the required field properties and/or any overrides
                    var override = null;
                    $scope.get = function (t) {
                        var p = null;
                        if (!override && $scope.override)
                            override = JSON.parse($scope.override);
                        if (override)
                            p = override[t];
                        if (!p && $scope.$parent && $scope.$parent[t])
                            p = $scope.$parent[t][($scope.relative || '') + $scope.field];
                        if (!p && t == 'Choices')
                            return [];
                        return p;
                    }

                    // on single selection change
                    $scope.selChangeS = function () {
                        // may get triggered on other type changes but ignore
                        if ($scope.get('TypeAsString') != 'Choice')
                            return;
                        
                        // set field value to match c either immediately 
                        if ($scope.c != $scope.v)
                            return $scope.object[$scope.field] = $scope.c;
                        // or slight delay for 'other' option
                        $timeout(function(){
                            $scope.object[$scope.field] = $scope.c;
                        },1);
                    }

                    // on multi selection change
                    $scope.selChangeM = function (event) {
                        // may get triggered on other type changes but ignore
                        if ($scope.get('TypeAsString') != 'MultiChoice')
                            return;

                        // set c to match new field value after creating below
                        $timeout(function(){
                            $scope.c = $scope.object[$scope.field] == null ? null : $scope.object[$scope.field].results;
                        },10);

                        // if no results set the field to the multiselect value
                        if (!$scope.object[$scope.field] || !$scope.object[$scope.field].results)
                            return $scope.object[$scope.field] = {
                                __metadata: {type: "Collection(Edm.String)"},
                                results: $scope.c && $scope.c.length >= 1 ? [$scope.c[0]] : []
                            }

                        // if there are selected results set the field to add/remove the most recent click
                        var i = $scope.object[$scope.field].results.indexOf(event.target.value.replace('string:',''));
                        if (~i)
                            $scope.object[$scope.field].results.splice(i,1);
                        else
                            $scope.object[$scope.field].results.push(event.target.value.replace('string:',''));
                    }

                    // on updates to model, field properties, visibility, override etc trigger to ensure no show, required state and info text are correct
                    // but dont run too much so set timeout for too many calls
                    var time = null;
                    var update = function () {
                        if (time)
                            clearTimeout(time);
                        time = $timeout(function () {
                            // fix ngShow needing to be explicit bool
                            if ($element && $element.attr('ng-show') && $element.attr('ng-show') != '' && $scope.ngShow == null)
                                $scope.ngShow = false;
                            if ($element && $element.attr('data-ng-show') && $element.attr('data-ng-show') != '' && $scope.ngShow == null)
                                $scope.ngShow = false;
                                // re-render the required based on enablement and info bubble from desc
                                $scope.r = $scope.get('Requireds') && ($scope.ngShow || $scope.ngShow == null) && !$scope.ngHide && !$scope.ngDisabled;
                                $scope.t = $scope.get('Descriptions');
                        },50).$$timeoutId;
                    }

                    $scope.$watch('s', update);
                    $scope.$watch('ngShow', update);
                    $scope.$watch('ngHide', update);
                    $scope.$watch('ngDisabled', update);
                    $scope.$watch(`override`, update);
                    $scope.$watch(`$parent.Titles[(relative || '') + field]`, update);
                    $scope.$watch(`$parent.Requireds[(relative || '') + field]`, update);
                    $scope.$watch(`$parent.Descriptions[(relative || '') + field]`, update);
                    $scope.$watch(`$parent.TypeAsString[(relative || '') + field]`, update);

                    // as field changes, ensure choice drop down and other fill-in setup correctly
                    $scope.$watch('object[field]',function () {
                        if (!$scope.object || $scope.object[$scope.field] == null) {
                            $scope.c = null;
                            $scope.s = null;
                            update();
                        } else if ($scope.get('TypeAsString') == 'MultiChoice') {
                            $scope.c = $scope.object[$scope.field].results;
                            $scope.s = false;
                            $scope.r = false; // dirty hack to remove required from actual inputs once populated due to timing bug
                        } else if ($scope.get('TypeAsString') == 'Choice' && ($scope.object[$scope.field] == '-' || $scope.get('Choices').indexOf($scope.object[$scope.field]) >= 0)) {
                            $scope.c = $scope.object[$scope.field];
                            $scope.s = false;
                            $scope.r = false; // dirty hack to remove required from actual inputs once populated due to timing bug
                        } else {
                            $scope.c = $scope.v; // other value on the choice field
                            $scope.s = true;
                            update();
                        }
                    });
                }
            }
        });
})();
