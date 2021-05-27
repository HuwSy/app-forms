"use strict";
(function () {
	var appMod = angular.module('upload', []);
	
    appMod.directive('upload',
        function () {
            /// <summary>Upload file directive with multiple support, list of current files, deletion commands etc</summary>
            return {
                restrict: 'E',
                scope: {
					attachments: "=",
					prefix: "=",
                    type: "=",
					mult: "=",
					addOnly: '=',
					web: "=",
					callback: "=",
					ngDisabled: "=",
					ngRequired: "=" // doesn't reject if invalid only highlight
                },
				template: '\
<ul style="padding-left: 20px; list-style: none;">\
	<li ng-show="attachments == null || (attachments | filter:search).length == 0" style="padding: 4px;" ng-class="{\'ng-invalid\':ngRequired}">No attachments</li>\
	<li ng-repeat="a in (attachments | filter:search)" style="padding: 4px; {{a.Deleted ? \'text-decoration: line-through\' : \'\'}}"><a {{web === false ? "download=\"" + a.FileName + "\"" : ""}} href="{{a.ServerRelativeUrl || \'#\'}}{{(web ? \'?Web=1\' : \'\')}}" target="_blank">{{a.FileName}}</a> <i class="fa fa-remove" ng-click="deleteLogo(a)" style="cursor: pointer;" ng-hide="ngDisabled || a.Deleted || (addOnly && a.ServerRelativeUrl)"></i> <i class="fa fa-upload" ng-show="a.ServerRelativeUrl == null"></i></li>\
</ul>\
<input ng-hide="ngDisabled" accept="{{type}}" type="file" style="display: none;" multiple onchange="angular.element(this).scope().addLogo(this)" />\
<a class="print-hide" style="color: black;padding: 7px 10px;border: 1px solid #ababab;background-color: #fdfdfd;font-size: 11px;float: right;" ng-hide="ngDisabled" onclick="this.previousElementSibling.tagName == \'DIV\' ? this.previousElementSibling.previousElementSibling.click() : this.previousElementSibling.click()">Add Files</a>\
				',
                controller: Upload
            }
        });

    function Upload($scope) {
		$scope.search = function (r) {
			return $scope.prefix == null
				|| $scope.prefix == ''
				|| r.FileName.toLowerCase().indexOf($scope.prefix.toLowerCase()+'-') == 0
		}

        $scope.deleteLogo = function (n) {
            /// <summary>Delete the selected uploaded object, or mark for pending server deletion</summary>
			if (n.ServerRelativeUrl != null) {
				n.Deleted = true;
			} else {
				var a = [];
                ($scope.attachments || []).forEach(function (f) {
					if (f.ServerRelativeUrl != null || f.FileName.toLowerCase() != n.FileName.toLowerCase())
						a.push(f);
				});
				$scope.attachments = a;
				if (typeof $scope.callback == "function")
					$scope.callback();
			}
		}
		
        $scope.addLogo = function (ths, txt) {
            /// <summary>Add new file data to the object fur later upload</summary>
			var files = [];
			// ensure the file name doenst already exist as duplicaes are not allowed
			for (var f = 0; f < ths.files.length; f++) {
				var dup = false;
				if ($scope.mult) {
					var n = ($scope.prefix && $scope.prefix != '' ? $scope.prefix.toLowerCase() + '-' : '');
					if ($scope.prefix && $scope.prefix != '' && ths.files[f].name.toLowerCase().indexOf($scope.prefix.toLowerCase()+'-') == 0)
						n += ths.files[f].name.toLowerCase().substring($scope.prefix.length + 1).replace(/[%'#]/g,'-');
					else
						n += ths.files[f].name.toLowerCase().replace(/[%'#]/g,'-');
					
					var dup = false;
					($scope.attachments || []).forEach(function (a) {
						if (!dup && !a.Deleted && a.FileName.toLowerCase() == n) {
							alert('File with name "' + a.FileName + '" already exists.');
							dup = true;
						}
					});
				}
				if (!dup)
			        files.push(ths.files[f]);
			}
			// load file(s) into js objects ready for save action
			ths.value = null;
			files.forEach(function (f) {
				var reader = new window.FileReader();
				reader.onload = function (event) {
					var data = '';
					if (!txt) {
						var bytes = new window.Uint8Array(event.target.result);
						var len = bytes.byteLength;
						for (var i = 0; i < len; i++) {
							data += String.fromCharCode(bytes[i]);
						}
					} else {
						data = event.target.result;
						len = data.length;
					}

					if ($scope.attachments == null)
					    $scope.attachments = [];

                    if (!$scope.mult)
                        for (var a = $scope.attachments.length - 1; a >= 0; a--)
                            $scope.deleteLogo($scope.attachments[a]);

					var n = ($scope.prefix && $scope.prefix != '' ? $scope.prefix + '-' : '');
					if ($scope.prefix && $scope.prefix != '' && f.name.toLowerCase().indexOf($scope.prefix.toLowerCase()+'-') == 0)
						n += f.name.substring($scope.prefix.length + 1).replace(/[%'#]/g,'-');
					else
						n += f.name.replace(/[%'#]/g,'-');
					
					$scope.attachments.push({
						FileName: n,
						ServerRelativeUrl: null,
						Data: data,
						Length: len
					});

					$scope.$apply();
					if (typeof $scope.callback == "function")
						$scope.callback();
				};
				reader.onerror = function () {
					alert("File reading error " + f.name);
				};
				if (!txt)
					reader.readAsArrayBuffer(f);
				else
					reader.readAsText(f);
			});
		}

		$scope.csvToJson = function (csv, header) {
			var lines = csv.split("\n");
			var result = [];
			var headers = [];

			if (lines.length <= 0)
				return null;

			if (header)
				headers = lines[0].split(",");
			else {
				var count = lines[0].split(",").length;
				for (var x = 0; x < count; x++)
					headers.push(x);
			}
			
			for (var i = (header ? 1 : 0); i < lines.length; i++) {
				var obj = {};
				var currentline = lines[i].split(",");
			
				for (var j = 0; j < headers.length; j++) {
					obj[headers[j]] = currentline[j];
				}
			
				result.push(obj);
			}
			
			return JSON.stringify(result);
		}
	}
})();
