"use strict";

(function ($) {
	var appMod = angular.module('Home', ['appforms']);
    
    appMod.controller("Home", function ($scope, $interval, listSvc, webSvc, $routeParams, $uibModal, $timeout, $location, $rootScope,
		listName, beforeLoad, filter, select, order, afterLoad, rows, search, refresh, saveFilter) {
		$scope.routeParams = $routeParams || {};
		$scope.go = function (g) {
			$rootScope.$apply(function() {
				$location.path(g);
			});
		}
		$scope.web = _spPageContextInfo.webServerRelativeUrl;
		$scope.userId = _spPageContextInfo.userId;

		if (beforeLoad)
			beforeLoad($scope, webSvc, listSvc, $uibModal, $timeout);

		// pagination
		$scope.Filter = '';
		$scope.currentPage = 1;
		$scope.itemsPerPage = '25';
		$scope.changePage = function (to) {
			$scope.currentPage = Math.ceil(to);
		}
		
		// decide if the current user is an admin before loading
		$scope.permissions = Override.Permissions;

		// on typing in search box or filter changes
		$scope.search = function (r) {
			if (search)
				return search($scope, r);
			else
				return true;
		}
		
		// on clicking headings
		$scope.orderKey = null;
		$scope.orderDir = false;
		$scope.sort = function (k) {
			if ($scope.orderKey == k)
				$scope.orderDir = !$scope.orderDir;
			else {
				$scope.orderKey = k;
				$scope.orderDir = false;
			}
			$scope.saveFilter('orderKey');
			$scope.saveFilter('orderDir');
		}
		
		// save specific filter field
		$scope.saveFilter = function (f) {
			var cur = JSON.parse(localStorage.getItem('Filter-' + document.location.href) || '{}');
			cur[f] = $scope[f];
			localStorage.setItem('Filter-' + document.location.href, JSON.stringify(cur));
		}
		
		// load data from lists
		$scope.loadData = function (restart) {
			var cur = saveFilter ? JSON.parse(localStorage.getItem('Filter-' + document.location.href) || '{}') : null;
			for(var f in cur) {
				$scope[f] = cur[f];
			}

			if (restart === true) {
				$scope.Loading = true;
				$scope.currentPage = 1;
			}
			listSvc.getItems(typeof listName == "function" ? listName($scope) : listName, typeof filter == "function" ? filter($scope) : filter, typeof select == "function" ? select($scope) : select, null, null, typeof order == "string" ? order : 'Id', 5000)
				.then(function (d) {
					if (rows)
						d.d.results.forEach(rows);
					
					if (afterLoad)
						afterLoad(d, $scope);
					          
					try {
						if (d.d.results.length > 0)
							$scope.Export = _spPageContextInfo.webServerRelativeUrl.replace(/\/$/,'')
								+ "/_vti_bin/owssvr.dll?CS=65001&Using=_layouts/15/query.iqy&List=%7B"
								+ d.d.results[0].__metadata.uri.match(/guid'[^']*/)[0].substring(5)
								+ "%7D&CacheControl=1";
					} catch (e) {}

					if (typeof order == "function")
						d.d.results.sort(order);
                    
					$scope.rows = d.d.results;
					$scope.Loading = false;
					$scope.$apply();
				});
		}
		
		$scope.Loading = true;
		if (listName)
			$scope.loadData();
		// stop reloading home.js loop
		if (typeof app.interval != "undefined")
			clearInterval(app.interval);
		if (refresh)
			app.interval = $interval($scope.loadData, refresh === true ? 15000 : refresh).$$intervalId;
	});
})(APP$ || $);