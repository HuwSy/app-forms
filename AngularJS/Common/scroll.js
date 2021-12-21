(function () {
	var appMod = angular.module('scroll', []);
	
    appMod.directive('scroll',
        function () {
            /// <summary>Scrolls to the element or the element with this attribute specified on a page</summary>
			return function(scope, element, attrs) {
				scope.$watch(attrs.scroll, function(value) {
					if (value) {
						setTimeout(function () {
							element[0].scrollIntoView({block: "end", behavior: "smooth"});
						}, value > 0 ? value : 1);
					}
				});
			}
		});
})();