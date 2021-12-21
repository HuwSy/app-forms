(function () {
	var appMod = angular.module('enter', []);
	
    appMod.directive('enter',
        function () {
            /// <summary>On enter key trigger specified action, such as begin a search when pressing enter on an input field</summary>
            return function (scope, element, attrs) {
				element.bind("keydown keypress", function (event) {
					if(event.which === 13) {
						scope.$apply(function (){
							scope.$eval(attrs.enter);
						});

						event.preventDefault();
					}
				});
			}
		});
})();