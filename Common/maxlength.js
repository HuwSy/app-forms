(function () {
    var appMod = angular.module('maxlength', []);
    
    appMod.directive('maxlength',
        function () {
            /// <summary>Add counter to maxlength</summary>
            return function (scope, element, attrs) {
                var l = element[0].getAttribute('maxlength');
                if (l == null || l == '')
                    return;
                
                var i = parseInt(l);
                if (i <= 0)
                    return;
                
				var t = null;
				element.bind('keyup', function (event) {
					if (element[0].value.length > i/2) {
                        var r = (i - element[0].value.length);
                        if (r < 0)
                            r = 0;

                        if (t)
                            return t.innerText = r + " characters remaining";
                        
                        t = document.createElement('div');
                        t.style.position = 'absolute';
                        t.style.right = '18px';
                        t.style.marginTop = '25px';
                        t.style.fontSize = '10px';
                        t.style.color = 'red';
                        
                        t.innerText = r + " characters remaining";
                        element.after(t)
                    } else if (t) {
                        t.remove();
                        t = null;
                    }
				});
			}
		});
})();