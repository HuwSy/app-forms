(function () {
    var appMod = angular.module('error', []);
    
    appMod.directive('error',
        function () {
            function toHide (element) {
                var e = !element ? null : !element.length ? element : element.length > 0 ? element[0] : null;
                if (!e || !e.tagName)
                    return false;
                
                e.onerror = function () {
                    if ((this.getAttribute('error') || '') != '') {
                        this.src = this.getAttribute('error');
                        this.style.opacity = "0.25";
                    } else
                        this.style.visibility = "hidden";
                };

                return true;
            }

            /// <summary>Hides via visibility (to not distort the view) the img item if the image doesnt load, i.e. 404</summary>
            return function (scope, element, attrs) {
                    if (!toHide(element))
                        setTimeout(function () {
                            toHide(element)
                        }, 1);
                }
        });
})();