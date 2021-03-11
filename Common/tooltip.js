(function () {
	var appMod = angular.module('tooltip', []);
	
    appMod.directive('tooltip',
        function () {
            /// <summary>On mouse over any element with tooltip</summary>
            return function (scope, element, attrs) {
				var t = null;
				element.bind('mouseover', function (event) {
					//if (t != null)
					//	return t.style = t.style.replace(/\;top\:[^\;]*\;/, ';top:' + (event.clientY + 2) + 'px;').replace(/\;left\:[^\;]*\;/, ';left:' + (event.clientX + 2) + 'px;');
					
					var text = element[0].getAttribute('tooltip');
					if (text == null || text == '')
						return;

					if (t == null)
						t = document.getElementById('ToolTip');
					if (t == null) {
						var c = document.createElement('div');
						c.id = 'ToolTip';
						c.style = '\
							position:absolute;\
							z-index:10000;\
							background:white;\
							padding:4px;\
							border:1px solid #ddd;\
							box-shadow:2px 2px 2px #ddd;\
							max-width:' + (text.length > 1000 ? 650 : text.length > 500 ? 500 : 350) + 'px;\
							top:-20px;\
							left:-20px;\
						';
						document.body.appendChild(c);
						t = document.getElementById('ToolTip');
					}

					t.innerHTML = text.replace(/\\n/g, '<br><br>').replace(/\n/g, '<br><br>');

					var y = event.clientY + 2;
					if (y + t.offsetHeight > window.innerHeight)
						y = window.innerHeight - t.offsetHeight;
					if (y < 0)
						y = 0;
					t.style.top = y + 'px';

					var x = event.clientX + 2;
					if (x + t.offsetWidth > window.innerWidth)
						x = event.clientX - 2 - t.offsetWidth;
					if (x < 0)
						x = 0;
					t.style.left = x + 'px';
				});
				element.bind('mouseout', function (event) {
					if (t != null) {
						t.style.display = "none";
						t.remove();
					}
					t = null;
				});
				element.bind('click', function (event) {
					if (t != null) {
						t.style.display = "none";
						t.remove();
					}
					t = null;
				});
			}
		});
})();