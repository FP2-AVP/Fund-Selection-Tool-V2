/* Auto-inject additional stylesheets */
(function() {
  var sheets = [
    'css/other-factors.css',
  ];
  sheets.forEach(function(href) {
    if (document.querySelector('link[href="' + href + '"]')) return;
    var link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = href;
    document.head.appendChild(link);
  });
})();
