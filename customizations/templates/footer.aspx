<script src="https://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js" type="text/javascript" async=""></script>
<script>
  WebFontConfig = {
    google: {families: ['Open+Sans:400,700:latin']}
  };
  (function() {
    var wf = document.createElement('script');
    wf.src = ('https:' == document.location.protocol ? 'https' : 'http') + '://ajax.googleapis.com/ajax/libs/webfont/1/webfont.js';
    wf.type = 'text/javascript';
    wf.async = 'true';
    var s = document.getElementsByTagName('script')[0];
    s.parentNode.insertBefore(wf, s);
  })();

  if (navigator.userAgent.match(/IEMobile\/10\.0/)) {
    var msViewportStyle = document.createElement('style');
    msViewportStyle.appendChild(
      document.createTextNode(
        '@-ms-viewport{width:auto!important}'
      )
    );
    document.getElementsByTagName('head')[0].appendChild(msViewportStyle);
  }
</script>
<link rel="stylesheet" media="not print" href="/_catalogs/masterpage/customizations/css/style.css">
<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Open+Sans:400,700&amp;subset=latin" media="all">
<footer role="contentinfo">
  <div class="wrapper">
      <div class="clearfix">
          <ul class="extras col small_4 large_7 clearfix">
              <li>
                  <a href="http://www.contoso.org/about-us">About us</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/business">Business</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/customers">Customers</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/news-and-events">News and Events</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/contact">Contacts</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/site-map">Site Map</a>
              </li>
          </ul>
          <ul class="social col small_4 large_5 large_switch no-external clearfix">
              <li>
                  <a href="http://www.facebook.com/spo-cdf" class="icon-facebook-rect">
                      <span class="away">Facebook</span>
                  <span class="away"> (external link)</span></a>
              </li>
              <li>
                  <a href="http://twitter.com/spo-cdf" class="icon-twitter-bird">
                      <span class="away">Twitter</span>
                  <span class="away"> (external link)</span></a>
              </li>
              <li>
                  <a href="http://www.youtube.com/spo-cdf" class="icon-youtube">
                      <span class="away">YouTube</span>
                  <span class="away"> (external link)</span></a>
              </li>
              <li>
                  <a href="http://www.flickr.com/photos/spo-cdf" class="icon-flickr-circled">
                      <span class="away">Flickr</span>
                  <span class="away"> (external link)</span></a>
              </li>
              <li>
                  <a href="http://www.linkedin.com/company/contoso" class="icon-linkedin-rect">
                      <span class="away">LinkedIn</span>
                  <span class="away"> (external link)</span></a>
              </li>
              <li>
                  <a href="http://www.pinterest.com/spo-cdf/" class="icon-pinterest">
                      <span class="away">Pinterest</span>
                  <span class="away"> (external link)</span></a>
              </li>
          </ul>
      </div>
      <div class="clearfix">
          <ul class="small-print col small_4 large_9">
              <li>
                  <a href="http://www.contoso.org/accessibility">Accessibility</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/languages">Languages</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/cookies">Cookies</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/terms">Terms and disclaimer</a>
              </li>
              <li>
                  <a href="http://www.contoso.org/privacy">Privacy</a>
              </li>
              <li>© The Contoso Corporation Ltd.</li>
          </ul>
          <a href="http://www.contoso.org" class="tab-logo col small_4 large_3 no-bottom">
            <img src="/_catalogs/masterpage/customizations/css/images/logo.png" />
          </a>
      </div>
  </div>
</footer>
<script type="text/javascript">
  if( typeof window.parent != "undefined" ) {
    function parentRedirect(url) {
      if( typeof window.parent != "undefined" ) {
        var message = '<div style="color:#810b19;text-align:left;font-weight:bold;">Your request is being virtually redirected to ' + url + '</div>';
        window.parent.SP.UI.Notify.addNotification(message, false);
        //window.parent.location.href = url;
      }
    }
    
    var allLinks = document.getElementsByTagName("a");
    for( var i = 0; i < allLinks.length; i++ ) {
      var link = allLinks[i];
      if( link.href.match(/^http/i) || link.href.match(/^\//) ) {
        (function(){
          var url = link.href;
          if( link.addEventListener ) { // For all major browsers, except IE 8 and earlier
            link.addEventListener("click", function(){parentRedirect(url);});
          } else if (link.attachEvent) { // For IE 8 and earlier versions
            link.attachEvent("onclick", function(){parentRedirect(url);});
          }
          link.href = "#";
        })();
      }
    }
  }
</script>
