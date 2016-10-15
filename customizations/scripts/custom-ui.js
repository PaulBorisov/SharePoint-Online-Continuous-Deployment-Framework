(function () {
  var allSettings = {
    AlternateCssUrl: "/_catalogs/masterpage/customizations/css/custom-ui.css",
    Verbose: true,
    JqueryVersion: "3.1.0",
    UrlRedirects: {
      SearchResults: {Source: "_layouts/15/osssearchresults.aspx", Target: "/search/pages/results.aspx"}
    }
    ,Header: {
      Enabled: true,
      Parent: "#DeltaPlaceHolderMain", // In a format of a JQuery selector
      CssClass: "content-header-image"
    }
    ,Footer: {
      Enabled: true,
      // Use the extension .aspx instead of .html. This forces loading the content instead of downloading .html.
      // You can either specify a leading slash to load the footer from the top root site (i.e. use "/_catalogs/...")
      // or omit a leading slash to load footer from the current web (i.e. use "_catalogs/...").
      Url: "/_catalogs/masterpage/customizations/templates/footer.aspx",
      Width: "100%",
      Height: "175px",
      Scrolling: "no",
      // Sets the initial styles to iframe. Note this can be overwritten in the scripts' logic. For example, 
      // a script can change fixed position on window resize to relative and back (see "EnsureFooter" below).
      Style: "position:relative;left:0px;bottom:0px;"
    }
  };

  var allIntervals = {
    AlterDefaultBranding: null,
    EnsureRedirects: null,
    VerifyStatuses: null
  };

  var allStatuses = {
    AlterDefaultBrandingSucceded: false,
    EnsureRedirects: false,
    CurrentVerificationAttempts: 0,
    MaxVerificationAttempts: 20
  }

  function getQueryStringParameterByName(name, url) {
    if (!url) url = window.location.href;
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
        results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }

  function successCommon(message, interval) {
    if( message && allSettings.Verbose ) {
      console.log(message);
    }
    if( interval != null ) {
      //console.log('Removing interval #' + interval);
      clearInterval(interval);
      //console.log('Done.');
    }
  }
  function errorCommon(message) {
    if (arguments != null && arguments.callee != null && 
        arguments.callee.caller != null && arguments.callee.caller.name != null) {
      successCommon("ERROR in " + arguments.callee.caller.name + ": " + message, null);
    } else {
      successCommon("ERROR: " + message, null);
    }
  }

  function EnsureFooter() {
    if( typeof allSettings["Footer"] != "object" || !allSettings["Footer"].Enabled ) return;
    
    function AppendFooter(siteUrl, footer) {
      if( typeof footer != "object" || typeof footer.Url == "undefined" ) {
        return;
      }

      var el = jQuery('<iframe id="footer"></iframe>');
      if( typeof footer.Width != "undefined" ) {
        el.attr("width", footer.Width);
      }
      if( typeof footer.Height != "undefined" ) {
        el.attr("height", footer.Height);
      }
      if( typeof footer.Scrolling != "undefined" ) {
        el.attr("scrolling", footer.Scrolling);
      }
      if( typeof footer.Style != "undefined" ) {
        el.attr("style", footer.Style);
      }
      
      jQuery(document).ready(function () {
        jQuery(window).bind("resize", function(){
          //console.log("Resize executed...");
          function AdjustPosition() {
            var currentHeight = jQuery(elementSelector).height();
            var currentWidth = jQuery(elementSelector).width();
            var hasChanged = false;
            //console.log("i=" + i);
            if( initHeight == currentHeight && initWidth == currentWidth ) {
              //console.log("initHeight=" + initHeight);
              //console.log("currentHeight=" + currentHeight);
              //console.log("Equal");
            } else {
              //console.log("Changed");
              hasChanged = true;
            }
            if( hasChanged ) {
              clearInterval(intervalResize);
              clearTimeout(timeoutResize);
            }
            if( hasChanged || i == 0 ) {
              i++;
              //SP.UI.Notify.addNotification(jQuery(elementSelector).height(), true);
              //SP.UI.Notify.addNotification(jQuery(elementSelector).prop("scrollHeight"), true);
              //SP.UI.Notify.addNotification(jQuery(window).height(), true);
              //SP.UI.Notify.addNotification(jQuery("#footer").offset().top, true);
              if( jQuery(elementSelector).height() >= jQuery(elementSelector).prop("scrollHeight") ) {
                el.css("position", "fixed");
                el.css("left", "0px");
                el.css("bottom", "0px");
                //SP.UI.Notify.addNotification("P1", true);
              } else {
                el.css("position", "relative");
                el.css("left", "0px");
                el.css("bottom", "0px");
                //SP.UI.Notify.addNotification("P2", true);
              }
            } else {
              i++;
            }
          }

          var elementSelector = "#s4-workspace";
          var initHeight = jQuery(elementSelector).height();
          var initWidth = jQuery(elementSelector).width();
          var i = 0;
          var intervalResize = setInterval(AdjustPosition, 10);
          var timeoutResize = setTimeout(function(){
            //console.log("Timeout...");
            clearInterval(intervalResize);
          }, 1000);
        });
      });
      var src = siteUrl.replace(/\/+$/,'') + '/' + footer.Url.replace(/^\/+/,'');
      el.attr("src", src);
      jQuery("#s4-workspace").append(el);
    }

    var footer = allSettings.Footer;
    jQuery(document).ready(function () {
      if( footer.Url.match(/^\//) ) {  // If URL starts with slash.
        AppendFooter("", footer);
      } else {
        FindCurrentWebUrlAndExecute(function(siteUrl) {
          AppendFooter(siteUrl, footer);
        });
      }
    });
  }

  function EnsureHeader() {
    if( typeof allSettings.Header != "object" || !allSettings.Header.Enabled ) return;
    // Avoid displaying a header on the system UIs.
    if( window.location.href.match(/((\/_layouts)|(\/_catalogs)|(\/_vti_))/i) != null ) return;
    jQuery(document).ready(function () {
      var parentElement = allSettings.Header.Parent ? allSettings.Header.Parent : "#DeltaPlaceHolderMain";
      var cssClass = 
        allSettings.Header.CssClass ? allSettings.Header.CssClass.replace(/^\.+/,'') : "content-header-image";
      jQuery(parentElement).prepend('<div class="' + cssClass + '"></div>');
      jQuery("#s4-workspace").scrollTop(0);
    });
  }
  
  function FindCurrentWebUrlAndExecute(callbackFunction) {
    try {
      jQuery(document).ready(function () {
        var loaded = false;
        var siteUrlToken = "/_vti_bin/spsdisco.aspx";
        var selector = "link[href$='" + siteUrlToken + "']";
        var siteUrl = jQuery(selector).attr('href');
        if( siteUrl ) {
          var slashIndex = siteUrl.indexOf(siteUrlToken);
          if( slashIndex > -1 ) {
            loaded = true;
            if( slashIndex > 0 ) {
              siteUrl = siteUrl.substr(0, slashIndex);
            } else {
              siteUrl = "/";
            }
          }
        }

        if( loaded ) {
          callbackFunction(siteUrl);
        } else {
          SP.SOD.executeFunc('sp.js', 'SP.ClientContext', function () {
            var ctx = new SP.ClientContext.get_current();
            var web = ctx.get_web();
            ctx.load(web);
            ctx.executeQueryAsync(successWebLoad, errorWebLoad);

            function successWebLoad() {
              try {
                siteUrl = web.get_serverRelativeUrl();
                callbackFunction(siteUrl);
              } catch (e) {
                errorCommon("FindCurrentWebUrlAndExecute, calling successWebLoad: " + e.message);
              }
            }
            function errorWebLoad(sender, args) {
              errorCommon("FindCurrentWebUrlAndExecute, loading web: " + args.get_message());
            }
          });
        }
      });
    } catch(e) {
      errorCommon("FindCurrentWebUrlAndExecute, using jQuery: " + e.message);
    }
  }
  
  function AlterDefaultBranding() {
    var interval = allIntervals.AlterDefaultBranding;
    var alternateCssUrl = allSettings.AlternateCssUrl.replace(/^\/+/, '');

    try {
      jQuery(document).ready(function () {
        if( jQuery(".custom-ms-breacrumb-top").length == 0 ) {
          var customBreadCrumb = jQuery('<div class="custom-ms-breacrumb-top"></div>');
          jQuery("#s4-titlerow").after(customBreadCrumb);
          var breadCrumb = jQuery(".ms-breadcrumb-top");
          breadCrumb.detach();
          customBreadCrumb.append(breadCrumb);
          customBreadCrumb.removeClass("s4-notdlg").addClass("s4-notdlg");
        }
        
        var selector = "link[href$='" + alternateCssUrl + "']";
        var loadCss = jQuery(selector).length == 0;

        if( loadCss ) {
          if( allSettings.AlternateCssUrl.match(/^\//) ) {
            var cssLink = '<link rel="stylesheet" type="text/css" href="' 
              + allSettings.AlternateCssUrl + '"/>';
            jQuery("head:first").append(cssLink);
            successCommon(null, interval)
            allStatuses.AlterDefaultBrandingSucceded = true;
          } else {
            FindCurrentWebUrlAndExecute(function(siteUrl) {
              function Callback(siteUrl, alternateCssUrl, interval) {
                try {
                  var cssLink = '<link rel="stylesheet" type="text/css" href="' 
                    + siteUrl.replace(/\/+$/,'') + '/' + alternateCssUrl + '"/>';
                  jQuery("head:first").append(cssLink);
                  successCommon(null, interval)
                  allStatuses.AlterDefaultBrandingSucceded = true;
                } catch (e) {
                  errorCommon("AlterDefaultBranding, executing Callback: " + e.message);
                }
              }
              Callback(siteUrl, alternateCssUrl, interval);
            });
          }
        } else {
          successCommon(null, interval)
          allStatuses.AlterDefaultBrandingSucceded = true;
        }
      });
    } catch(e) {
      errorCommon("AlterDefaultBranding, attempt to use jQuery: " + e.message);
    }
  }

  function EnsureRedirects() {
    var interval = allIntervals.EnsureRedirects;
    var url = window.location.href;
    for( var key in allSettings.UrlRedirects ) {
      var source = allSettings.UrlRedirects[key].Source;
      var urlMatch = new RegExp(source, "i");
      if( url.match(urlMatch) ) {
        successCommon(null, interval); 
        var newUrl = allSettings.UrlRedirects[key].Target;
        if( newUrl.match(/^\//) ) {
          newUrl = window.location.protocol + "//" + window.location.host + newUrl 
            + (newUrl.indexOf('?') == -1 ? window.location.search : window.location.search.replace(/\?/g, '&'));
        } else if( newUrl.match(/\:\/\//) ) {
          newUrl += (newUrl.indexOf('?') == -1 ? window.location.search : window.location.search.replace(/\?/g, '&'));
        } else {
          newUrl = url.replace(urlMatch, allSettings.UrlRedirects[key].Target);
        }
        window.location.href = newUrl;
        break;
      }
    }
    allStatuses.EnsureRedirects = true;
    successCommon(null, interval); 
  }
  
  function VerifyStatuses() {
    var interval = allIntervals.VerifyStatuses;
    allStatuses.CurrentVerificationAttempts++;
    if( allStatuses.AlterDefaultBrandingSucceded ) {
      successCommon("Branding completed after " + allStatuses.CurrentVerificationAttempts + " verification attempt(s).", interval);
      for( var key in allIntervals ) {
        successCommon(null, allIntervals[key]);
      }
    } else {
      if (allStatuses.CurrentVerificationAttempts >= allStatuses.MaxVerificationAttempts) {
        successCommon("Branding was not completed after " + allStatuses.CurrentVerificationAttempts
          + " verification attempt(s). Refresh the page if you want to repeat it.", interval);
        for( var key in allIntervals ) {
          successCommon(null, allIntervals[key]);
        }
      }
    }
  }
  
  // Load jQuery from CDN's URL (the code executes only in the case when it was not loaded by the custom action).
  (window.jQuery ||
   document.write('<script src="//ajax.aspnetcdn.com/ajax/jquery/jquery-' + allSettings.JqueryVersion + '.min.js"><\/script>'));

  var revert = getQueryStringParameterByName("nobranding");
  if (revert == null || revert.match(/((1)|(true))/i) == null) {
    // setInterval is used to guarantee execution of the code inside SP.SOD.executeFunc, which execution can delay due to a long loading time.
    allIntervals.AlterDefaultBranding = setInterval(function () { AlterDefaultBranding(); }, 100);
    allIntervals.EnsureRedirects = setInterval(function () { EnsureRedirects(); }, 100);
    allIntervals.VerifyStatuses = setInterval(function () { VerifyStatuses(); }, 1000);
    EnsureHeader();
    EnsureFooter();
  }
})();