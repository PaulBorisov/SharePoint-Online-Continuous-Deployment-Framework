(function(){ // Use this logic to customize actions related to asynchronous creation of a new site collection. 
  var allSettings = {
    DialogTitleDefault: "Create site collection",
    DialogUrl: "_layouts/15/newsbweb.aspx",
    JqueryVersion: "3.1.0",
    ListForDeploymentRequests: {
      Title:"Deployment Requests", 
      Url:"Lists/DeploymentRequests",
      ContentTypes: {
        "Site Collection": {
          Fields: [
            "<Field Type='Choice' DisplayName='Managed Path' Name='ManagedPath' FillInChoice='FALSE' Format='Dropdown'>" +
              "<Default>/sites/</Default><CHOICES><CHOICE>/sites/</CHOICE><CHOICE>/teams/</CHOICE></CHOICES></Field>",
            "<Field Type='Text' DisplayName='Site URL' Name='SiteUrl' />",
            "<Field Type='Text' DisplayName='Template' Name='WebTemplate' />",
            "<Field Type='Text' DisplayName='Description' Name='Description' />",
            "<Field Type='Choice' DisplayName='Language' Name='Language' FillInChoice='TRUE' Format='Dropdown'>" +
              "<Default>1033</Default><CHOICES><CHOICE>1033</CHOICE><CHOICE>1035</CHOICE><CHOICE>1049</CHOICE>" +
              "<CHOICE>1053</CHOICE></CHOICES></Field>",
            //"<Field Type='Text' DisplayName='Requestor Email' Name='RequestorEmail' />",
            //"<Field Type='Text' DisplayName='Notify Email' Name='NotifyEmail' />",
            "<Field Type='Choice' DisplayName='Status' Name='Status' FillInChoice='FALSE' Format='Dropdown'>" + 
              "<Default>Requested</Default><CHOICES><CHOICE>Requested</CHOICE><CHOICE>Created</CHOICE>" + 
              "<CHOICE>CreatedNeedsCustomTemplate</CHOICE><CHOICE>CreatedCustomTemplateFailed</CHOICE>" + 
              "<CHOICE>CreatedCustomTemplateApplied</CHOICE><CHOICE>Deleted</CHOICE><CHOICE>Failed</CHOICE>" + 
              "<CHOICE>Creating</CHOICE><CHOICE>Deleting</CHOICE></CHOICES></Field>",
            "<Field Type='Text' DisplayName='Status Message' Name='StatusMessage' />",
            "<Field Type='Text' DisplayName='Updated Status Message' Name='UpdatedStatusMessage' />"
          ]
        }
      }
    },
    ScriptName: "custom-sa-create-site.js",
    WebTemplates:{
      // How to load jQuery to the browser using Chrome's console:
      // var el = document.createElement("script");el.type = "text/javascript";el.src = "//ajax.aspnetcdn.com/ajax/jquery/jquery-3.1.0.min.js";document.getElementsByTagName("head")[0].appendChild(el);
      // How to get site templaces from the currently selected group in the SPO Admin Center using Chrome's console:
      // var items={};$("select[name$='LbWebTemplate'] > option").each(function(){var a=$(this);items[a.text()]=a.val();});window.JSON.stringify(items);
      // var items={};$("select[name$='LbWebTemplate'] > option").each(function(i){var a=$(this);items[a.val()]={Title:a.text(),Description:$("#HidDescription"+i).val()};});window.JSON.stringify(items);
      Default:1033
      ,1033:{
        Collaboration: {"STS#0":{"Title":"Team Site","Description":"A place to work together with a group of people."},"BLOG#0":{"Title":"Blog","Description":"A site for a person or team to post ideas, observations, and expertise that site visitors can comment on."},"DEV#0":{"Title":"Developer Site","Description":"A site for developers to build, test and publish apps for Office"},"PROJECTSITE#0":{"Title":"Project Site","Description":"A site for managing and collaborating on a project. This site template brings all status, communication, and artifacts relevant to the project into one place."},"COMMUNITY#0":{"Title":"Community Site","Description":"A place where community members discuss topics of common interest. Members can browse and discover relevant content by exploring categories, sorting discussions by popularity or by viewing only posts that have a best reply. Members gain reputation points by participating in the community, such as starting discussions and replying to them, liking posts and specifying best replies."}},
        Enterprise: {"BDR#0":{"Title":"Document Center","Description":"A site to centrally manage documents in your enterprise."},"EDISC#0":{"Title":"eDiscovery Center","Description":"A site to manage the preservation, search, and export of content for legal matters and investigations."},"OFFILE#1":{"Title":"Records Center","Description":"This template creates a site designed for records management. Records managers can configure the routing table to direct incoming files to specific locations. The site also lets you manage whether records can be deleted or modified after they are added to the repository."},"EHS#1":{"Title":"Team Site - SharePoint Online configuration","Description":"A Team Site configured to allow organization members to edit, create new sites, and share with external users."},"BICenterSite#0":{"Title":"Business Intelligence Center","Description":"A site for presenting Business Intelligence content in SharePoint."},"POLICYCTR#0":{"Title":"Compliance Policy Center","Description":"Use the Document Deletion Policy Center to manage policies that can delete documents after a specified period of time. These policies can then be assigned to specific site collections or to site collection templates."},"SRCHCEN#0":{"Title":"Enterprise Search Center","Description":"A site focused on delivering an enterprise-wide search experience. Includes a welcome page with a search box that connects users to four search results page experiences: one for general searches, one for people searches, one for conversation searches, and one for video searches. You can add and customize new results pages to focus on other types of search queries."},"SPSMSITEHOST#0":{"Title":"My Site Host","Description":"A site used for hosting personal sites (My Sites) and the public People Profile page. This template needs to be provisioned only once per User Profile Service Application, please consult the documentation for details."},"COMMUNITYPORTAL#0":{"Title":"Community Portal","Description":"A site for discovering communities."},"SRCHCENTERLITE#0":{"Title":"Basic Search Center","Description":"A site focused on delivering a basic search experience. Includes a welcome page with a search box that connects users to a search results page, and an advanced search page. This Search Center will not appear in navigation."},"visprus#0":{"Title":"Visio Process Repository","Description":"A site for viewing, sharing, and storing Visio process diagrams. It includes a versioned document library and templates for Basic Flowcharts, Cross-functional Flowcharts, and BPMN diagrams."}},
        Publishing: {"BLANKINTERNETCONTAINER#0":{"Title":"Publishing Portal","Description":"A starter site hierarchy for an Internet-facing site or a large intranet portal. This site can be customized easily with distinctive branding. It includes a home page, a sample press releases subsite, a Search Center, and a login page. Typically, this site has many more readers than contributors, and it is used to publish Web pages with approval workflows."},"ENTERWIKI#0":{"Title":"Enterprise Wiki","Description":"A site for publishing knowledge that you capture and want to share across the enterprise. It provides an easy content editing experience in a single location for co-authoring content, discussions, and project management."}},
        Custom: {"__SELECTLATER":{"Title":"< Select template later... >","Description":"Create an empty site and pick a template for the site at a later time."}},
        Legacy: {"SiteTemplateEng":{"Title":"Legacy site collection", "Description":""}}
      }
      ,1035:{
        Yhteiskäyttö: {"STS#0":{"Title":"Työryhmäsivusto","Description":"Paikka, jossa voi työskennellä muiden kanssa."},"BLOG#0":{"Title":"Blogi","Description":"Sivusto, johon henkilö tai ryhmä voi julkaista ideoita, havaintoja ja muita tietoja ja jossa muut voivat kommentoida näitä tietoja."},"DEV#0":{"Title":"Kehittäjäsivusto","Description":"Sivusto, jossa kehittäjät voivat luoda, testata ja julkaista Office-sovelluksia"},"PROJECTSITE#0":{"Title":"Projektisivusto","Description":"Sivusto projektin hallintaa ja yhteiskäyttöä varten. Tämä sivustomalli tuo kaikki projektiin liittyvät tilat ja viestinnän sekä tiedot yhteen paikkaan."},"COMMUNITY#0":{"Title":"Yhteisön sivusto","Description":"Paikka, jossa yhteisön jäsenet voivat keskustella yhteisistä mielenkiinnon kohteista. Jäsenet voivat selata ja etsiä sopivaa sisältöä tutustumalla luokkiin, järjestelemällä keskustelut suosion mukaan tai katsomalla vain kirjoitukset, joissa on paras vastaus. Jäsenet saavat mainepisteitä osallistumalla yhteisön toimintaan esimerkiksi aloittamalla keskusteluja ja vastaamalla niihin, tykkäämällä kirjoituksista ja määrittämällä parhaita vastauksia."}},
        Yritys: {"BDR#0":{"Title":"Tiedostokeskus","Description":"Sivusto, jossa yrityksesi tiedostoja voidaan hallita keskitetysti."},"EDISC#0":{"Title":"eDiscovery-keskus","Description":"Sivusto oikeudellisen ja tutkintaan liittyvän sisällön säilyttämiseen, hakemiseen ja vientiin."},"OFFILE#1":{"Title":"Tietuekeskus","Description":"Tämä malli luo tietueiden hallintaan tarkoitetun sivuston. Tietueiden ylläpitäjät voivat määrittää reititystaulukon, joka ohjaa saapuvat tiedostot tiettyihin sijainteihin. Sivustossa voi myös määrittää, voiko tietueita poistaa tai muokata sen jälkeen, kun ne on lisätty säilöön."},"EHS#1":{"Title":"Ryhmäsivusto - SharePoint Online -kokoonpano","Description":"Sivusto, jonka avulla käyttäjät voivat nopeasti luoda yhteistyötilan, joka sisältää julkisia verkkosivuja ja vain jäsenille tarkoitetun alueen."},"BICenterSite#0":{"Title":"Liiketoimintatietokeskus","Description":"Sivusto liiketoimintatietosisällön esittämiseen SharePointissa."},"POLICYCTR#0":{"Title":"Yhteensopivuuskäytäntökeskus","Description":"Tiedoston poistokäytännön keskuksessa voit hallita käytäntöjä, joilla tiedostoja poistetaan tietyn ajanjakson jälkeen. Nämä käytännöt voidaan liittää tiettyihin sivustokokoelmiin tai sivustokokoelman malleihin."},"SRCHCEN#0":{"Title":"Enterprise-hakukeskus","Description":"Sivusto koko yrityksen hakutoimintoja varten. Sisältää aloitussivun, jossa on hakuruutu, joka yhdistää käyttäjät neljään eri hakutulossivuun: yksi yleishauille, yksi henkilöhauille, yksi keskusteluhauille ja yksi videohauille. Voit lisätä ja mukauttaa uusia hakutulossivuja, jos haluat keskittyä muihin etsintäalueisiin."},"SPSMSITEHOST#0":{"Title":"Oman sivuston isäntä","Description":"Sivusto, jossa henkilökohtaiset (omat) sivustot ja julkinen henkilöprofiilisivu sijaitsevat. Tämä malli on valmisteltava vain kerran kullekin käyttäjäprofiilien palvelusovellukselle. Lisätietoja on dokumentaatiossa."},"COMMUNITYPORTAL#0":{"Title":"Yhteisöportaali","Description":"Yhteisöjen etsintään tarkoitettu sivusto."},"SRCHCENTERLITE#0":{"Title":"Perushakukeskus","Description":"Sivusto hakutoimintoja varten. Sisältää aloitussivun, jossa oleva hakuruutu yhdistää käyttäjät hakutulossivulle, sekä tarkennetun haun sivun."},"vispr#0":{"Title":"Vision prosessisäilö","Description":"Sivusto, jota voi käyttää Visio-prosessikaavioiden tarkasteluun, jakamiseen ja tallentamiseen. Sivustossa on versionhallintaa tukeva tiedostokirjasto sekä perusvuokaavioiden, toimintojen välisten vuokaavioiden ja BPMN-kaavioiden mallit."}},
        Julkaiseminen: {"BLANKINTERNETCONTAINER#0":{"Title":"Julkaisemisportaali","Description":"Internetiin yhteydessä olevan sivuston tai suuren intranet-portaalin aloitussivusto. Tätä sivustoa voidaan mukauttaa helposti lisäämällä yrityskuvaan liittyviä selkeästi tunnistettavia ominaisuuksia. Se sisältää kotisivun, lehdistötiedotteiden mallialisivuston, hakukeskuksen ja kirjautumissivun. Yleensä tällä sivustolla on paljon enemmän lukijoita kuin tekijöitä, ja sitä käytetään hyväksymisen työnkulkuja sisältävien verkkosivujen julkaisemiseen."},"ENTERWIKI#0":{"Title":"Yrityswiki","Description":"Sivusto, jossa voit julkaista löytämiäsi tietoja ja jakaa ne yrityksen muiden työntekijöiden kanssa. Sivusto toimii keskitettynä sisällönmuokkaussijaintina, jossa useat käyttäjät voivat käyttää sisältöä, keskusteluja ja projektinhallintaa samanaikaisesti."}},
        Mukautettu: {"__SELECTLATER":{"Title":"< Valitse malli myöhemmin... >","Description":"Luo tyhjä sivusto ja valitse siihen malli myöhemmin."}},
        Legacy: {"SiteTemplateFin":{"Title":"Vanha sivustokokoelma", "Description":""}}
      }
    },
    // Advanced web templates use dynamically generated js-scripts that contain advanced localization 
    // of templates' titles and descriptions in all languages available in Tenant. 
    // This functionality is potentially error prone and can be enabled or disabled by correspondent setting.
    WebTemplatesAdvanced:{ 
      Enabled:true,
      ScriptForCustomTemplates: "_catalogs/masterpage/customizations/scripts/wt-custom.js",
      ScriptForStandardTemplates: "_catalogs/masterpage/customizations/scripts/wt-standard.js"
    }
  };

  function AddDeploymentRequest(data) {
    var note = SP.UI.Notify.addNotification("Processing your request, please wait...", false);
    EnsureListForDeploymentRequests(function(list){AddDeploymentRequestCallback(list, data, note);});
  }
  
  function AddDeploymentRequestCallback(list, data, note) {
    var info = new SP.ListItemCreationInformation();
    var listItem = list.addItem(info);
    listItem.set_item("Title", data["Title"]);
    listItem.set_item("Description", data["Description"]);
    listItem.set_item("ManagedPath", data["ManagedPath"]);
    listItem.set_item("SiteUrl", data["SiteUrl"]);
    listItem.set_item("Language", data["Language"]);
    listItem.set_item("WebTemplate", data["WebTemplate"]);
    listItem.set_item("Status", "Requested");
    listItem.set_item("StatusMessage", "Request stored " + (new Date()).format("dd.MM.yyyy HH:mm:ss"));
    listItem.update();

    var ctx = SP.ClientContext.get_current();
    ctx.executeQueryAsync(function(){
      if( note ) {
        SP.UI.Notify.removeNotification(note);
      }
      var listUrl = list.get_rootFolder().get_serverRelativeUrl();
      console.log("Request was successfully added as a list item to " + listUrl);
      var strHtml = "<div style='text-align:left;color:darkgreen;font-weight:bold;'>" +
        "<div>Your request has been successfully added to the list '" + 
        "<a style='text-decoration:underline;' href='" +  listUrl + "' target='_blank'>" +
        allSettings.ListForDeploymentRequests.Title + "</a>'</div>" +
        "<div>Title: " + listItem.get_item("Title") + "</div>" +
        "<div>Description: " + listItem.get_item("Description") + "</div>" +
        "<div>Managed Path: " + listItem.get_item("ManagedPath") + "</div>" +
        "<div>Site Url: " + listItem.get_item("SiteUrl") + "</div>" +
        "<div>Language: " + listItem.get_item("Language") + "</div>" +
        "<div>Template: " + listItem.get_item("WebTemplate") + "</div>" +
        "<div>Date and time: " + (new Date()).format("dd.MM.yyyy HH:mm:ss") + "</div>" +
        "</div>";
      SP.UI.Notify.addNotification(strHtml, true);
    }, 
    function (sender, args){
      if( note ) {
        SP.UI.Notify.removeNotification(note);
      }
      errorCommon("AddDeploymentRequestCallback, creating list item: " + args.get_message(), true);
    });
  }

  function AdjustDialogLogicAndUI(dlg) {
    var doc = dlg.contents();
    var dlgWindow = dlg[0].contentWindow;
    var btnOk = doc.find("input[name$='BtnCancel']").prev("input");
    var languageSelector = doc.find("select[name$='LanguageWebTemplate']:first");
    btnOk.click(function(){
      var isValid = true;
      if( typeof(dlgWindow.Page_ClientValidate) == "function" ) {
        dlgWindow.Page_ClientValidate();
        isValid = dlgWindow.Page_IsValid;
      }
      if( !isValid ) return;
      SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog",function(){
        var inputData = {};
        inputData["Title"] = doc.find("input[name$='TxtCreateSubwebTitle']:first").val();
        inputData["Description"] = doc.find("textarea[name$='TxtCreateSubwebDescription']:first").val();
        inputData["ManagedPath"] = doc.find("select[name$='ManagedPath']:first").val();
        inputData["SiteUrl"] = doc.find("input[name$='TxtCreateSubwebName']:first").val();
        inputData["Language"] = languageSelector.val();
        inputData["WebTemplate"] = doc.find("select[name$='LbWebTemplate']:first").val();
        SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.ok, inputData);
      });
    });
    doc.find("tr[id*='idPermSection']").nextUntil("tr[id$='HiddenSection']").hide();
    doc.find("tr[id$='HiddenSection']").hide();
    var urlInput = doc.find("input[id$='TxtCreateSubwebName']").first();
    urlInput.parent().prev().text(window.location.protocol + "//" + window.location.host);
    urlInput.before(
      "<select name='ManagedPath' style='margin:0px 5px 0px 5px;'>" +
      "<option selected='true'>/sites/</option>" +
      "<option>/teams/</option>" +
      "</select>");
    urlInput.attr("style", "position:absolute;margin-top:30px;left:30%;width:65%;");
    urlInput.next("span").attr("style", "position:absolute;margin-top:-30px;margin-left:110px;display:none;");
    //urlInput.css("left", (urlInput.parent().prev().offset().left - 22) + "px");
    doc.find("input[name$='TxtCreateSubwebTitle']:first").attr("style", "width:95%");
    doc.find("textarea[name$='TxtCreateSubwebDescription']:first").attr("style", "width:95%");
    AppendMissingWebTemplates(doc, languageSelector);
    EnableCheckForFrameReload("select[name='ManagedPath']");
  }

  function EnableCheckForFrameReload(selector) {
    SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog",function(){
      var intervalCheckForReload = setInterval(function(){
        //console.log("Check for the frame reload has been executed.");
        var dlg = SP.UI.ModalDialog.get_childDialog();
        if( typeof dlg == "undefined" ) {
          // The dialog was probably closed; just clear the interval to prevent any further checks.
          clearInterval(intervalCheckForReload);
          //console.log("Dialog was closed.");
          return;
        }
        try {
          var frameElement = dlg.get_frameElement();
          var doc = jQuery(frameElement).contents();
          //console.log(doc.find(selector));
          if( doc.find(selector).length == 0 ) {
            clearInterval(intervalCheckForReload);
            errorCommon("frame reload detected in the open modal dialog");
            SP.UI.ModalDialog.commonModalDialogClose(SP.UI.DialogResult.cancel, null);
          }
        } catch(e) {
          clearInterval(intervalCheckForReload);
          //console.log("Dialog was closed.");
        }
      }, 2500);
    });
  }
  
  function AppendMissingWebTemplates(doc, languageSelector) {
    doc.find(".ms-templatepickerunselected").detach();                      // Find and remove inactive headers.
    var selectedHeader = doc.find(".ms-templatepickerselected a").first();  // Find the active header.
    selectedHeader.attr("href","javascript:;");
    selectedHeader.click(EnsureWebTemplatesForSelectedGroup);
    languageSelector.attr("onchange","javascript:;");
    languageSelector.on("change", EnsureWebTemplatesForSelectedLanguage);
    languageSelector.trigger("change");
    if( allSettings.WebTemplatesAdvanced.Enabled ) {
      TryLoadAdvancedWebTemplates(function(){languageSelector.trigger("change");});
    }
  }

  function DialogCallback(dialogResult, returnValue) {
    if (dialogResult == SP.UI.DialogResult.ok) {
      var inputData = returnValue;
      AddDeploymentRequest(inputData);      
    }
  }

  function DialogLoaded(dlg) {
    AdjustDialogLogicAndUI(dlg);
  }
  
  function EnsureJQuery(version) {
    // Load jQuery from CDN's URL (the code executes only in the case when it was not loaded by the custom action).
    if( typeof window.jQuery == "undefined" ) {
      var el = document.createElement("script");
      el.type = "text/javascript";
      el.src = "//ajax.aspnetcdn.com/ajax/jquery/jquery-" + version + ".min.js";
      document.getElementsByTagName("head")[0].appendChild(el);
    }
    return true;
  }

  function EnsureListForDeploymentRequests(callback) {
    var listTitle = allSettings.ListForDeploymentRequests.Title;
    var listUrl = allSettings.ListForDeploymentRequests.Url.replace(/^\/+/,'').replace(/\/+$/,'');

    var ctx = new SP.ClientContext.get_current();
    var rootWeb = ctx.get_site().get_rootWeb();
    //var contentTypes = rootWeb.get_contentTypes();
    var lists = rootWeb.get_lists();
    //ctx.load(contentTypes);
    ctx.load(lists,"Include(RootFolder)");
    ctx.executeQueryAsync(function(){
      var listEnumerator = lists.getEnumerator();
      var re = new RegExp("\/" + listUrl + "$", "i");
      while( listEnumerator.moveNext() ) {
        var list = listEnumerator.get_current();
        var serverRelativeUrl = list.get_rootFolder().get_serverRelativeUrl();
        if( serverRelativeUrl.match(re) ) {
          callback(list);
          return;
        }
      }
      
      // List not found; create a new one.
      var listCreationInfo = new SP.ListCreationInformation();
      listCreationInfo.set_title(listTitle);
      listCreationInfo.set_templateType(SP.ListTemplateType.genericList);
      listCreationInfo.set_url(listUrl);
      var list = lists.add(listCreationInfo);
      list.set_contentTypesEnabled(true);
      list.update();  //Update operation is required to apply list changes.

      // Add a custom content type on the list level only (i.e. not on a web).
      var listContentTypes = list.get_contentTypes();
      //var parentContentType = contentTypes.getById("0x0104");     // For example, Announcement
      for( var name in allSettings.ListForDeploymentRequests.ContentTypes ) {
        var contentType = new SP.ContentTypeCreationInformation();
        //contentType.set_group(listTitle);
        //contentType.set_parentContentType(parentContentType);
        contentType.set_name(name.replace(/s$/,''));
        contentType.set_description(listTitle);
        listContentTypes.add(contentType);
      }
      
      console.log("Creating a list ... '" + listTitle + "'(" + ctx.get_url().replace(/\/+/,'') + '/' + listUrl + ")");
      var rootFolder = list.get_rootFolder();
      ctx.load(listContentTypes);
      ctx.load(list, "RootFolder");
      ctx.load(rootFolder, "ContentTypeOrder");
      ctx.executeQueryAsync(function(){
        var enumerator = listContentTypes.getEnumerator();
        var itemContentType = null;
        while( enumerator.moveNext() ) {
          var ct = enumerator.get_current();
          //if( ct.get_name() == "Item" ) {
          itemContentType = ct;
          //}
          break;
        }

        // Set the custom content type as default one.
        rootFolder.set_uniqueContentTypeOrder(rootFolder.get_contentTypeOrder().reverse());
        rootFolder.update();
        
        // Add custom fields to the list.
        var allFields = list.get_fields();
        // TODO: implement a better logic later. Now all fields are added to all content types by default.
        for( var name in allSettings.ListForDeploymentRequests.ContentTypes ) {
          for( var i = 0; i < allSettings.ListForDeploymentRequests.ContentTypes[name].Fields.length; i++ ) {
            // https://msdn.microsoft.com/en-us/library/office/ee553410(v=office.14).aspx
            var field = allSettings.ListForDeploymentRequests.ContentTypes[name].Fields[i];
            allFields.addFieldAsXml(field, true, SP.AddFieldOptions.addFieldInternalNameHint);
          }
        }
        ctx.load(allFields);

        // Remove the previous default system content type (Item).
        itemContentType.deleteObject();
        
        ctx.executeQueryAsync(function(){
          callback(list);
        },function (sender, args){
          errorCommon("EnsureListDeploymentRequest, adding list fields: " + args.get_message(), true);
        });
      },function (sender, args){
        errorCommon("EnsureListDeploymentRequest, creating list: " + args.get_message(), true);
      });

    }, function (sender, args){
      errorCommon("EnsureListDeploymentRequest, loading lists: " + args.get_message(), true);
    });
  }
    
  function EnsureWebTemplatesForSelectedGroup() {
    // How to get reference to dialog's document using Chrome's console:
    // var doc = $(".ms-dlgFrame").contents(); var aSel = doc.find(".ms-templatepickerselected a");aSel.parents(".ms-templatepicker:first");
    var selectedDiv = jQuery(this).parents(".ms-templatepickerselected,.ms-templatepickerunselected");
    selectedDiv.removeClass("ms-templatepickerunselected").addClass("ms-templatepickerselected");
    selectedDiv.siblings(".ms-templatepickerselected,.ms-templatepickerunselected")
      .removeClass("ms-templatepickerselected").removeClass("ms-templatepickerunselected")
      .addClass("ms-templatepickerunselected");

    var parentTable = selectedDiv.parents(".ms-propertysheet:first");
    var lcid = parentTable.find("select[name$='LanguageWebTemplate']:first").val();
    var templates = allSettings.WebTemplates[lcid];
    if( templates == null ) {
      templates = allSettings.WebTemplates[allSettings.WebTemplates.Default];
    }
    var groupTemplates = templates[jQuery(this).text().trim()];
    if( groupTemplates ) {
      var templateSelector = parentTable.find(".ms-templatepicker-select:first");
      if( templateSelector.length > 0 ) {
        templateSelector.empty();
        //templateSelector.off("change");
        var parentForm = parentTable.parents("form:first");
        parentForm.find("input[name^='HidDescription']").detach();
        var aspNetHidden = parentForm.find(".aspNetHidden");
        if( aspNetHidden == null ) {aspNetHidden = parentForm};
        var i = 0;
        for( var name in groupTemplates ) {
          var title = groupTemplates[name]["Title"];
          var description = groupTemplates[name]["Description"];
          if( title == null ) {
            title = name;
          }
          var option = jQuery("<option value='" + name + "'>" + title + "</option>");
          templateSelector.append(option);
          if( description ) {
            var el = jQuery('<input type="hidden" name="HidDescription{0}" id="HidDescription{0}" value="{1}">'
              .replace(/\{0\}/g,i).replace(/\{1\}/g,description));
            aspNetHidden.append(el);
          }
          i++;
        }
        templateSelector.prop('selectedIndex', 0);
        templateSelector.trigger("change");
      }
    }
  }
  
  function EnsureWebTemplatesForSelectedLanguage() {
    var templates = allSettings.WebTemplates[jQuery(this).val()];
    if( templates == null ) {
      templates = allSettings.WebTemplates[allSettings.WebTemplates.Default];
    }
    
    var parentTable = jQuery(this).parents(".ms-propertysheet:first");
    parentTable.find(".ms-templatepickerunselected").detach(); // Find and remove inactive headers.
    // Find the active header.
    var selectedGroup = jQuery(this).parents(".ms-propertysheet:first").find(".ms-templatepickerselected");
    var selectedLink = selectedGroup.find("a:first");
    var i = 0;
    for( var name in templates ) {
      if( i == 0 ) {
        selectedLink.text(name);
        i++;
        continue;
      }
      var cloned = selectedGroup.clone();
      selectedGroup.parent().append(cloned);
      cloned.find("a:first").text(name).click(EnsureWebTemplatesForSelectedGroup);
      cloned.removeClass("ms-templatepickerselected").addClass("ms-templatepickerunselected");
    }
    selectedLink.trigger("click");
  }
  
  function errorCommon(message, displayNotification) {
    console.log("ERROR: " + message);
    if( displayNotification ) {
      var text = 
        "<div style='text-align:left;max-width:400px;color:#f00;font-weight:bold;'>" +
        "<div>ERROR: " + message + "</div>" + 
        "<div>Time: " + (new Date()).format("dd.MM.yyyy HH:mm:ss") + "</div>"
        "</div>";
      SP.UI.Notify.addNotification(text, true);
    }
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
  
  function OpenModalDialog(dlgTitle, dlgUrl, callbackFunction) {
    SP.SOD.executeFunc("sp.ui.dialog.js", "SP.UI.ModalDialog",
      function() {
        var _options = {
          title: dlgTitle,
          url: dlgUrl,
          autoSize: true,
          height:730,
          //showMaximized: true,
          //allowMaximize: true,
          //showClose: true,
          dialogReturnValueCallback: DialogCallback
        };
        var dialog = SP.UI.ModalDialog.showModalDialog(_options);

        var timeoutLoadContent = setTimeout(function(){clearInterval(intervalLoadContent);}, 10000);
        var intervalLoadContent = setInterval(function() {
          var dlg = jQuery(dialog.get_frameElement());
          if( dlg.contents().find("input[name$='BtnCancel']").length > 0 ) {
            clearInterval(intervalLoadContent);
            clearTimeout(timeoutLoadContent);
            callbackFunction(dlg);
            return true;
          };
        }, 100);
      }
    );
  }

  function OpenModalDialogOnThisWeb(callbackFunction) {
    FindCurrentWebUrlAndExecute(function(siteUrl){
      // This callback function is executed when siteUrl is calculated by FindCurrentWebUrlAndExecute.
      function OpenModalDialogOnThisWebCallback(siteUrl, callbackFunction) {
        try {
          var dlgTitle = jQuery("ie\\:menuitem[onmenuclick*='" + allSettings.ScriptName + "']").attr("text");
          if( dlgTitle == null ) {
            dlgTitle = allSettings.DialogTitleDefault;
          }
          
          var dialogUrl = allSettings.DialogUrl.replace(/^\/+/,'');
          
          var dlgUrl = siteUrl.replace(/\/+$/,'') + '/' + dialogUrl;
          OpenModalDialog(dlgTitle, dlgUrl, callbackFunction); // This callback function is executed on closing a dialog.
        } catch(e) {
          errorCommon("OpenModalDialogOnThisWeb, using jQuery: " + e.message);
        }
      }
      OpenModalDialogOnThisWebCallback(siteUrl, callbackFunction);
    });
  }

  function TryLoadAdvancedWebTemplates(callbackFunction) {
    try {
      //console.log("P1");
      FindCurrentWebUrlAndExecute(function(siteUrl) {
        //console.log("P2");
        jQuery.getScript(siteUrl + allSettings.WebTemplatesAdvanced.ScriptForStandardTemplates.replace(/^\/+/,''))
        .done(jQuery.getScript(siteUrl + allSettings.WebTemplatesAdvanced.ScriptForCustomTemplates.replace(/^\/+/,'')))
        .done(function(){
          //console.log("P3");
          function ContinueAfterLoad() {
            try {
              if( typeof webTemplates == "undefined"
                  || typeof webTemplates.Standard != "function" 
                  || typeof webTemplates.Custom != "function" ){
                setTimeout(ContinueAfterLoad, 100);
                //console.log("P4");
                return;
              }
              //console.log("P5");
              webTemplates.All = jQuery.extend(true, webTemplates.Standard(), webTemplates.Custom());
              var itemCount = Object.keys(webTemplates.All[webTemplates.All.Default]).length;
              //console.log("P6: " + itemCount);
              for(var key in webTemplates.All) {
                var localeData = webTemplates.All[key];
                if( typeof localeData != "object" ) continue; // Default is a number.
                if(Object.keys(localeData).length < itemCount) {
                  jQuery.extend(true, localeData, webTemplates.Custom()[webTemplates.All.Default]);
                }
              }
              //console.log("P7");
              allSettings.WebTemplates = webTemplates.All;
              callbackFunction();
            } catch(e) {
              errorCommon("TryLoadAdvancedWebTemplates, processing dynamic web-templates: " + e.message, false);
            }                  
          }
          setTimeout(ContinueAfterLoad, 100);
        });
      });
    } catch(e) {
      errorCommon("TryLoadAdvancedWebTemplates, loading dynamic web-templates: " + e.message, false);
    }
  } 

  EnsureJQuery(allSettings.JqueryVersion);
  OpenModalDialogOnThisWeb(DialogLoaded);
  jQuery("script[src*='" + allSettings.ScriptName + "']").detach();
})();