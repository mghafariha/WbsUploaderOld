<script type="text/javascript">//<![CDATA[
          $(document).ready(function () {       
            
              var userid = _spPageContextInfo.userId;

              var requestUri = _spPageContextInfo.webAbsoluteUrl + "/_api/web/getuserbyid(" + userid + ")";
              var requestHeaders = { "accept": "application/json;odata=verbose" };
              $.ajax({
                  url: requestUri,
                  contentType: "application/json;odata=verbose",
                  headers: requestHeaders,
                  success: onSuccess,
                  error: onError
              }); function onSuccess(data, request) {
       
                  var itemType = GetItemTypeForListName("Statistics");
                  var item = {
                      "__metadata": { "type": itemType },
                      "Title": "HNLog",
                      "url":window.location.href,
                      "date":new Date(),
                      "uname":data.d.Title
                  };
      
                  $.ajax({
                      url: _spPageContextInfo.siteAbsoluteUrl + "/_api/web/lists('{C7FE4249-C610-4A4F-8B35-FCF69A5C97FD}')/items",
                      type: "POST",
                      contentType: "application/json;odata=verbose",
                      data: JSON.stringify(item),
                      headers: {
                          "Accept": "application/json;odata=verbose",
                          "X-RequestDigest": $("#__REQUESTDIGEST").val()
                      },
                      success: function (data) {
                          //  success(data);
                          //console.log(data);
                      },
                      error: function (data) {
                          // failure(data);
                          //console.log(data);
                      }
                  });
              
                  function GetItemTypeForListName(name) {
                      return "SP.Data." + name.charAt(0).toUpperCase() + name.split(" ").join("").slice(1) + "ListItem";
                  }
       
       
              } function onError(error) {
                  console.log("Error on retrieving current user.");
              }
          
          
          
          
              //addNewItem("/_api/Web/lists('{C7FE4249-C610-4A4F-8B35-FCF69A5C97FD}')/items",{Title:'hnLog',url:loc,date:date,uname:userid.toString()},function(){console.log('done')},function(e){console.log(e)});
          
          });
//]]></script>