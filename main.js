$(document).ready(function() {
  $("#run").click(() => tryCatch(getkakuin));
});

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

Office.initialize = function(reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

function getkakuin() {
  var authenticator;
  var client_id = "2e1be2b2-01f2-466e-84cd-65f2b689fbce";
  var redirect_url = "https://mikiyks.github.io/inkan/";
  var scope = "https://graph.microsoft.com/Sites.Read.All";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

  authenticator
    .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
    .then(function(token) {
      access_token = token.access_token;
      //API呼び出し
      $(function() {
        $.ajax({
          url:
            "https://graph.microsoft.com/v1.0/sites/20531fc2-c6ab-4e1e-a532-9c8e15afed0d/drive/items/01SG44IHMJY6HM4OB2XJGZ34EYB77ZANB2/content",
          type: "GET",
          beforeSend: function(xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
            xhr.overrideMimeType("image/png");
          },
          success: function(data, status) {

            console.log(data);
            console.log(status);
            console.log(data.responseText);
            console.log(data.responseURL);
            console.log(status.responseURL);
            console.log(JSON.parse(data));

          },
          error: function(data) {
            console.log(data);
          }
        });
      });
      return { access_token: access_token };
    })
    .catch(OfficeHelpers.Utilities.log);
}
