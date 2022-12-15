$(document).ready(function () {

  $("#run").click(() => tryCatch(run));
  
});

async function run() {
  await Excel.run(async (context) => {
logtoSPList();
    }
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error(error);
  }
}

function logtoSPList() {

  var authenticator;
  var client_id = "2e1be2b2-01f2-466e-84cd-65f2b689fbce";
  var redirect_url = "https://mikiyks.github.io/inkan/";
  var scope = "https://graph.microsoft.com/sites.readwrite.all";
  var access_token;

  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  authenticator.endpoints.registerMicrosoftAuth(client_id, {
    redirectUrl: redirect_url,
    scope: scope
  });

    authenticator
      .authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
      .then(function (token) {
        access_token = token.access_token;

      //API呼び出し
      $(function () {
        $.ajax({
          url: "https://graph.microsoft.com/v1.0/sites/everyone/lists/6aac0560-622e-4ee1-ba8f-73b32d8e9f05/items",
          type: "POST",
          beforeSend: function (xhr) {
            xhr.setRequestHeader("Authorization", "Bearer " + access_token);
          },
          Authorization: "Bearer " + access_token,
          data: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'FileName': "excel",
            'Title': "testMAN"
          }),
          headers: {
            "Accept": "application/json; odata=nometadata",
            "Content-Type": "application/json;odata=verbose"
          },
          success: function (data) {
            console.log("logOK");
          },
          error: function (data) {
            console.log(data);
          }
        });
      });
    })
    .catch(OfficeHelpers.Utilities.log);
}
