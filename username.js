var authenticator;
var client_id = "2e1be2b2-01f2-466e-84cd-65f2b689fbce";
var redirect_url = "https://mikiyks.github.io/inkan/";
var scope = "https://graph.microsoft.com/user.read";
var access_token;

Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
}

function getUser() {
  authenticator = new OfficeHelpers.Authenticator();

  //access_token取得
  $(function () {
    authenticator.endpoints.registerMicrosoftAuth(client_id, {
      redirectUrl: redirect_url,
      scope: scope
    });

    authenticator.authenticate(OfficeHelpers.DefaultEndpoints.Microsoft)
      .then(function (token) {
        access_token = token.access_token;
        $("#exec").prop("disabled", false);
        //API呼び出し
        $(function () {
          $.ajax({
            url: 'https://graph.microsoft.com/v1.0/me',
            type: 'GET',
            beforeSend: function (xhr) {
              xhr.setRequestHeader('Authorization', 'Bearer ' + access_token);
            },
            success: function (data) {
              //取得したい365データ
              var userSurName = data.surname;
              var userGiveName = data.giveName;
		console.log(userSurName);
		console.log("mae");
              return userSurName;
            },
            error: function (data) {
              console.log(data);
            }
          });
        });
      })
      .catch(OfficeHelpers.Utilities.log);
  });
};
