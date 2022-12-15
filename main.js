$(document).ready(function () {

  getUser();


  var dt = new Date();
  var txtDate = dt.getFullYear().toString() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate();
  $("#date").val(txtDate);
  $("#run").click(() => tryCatch(run));
  //日付不要にチェック入れたら日付グレーアウト
  $("#dateCheckBox").change(() => tryCatch(change));
  function change() {
    if ($("#dateCheckBox").prop("checked")) {
      $("#date").prop("disabled", true);
    } else {
      $("#date").prop("disabled", false);
    }
  }
});

async function run() {
  await Excel.run(async (context) => {
    getfilename();
    //名前が空なら処理なし
    if (
      !$("#name")
        .val()
        .toString()
    ) {
    } else {
      //アクティブセルの位置取得
      const cell = context.workbook.getActiveCell();
      cell.load("left").load("top");
      await context.sync();
      //印鑑生成実行
      onWorkSheetSingleClick(cell.left, cell.top);
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

//アクティブセルに押印
async function onWorkSheetSingleClick(x, y) {
  await Excel.run(async (context) => {
    //変数宣言//
    var fontName = "HGS行書体"; //名前テキストのフォント
    var objectColor = "FF2000"; //線色・文字色
    var txtName = $("#name")
      .val()
      .toString(); //名前テキスト
    var lenName = txtName.length; //名前文字数
    if (lenName < 3) {
      var fontSize = 22;
    } else {
      var fontSize = 19;
    } //名前文字数でフォントサイズ調整
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const shapes = sheet.shapes;

    //印鑑枠作成
    const ellipse = shapes.addGeometricShape(Excel.GeometricShapeType.ellipse);
    if (lenName < 3) {
      ellipse.left = 14.3;
      ellipse.top = 15.9;
    } else {
      ellipse.left = 10;
      ellipse.top = 10;
    }
    ellipse.height = 31.1;
    ellipse.width = 31.1;
    ellipse.fill.transparency = 1;
    ellipse.lineFormat.weight = 1;
    ellipse.lineFormat.color = objectColor;

    //日付テキスト, 日付不要ならスルー
    if ($("#dateCheckBox").prop("checked")) {
    } else {
      var shpDateText = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
      if (lenName < 3) {
        shpDateText.left = 5.5;
        shpDateText.top = 47.9;
      } else {
        shpDateText.left = 0.6;
        shpDateText.top = 42;
      }
      shpDateText.height = 12;
      shpDateText.width = 50;
      shpDateText.textFrame.leftMargin = 0;
      shpDateText.textFrame.bottomMargin = 0;
      shpDateText.textFrame.rightMargin = 0;
      shpDateText.textFrame.topMargin = 0;
      shpDateText.fill.transparency = 1;
      shpDateText.lineFormat.transparency = 1;
      shpDateText.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
      shpDateText.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
      shpDateText.textFrame.verticalOverflow = Excel.ShapeTextVerticalOverflow.overflow;
      shpDateText.textFrame.horizontalOverflow = Excel.ShapeTextHorizontalOverflow.overflow;
      var trngDateText = shpDateText.textFrame.textRange;
      trngDateText.font.color = objectColor;
      trngDateText.font.name = "Calibri"; //fontName;
      trngDateText.font.size = 8;
      trngDateText.text =
        "'" +
        $("#date")
          .val()
          .toString()
          .slice(2, 4) +
        "." +
        $("#date")
          .val()
          .toString()
          .slice(5, 7) +
        "." +
        $("#date")
          .val()
          .toString()
          .slice(8, 10);
    }

    //名前テキスト
    const shpNameText = shapes.addGeometricShape(Excel.GeometricShapeType.rectangle);
    if (lenName < 3) {
      shpNameText.height = lenName * 23;
    } else {
      shpNameText.height = lenName * 20;
    }
    shpNameText.width = 27.7;
    shpNameText.fill.transparency = 1;
    shpNameText.lineFormat.transparency = 1;
    shpNameText.textFrame.verticalAlignment = Excel.ShapeTextVerticalAlignment.middle;
    shpNameText.textFrame.horizontalAlignment = Excel.ShapeTextHorizontalAlignment.center;
    shpNameText.textFrame.leftMargin = 0;
    shpNameText.textFrame.bottomMargin = 0;
    shpNameText.textFrame.rightMargin = 0;
    shpNameText.textFrame.topMargin = 0;
    shpNameText.textFrame.verticalOverflow = Excel.ShapeTextVerticalOverflow.overflow;
    shpNameText.textFrame.horizontalOverflow = Excel.ShapeTextHorizontalOverflow.overflow;
    shpNameText.textFrame.orientation = "EastAsianVertical";
    const trngNameText = shpNameText.textFrame.textRange;
    trngNameText.font.color = objectColor;
    trngNameText.font.name = fontName;
    trngNameText.font.size = fontSize;
    trngNameText.text = txtName;

    //名前テキストの画像化
    const preImgNameText = shpNameText.getAsImage(Excel.PictureFormat.png);
    await context.sync();
    const imgNameText = shapes.addImage(preImgNameText.value);
    if (lenName < 3) {
      imgNameText.height = 52;
      imgNameText.left = 2.6;
      imgNameText.top = 6;
    } else {
      imgNameText.height = 40;
      imgNameText.left = 1.5;
      imgNameText.top = 6;
    }

    //グループ+画像化
    if ($("#dateCheckBox").prop("checked")) {
      var shpStamp = shapes.addGroup([ellipse, imgNameText]);
    } else {
      var shpStamp = shapes.addGroup([ellipse, shpDateText, imgNameText]);
    }
    const shpStampPreImage = shpStamp.getAsImage(Excel.PictureFormat.png);
    await context.sync();
    const shpStampImage = shapes.addImage(shpStampPreImage.value);
    shpStampImage.name = "印鑑";

    //素材削除
    shpStamp.group.ungroup();
    ellipse.delete();
    if ($("#dateCheckBox").prop("checked")) {
    } else {
      shpDateText.delete();
    }
    imgNameText.delete();
    shpNameText.delete();

    shpStampImage.left = x;
    shpStampImage.top = y;
    await context.sync();
  });
}

Office.initialize = function (reason) {
  if (OfficeHelpers.Authenticator.isAuthDialog()) return;
};

function getUser() {
  var authenticator;
  var client_id = "2e1be2b2-01f2-466e-84cd-65f2b689fbce";
  var redirect_url = "https://mikiyks.github.io/inkan/";
  var scope = "https://graph.microsoft.com/user.read";
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
            url: "https://graph.microsoft.com/v1.0/me",
            type: "GET",
            beforeSend: function (xhr) {
              xhr.setRequestHeader("Authorization", "Bearer " + access_token);
            },
            success: function (data) {
              //取得した苗字をセット
              $("#name").val(data.surname);
            },
            error: function (data) {
              console.log(data);
            }
          });
        });
        return { access_token: access_token };
      })
      .catch(OfficeHelpers.Utilities.log);
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
          data: JSON.stringify({
            '__metadata': { 'type': 'SP.List' },
            'FileName': $("#filename").val(),
            'Title': $("#name").val()
          }),
          headers: {
            "Accept": "application/json; odata=nometadata",
            "X-RequestDigest": $("#__REQUESTDIGEST").val()
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

function getfilename() {
  Office.context.document.getFilePropertiesAsync(function (asyncResult) {
    let filename = asyncResult.value.url
      .split("/")
      .reverse()[0]
      .split(".")[0];
    let extend = asyncResult.value.url
      .split("/")
      .reverse()[0]
      .split(".")[1];
    //console.log(filename + "." + extend);
    $("#filename").val(filename + "." + extend);
  });
}
