$(document).ready( function(){
var dt = new Date();
var txtDate = dt.getFullYear().toString() + "-" + (dt.getMonth() + 1) + "-" + dt.getDate();
$('#date').val(txtDate);
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

      var access_token;
      var client_id = '507591d2-cfb9-4a52-bdf6-6053cfcc3ff3';
      var scope = 'https://graph.microsoft.com/user.read';
      var url_auth = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize' +
                     '?response_type=token' +
                     '&client_id=' + client_id +
                     '&redirect_uri=' + encodeURIComponent(location.protocol + "//" + location.host + location.pathname) +
                     '&scope=' + encodeURIComponent(scope);
       
      $(function(){
        //access_tokenの取得
        if(location.hash){
          var hashary = location.hash.substr(1).split('&');
          $.each(hashary,function(i, v){
            var ary = v.split('=');
            if(ary[0] == 'access_token'){
              access_token = ary[1];
              $('#exec').prop('disabled', false);
              return false;
            }
          });
        }
         
        $('#login').click(function(){
          location.href = url_auth;
        });
         
        //API呼び出し
        $('#exec').click(function(){
          //alert(access_token); //確認用
          $.ajax({
            url: 'https://graph.microsoft.com/v1.0/me',
            type: 'GET',
            beforeSend: function(xhr){
              xhr.setRequestHeader('Authorization', 'Bearer ' + access_token);
            },
            success: function(data){
              //console.log(data); //確認用
              alert(data.displayName);
            },
            error: function(data){
              console.log(data);
            }
          });
        });
      });