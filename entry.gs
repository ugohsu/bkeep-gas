// ID からスプレッドシートの取得
const SSID = '1TB_sYFJ0QccVXWC1WV0-S2pmusmLM0Qy81rlDL_3XD8';
const SS = SpreadsheetApp.openById(SSID);
const SSURL = 'https://docs.google.com/spreadsheets/d/' + SSID + '/edit'

// index.html の URL
const URL = "https://script.google.com/macros/s/AKfycbxiPVp-bIj3j-SwgaHhbieBGqx6zQMNVbwC2nKa9A/exec"

// アクセス可能なパスワード
const Password = ["1111"];

// 勘定科目名の取得
const fullComp = getData(SS.getSheetByName('definition')).slice(1);
const Comp = fullComp.map(x => x[1]);
const CompWithLabel = mkComp(); // html タグ

// 確認用データの宣言
let confirmdata = "";

// 関数定義
// doGet
function doGet() {
    const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
    htmlOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1');
    return htmlOutput;
}

// スプレッドシートのデータ取得
function getData (sheet) {
    let values = sheet.getDataRange().getValues();
    return values;
}

// 補完文字列の取得 (入力補完用)
function mkComp () {
    otext = '<datalist id="completion">';
    for (let i = 0; i < Comp.length; i++) {
        otext = otext + '<option value="' + Comp[i] + '" label="' + fullComp[i][0] + '"></option>';
    }
    otext = otext + '</datalist>';
    return otext;
}

function doPost (postdata) {

    
    // エイリアスの設定
    let pm = postdata.parameters
    
    // 認証
    /*
    if (Password.indexOf(pm.pass.toString()) === -1) {
        // Logger.log("あなたには実行権限がありません。");
        return;
    }
    */
    
    // 日付の取得
    let date = pm.date.toString();
    let chdate = date.replace(/-/g, "")
    
    // シート取得
    let sheet = setSheet(SS, chdate);
    
    // 種々の変数の定義
    let time = new Date();
    let note = pm.note.toString();
    let n = pm.DC.length;
    let adj = pm.adj;
    if (adj == null)
        adj = 0;
    else
        adj = 1;
    
    // 入力データの作成
    let itemdata = [];
    for (let i = 0; i < n; i++) {
        // 変数定義
        let item = pm.item[i].toString();
        
        if (item !== "") {
            // item が Comp に含まれない場合、post しない
            if (Comp.indexOf(item) === -1) {
                let resultOutput = HtmlService.createTemplateFromFile('errinput').evaluate();
                resultOutput
                // スマホ対応
                    .addMetaTag('viewport', 'width=device=width, initial-scale=1');
                return resultOutput;
            }
            
            // 金額の定義
            let amount = Number(pm.amount[i]);
            
            // 金額の修正 (貸方の場合 -1 をかける)
            if (pm.DC[i] === "Cr")
                amount = amount * (-1);
            
            // 出力要素の追加
            itemdata.push([time, chdate, item, amount, adj, 0, note]);
        }
    }
    
    // 入力フォーマットのチェック
    if (itemdata.length < 2 || itemdata.filter(x => x[3] === 0).length > 1) {
        let resultOutput = HtmlService.createTemplateFromFile('errformat').evaluate();
        resultOutput
        // スマホ対応
            .addMetaTag('viewport', 'width=device=width, initial-scale=1');
        return resultOutput;
    }
    
    // 金額の補完
    zeroind = itemdata.map(x => x[3]).indexOf(0);
    if (zeroind !== -1) {
        itemdata[zeroind][3] = (-1) * 
            itemdata.map(x => x[3]).reduce((sum, element) => sum + element, 0);
    }
    
    // 借方合計と貸方合計の一致
    if (itemdata.map(x => x[3]).reduce((sum, element) => sum + element, 0) !== 0) {
        let resultOutput = HtmlService.createTemplateFromFile('errbalance').evaluate();
        resultOutput
        // スマホ対応
            .addMetaTag('viewport', 'width=device=width, initial-scale=1');
        return resultOutput;
    }
    
    // データの挿入
    for (let i = 0; i < itemdata.length; i++) {
        sheet.appendRow(itemdata[i]);
    }
    
    // 確認データの取得
    mkConfirm(sheet, chdate);
    
    // 出力
    let resultOutput = HtmlService.createTemplateFromFile('result').evaluate();
    resultOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1');
    return resultOutput;
}

// 確認データの作成
function mkConfirm (sheet, chdate) {
    let contable = getData(sheet).map(x => x.slice(1));
    let colnames = contable.shift();
    contable = contable.filter(x => x[0] == chdate);
    confirmdata = '<table><tr>';
    // ヘッダー
    colnames.forEach(item => confirmdata = confirmdata + '<th>' + item + '</th>');
    confirmdata = confirmdata + '</tr>';
    // 本体
    for (let i = 0; i < contable.length; i++) {
        confirmdata = confirmdata + '<tr>';
        contable[i].forEach(item => confirmdata = confirmdata + '<td>' + item + '</td>');
        confirmdata = confirmdata + '</tr>';
    }
    confirmdata = confirmdata + '</table>';
}

// 入力先のシートを作成・選択する
function setSheet (ss = SS, chdate) {  
    // 年月を取得
    let YM = chdate.slice(0, 6);
    
    // シート名一覧を取得
    let SheetNames = ss.getSheets().map(x => x.getSheetName());
    
    // 指定年月のシートが存在するか否かで条件分岐
    if (SheetNames.indexOf(YM) === -1) {
        ss.insertSheet(YM);
        let sheet = ss.getSheetByName(YM);
        sheet.appendRow(['time', 'date', 'item', 'amount', 'adj', 'transfer', 'note']);
    } 
    
    let sheet = ss.getSheetByName(YM);
    return sheet;
}

