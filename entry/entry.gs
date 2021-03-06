/*
定数・グローバル変数の設定
*/

// 外生的に決定する定数
const JNLID = '1FO9Lpw8qpjuWtOi0hcdNxUH0O0_iASFEijPZs3q16uU';
const FSID = '1o0hYHuFBy7f9336fCpZsC4up9N1qNjjyhr-KzkKfX-I'; // 財務諸表シートの ID
const ENTURL = "https://script.google.com/macros/s/AKfycbypB1J210SxBc8sycedUaLlL6VNsVIPEFuvWNdN/exec";
const CALCURL = 'https://script.google.com/macros/s/AKfycbxFGIqy1AWG5Gk679FAqWu4EMUClLrQcLY0SXH8uQ/exec';
const ANALURL = 'https://datastudio.google.com/s/oU7k5aPgNJA';

// ID からスプレッドシートの取得
const JNL = SpreadsheetApp.openById(JNLID);
const JNLURL = 'https://docs.google.com/spreadsheets/d/' + JNLID + '/edit';
const JNLNames = JNL.getSheets().map(x => x.getSheetName());
const FSURL = 'https://docs.google.com/spreadsheets/d/' + FSID + '/edit';

// エラーメッセージ
let ERR;

// 勘定科目名の取得
const fullComp = getData(JNL.getSheetByName('COA')).slice(1);
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
        .addMetaTag('viewport', 'width=device=width, initial-scale=1')
        .setTitle('bkeep-entry');
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
    
    // 日付の取得
    let date = pm.date.toString();
    let chdate = date.replace(/-/g, "")
    
    // シート取得
    let sheet = setSheet(chdate);
    
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
    let accountdata = [];
    for (let i = 0; i < n; i++) {
        // 変数定義
        let account = pm.account[i].toString();
        
        if (account !== "") {
            // account が Comp に含まれない場合、post しない
            if (Comp.indexOf(account) === -1) {
                ERR = "予定されていない勘定科目名を使っています。";
                let resultOutput = HtmlService.createTemplateFromFile('errmsg').evaluate();
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
            accountdata.push([time, chdate, account, amount, adj, note]);
        }
    }
    
    // 入力フォーマットのチェック
    if (accountdata.length < 2) {
        ERR = "2 行以上の入力が必要です。";
        let resultOutput = HtmlService.createTemplateFromFile('errmsg').evaluate();
        resultOutput
        // スマホ対応
            .addMetaTag('viewport', 'width=device=width, initial-scale=1');
        return resultOutput;
    } else if (accountdata.filter(x => x[3] === 0).length > 1) {
        ERR = "金額未入力が 2 箇所以上存在します。";
        let resultOutput = HtmlService.createTemplateFromFile('errmsg').evaluate();
        resultOutput
        // スマホ対応
            .addMetaTag('viewport', 'width=device=width, initial-scale=1');
        return resultOutput;
    }
    
    // 金額の補完
    zeroind = accountdata.map(x => x[3]).indexOf(0);
    if (zeroind !== -1) {
        accountdata[zeroind][3] = (-1) * 
            accountdata.map(x => x[3]).reduce((sum, element) => sum + element, 0);
    }
    
    // 借方合計と貸方合計の一致
    if (accountdata.map(x => x[3]).reduce((sum, element) => sum + element, 0) !== 0) {
        ERR = "借方合計と貸方合計が一致しません。";
        let resultOutput = HtmlService.createTemplateFromFile('errmsg').evaluate();
        resultOutput
        // スマホ対応
            .addMetaTag('viewport', 'width=device=width, initial-scale=1');
        return resultOutput;
    }
    
    // データの挿入
    for (let i = 0; i < accountdata.length; i++) {
        sheet.appendRow(accountdata[i]);
    }
    
    // 確認データの取得
    mkConfirm(sheet, chdate);
    
    // 出力
    let resultOutput = HtmlService.createTemplateFromFile('result').evaluate();
    resultOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1')
        .setTitle('bkeep-entry-result');
    return resultOutput;
}

// 確認データの作成
function mkConfirm (sheet, chdate) {
    let contable = getData(sheet).map(x => x.slice(1));
    let colnames = contable.shift();
    contable = contable.filter(x => x[0] == chdate);
    confirmdata = '<table><tr>';
    // ヘッダー
    colnames.forEach(account => confirmdata = confirmdata + '<th>' + account + '</th>');
    confirmdata = confirmdata + '</tr>';
    // 本体
    for (let i = 0; i < contable.length; i++) {
        confirmdata = confirmdata + '<tr>';
        contable[i].forEach(account => confirmdata = confirmdata + '<td>' + account + '</td>');
        confirmdata = confirmdata + '</tr>';
    }
    confirmdata = confirmdata + '</table>';
}

// 入力先のシートを作成・選択する
function setSheet (chdate) {  
    // 年月を取得
    let YM = chdate.slice(0, 6);
    
    // 指定年月のシートが存在するか否かで条件分岐
    if (JNLNames.indexOf(YM) === -1) {
        JNL.insertSheet(YM);
        let sheet = JNL.getSheetByName(YM);
        sheet.appendRow(['time', 'date', 'account', 'amount', 'adj', 'note']);
    } 
    
    let sheet = JNL.getSheetByName(YM);
    return sheet;
}
