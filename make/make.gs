/*
定数・グローバル変数の設定
*/

// 外生的に決定する定数
const JNLID = '1FO9Lpw8qpjuWtOi0hcdNxUH0O0_iASFEijPZs3q16uU';
const FSID = '1o0hYHuFBy7f9336fCpZsC4up9N1qNjjyhr-KzkKfX-I'; // 財務諸表シートの ID
const ENTURL = "https://script.google.com/macros/s/AKfycbypB1J210SxBc8sycedUaLlL6VNsVIPEFuvWNdN/exec";
const CALCURL = 'https://script.google.com/macros/s/AKfycbxFGIqy1AWG5Gk679FAqWu4EMUClLrQcLY0SXH8uQ/exec';
const ANALURL = 'https://datastudio.google.com/s/oU7k5aPgNJA';

// スプレッドシートの取得
const JNL = SpreadsheetApp.openById(JNLID); // 帳簿
const JNLURL = 'https://docs.google.com/spreadsheets/d/' + JNLID + '/edit';
const JNLNames = JNL.getSheets().map(x => x.getSheetName());
const FSS = SpreadsheetApp.openById(FSID); // 財務諸表
const FSSHEET = FSS.getSheetByName('fs'); // メインデータ
const FSTIDY = FSS.getSheetByName('tidy'); // 整然データ
const FSURL = 'https://docs.google.com/spreadsheets/d/' + FSID + '/edit';
let FSARRAY = convertArray(FSS.getSheetByName('fs').getDataRange().getValues());


// 勘定科目名の取得
const Def = convertArray(JNL.getSheetByName('COA').getDataRange().getValues());
const fullComp = Def['body'];
const Comp = fullComp.map(x => x[1]);
const TAGLIST = mkTagList();

// その他の変数・定数
const today = new Date();
let TIDY = {};
let TABLELIST = {};
let WARNINGS = '';

// 当月の計算
// init
initFS(Comp);

// 当月の取得
const YM = thisMonth();
let MONTHFS;

// 指定年月のシートが存在するか否かで条件分岐
if (JNLNames.indexOf(YM) != -1) {
    let monthjnl = convertArray(JNL.getSheetByName(YM).getDataRange().getValues());
    reflectProgress(monthjnl);
    MONTHFS = calcFS(monthjnl, Number(YM));
    registerFS(MONTHFS, Number(YM));
    mkTidy();
    writeSpreadSheet();
} 

mkTable();

/*
メイン関数
*/

// doGet
function doGet() {
    const htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
    htmlOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1')
        .setTitle('bkeep-make');
    return htmlOutput;
}

// doPost (すべての取得可能な決算月について再計算をする)
function doPost (postdata) {

    // JNLNames のうち、YYMMMM で構成されるシート名を取得
    let trgNames = JNLNames
        .filter(x => x.match(/^[0-9]{6}$/))
        .map(x => Number(x))
        .sort((a, b) => a - b);
        
    // 実施
    trgNames.forEach(
        function (trg) {
            let monthjnl = convertArray(JNL.getSheetByName(trg.toString()).getDataRange().getValues());
            if (trg == Number(YM)) reflectProgress(monthjnl);
            MONTHFS = calcFS(monthjnl, trg);
            registerFS(MONTHFS, trg);
        }
    );
    mkTidy();
    writeSpreadSheet();
    mkTable();
    
    // 出力
    let resultOutput = HtmlService.createTemplateFromFile('index').evaluate();
    resultOutput
    // スマホ対応
        .addMetaTag('viewport', 'width=device=width, initial-scale=1')
        .setTitle('bkeep-make');
    return resultOutput;
}

// 日割計算を反映して仕訳帳データを修正する
function reflectProgress (monthdata) {
    // 帳簿の特定月のデータの読み込み
    let account = monthdata.index['account'];
    let adj = monthdata.index['adj'];
    let amount = monthdata.index['amount'];
    
    // progress を適用
    let progress = calcProgress(today);
    for (let i = 0; i < monthdata['body'].length; i++) {
        // 日割計算科目については amount に進捗度を乗じる
        if (monthdata['body'][i][adj] == 1) {
            let x = monthdata['body'][i][amount] * progress;
            if (x > 0) {
                x = Math.floor(x);
            } else {
                x = Math.ceil(x);
            }
            monthdata['body'][i][amount] = x;
        }
    }
}


// 財務諸表要素の計算
function calcFS (monthdata, YM) {

    // 要素ごとに勘定を計算する
    let ASSETS = calcAccounts(monthdata, 'assets');
    let LIABILITIES = calcAccounts(monthdata, 'liabilities');
    let EQUITY = calcAccounts(monthdata, 'equity');
    let REVENUES = calcAccounts(monthdata, 'revenues');
    let EXPENSES = calcAccounts(monthdata, 'expenses');
    let EARNINGS = {};
    
    // 当期純利益の計算
    EARNINGS['当期純利益'] = 
        Object.values(REVENUES).reduce((sum, element) => sum + element, 0) -
        Object.values(EXPENSES).reduce((sum, element) => sum + element, 0);
        
    // 当期純利益の反映
    EQUITY['留保利益'] = EQUITY['留保利益'] + EARNINGS['当期純利益'];
     
    // 財務諸表の作成
    let FS = {
        'assets': ASSETS, 'liabilities': LIABILITIES, 'equity': EQUITY,
        'revenues': REVENUES, 'expenses': EXPENSES, 'earnings': EARNINGS
    };
    
    // 前期の貸借対照表価額を足しあわせる
    mkStocks(FS, YM)
    
    // 財務諸表の要素の集計値
    FS['elements'] = {};
    ["assets", "liabilities", "equity", "revenues", "expenses"].forEach(
        function (element) {
            FS.elements[element] = Object.values(FS[element])
                .reduce((sum, x) => sum + x, 0);
        }
    )
    
    let diff = FS.elements.assets - FS.elements.liabilities - FS.elements.equity;
    if (diff != 0)
        WARNINGS = WARNINGS + YM + ' における貸借対照表の貸借差額: ' + diff.toString() + '<br>';
    
    // タグの集計値
    FS['tags'] = {};
    uniq(TAGLIST.map(x => x[1])).forEach(
        function (tag) {
            let value = 0;
            TAGLIST.filter(x => x[1] == tag).map(x => x[0]).forEach(
                function (item) {
                    if (Object.keys(FS.assets).indexOf(item) != -1) {
                        value = value + FS.assets[item];
                    } else if (Object.keys(FS.liabilities).indexOf(item) != -1) {
                        value = value - FS.liabilities[item];
                    } else if (Object.keys(FS.equity).indexOf(item) != -1) {
                        value = value - FS.equity[item];
                    } else if (Object.keys(FS.revenues).indexOf(item) != -1) {
                        value = value - FS.revenues[item];
                    } else if (Object.keys(FS.expenses).indexOf(item) != -1) {
                        value = value + FS.expenses[item];
                    }
                }
            );
            FS.tags[tag] = value;
        }
    );
    
    delete FS.tags[''];
    
    // 出力
    return FS;
}

// 財務諸表の格納場所の作成
function initFS (Accounts) {
    let colnames = ["決算月", "決算日", "当期純利益"].concat( 
        ["assets", "liabilities", "equity", "revenues", "expenses"],
        Accounts,
        uniq(TAGLIST.map(x => x[1]).filter(x => x != ""))
    );
    let index = {};
    let body = [];
    
    // FSARRAY が存在する場合には、colnames を修正し、body の部分も変更する
    if (FSARRAY.body.length > 0) {
    
        // colnames の更新
        colnames = joinArray(colnames, FSARRAY.indexarray);
        
        // 元の body の行数によるループ
        
        for (let i = 0; i < FSARRAY.body.length; i++) {
        
            let bodyarray = [];
            
            // 新しい colnames の要素数 (新たな body の列数) によるループ
            for (let j = 0; j < colnames.length; j++) {
                // 新旧 colnames で共通する要素の取り扱い
                if (FSARRAY.indexarray.indexOf(colnames[j]) != -1) {
                    let n = FSARRAY.index[colnames[j]];
                    bodyarray.push(FSARRAY.body[i][n]);
                    
                // 元の colnames に含まれない要素の取り扱い
                } else {
                    bodyarray.push(0);
                }
            }
            
            body.push(bodyarray);
        }
    }
    
   
    for (let i = 0; i < colnames.length; i++) {
        index[colnames[i]] = i;
    }
    
    FSARRAY = {"index": index, "indexarray": colnames, "body": body};
}


// 財務諸表の数値を格納場所に記録
function registerFS (FS, YM) {
    // 入力情報の形成
    let accountdata = {};
    let elements = Object.keys(FS);
    elements.forEach(
        function (element) {
            accountdata = Object.assign(accountdata, FS[element]);
        }
    );
    
    // 入力行の作成
    let accounts = Object.keys(accountdata);
    let fsdata = [];
    FSARRAY.indexarray.forEach(
        function (account) {
            if (account == '決算月') {
                fsdata.push(YM);
            } else if (account == '決算日') {
                fsdata.push(new Date(Number(YM.toString().substr(0, 4)),
                    Number(YM.toString().substr(4)), 0));
            } else if (accounts.indexOf(account) != -1) {
                fsdata.push(accountdata[account]);
            } else {
                fsdata.push(0);
            }
        }
    );
    
    // 適用
    let flg = FSARRAY.body.map(x => x[FSARRAY.index['決算月']]).indexOf(YM);
    if (flg == -1) {
        FSARRAY.body.push(fsdata);
    } else {
        FSARRAY.body[flg] = fsdata;
    }
    
    // 決算月でソート
    arraySort(FSARRAY.body, FSARRAY.index['決算月']);
    
}

// Google データポータル用のデータベースの作成
function mkTidy () {
    let colnames = ["決算月", "決算日", "element", "account", "value", "ratio"];
    let index = {};
    let body = [];
   
    for (let i = 0; i < colnames.length; i++) {
        index[colnames[i]] = i;
    }
    
    
    // TIDY にデータの格納
    for (let i = 0; i < FSARRAY.body.length; i++) {
        Comp.forEach(
            function (account) {
                // bodyarray の宣言
                let bodyarray = [FSARRAY.body[i][FSARRAY.index["決算月"]], FSARRAY.body[i][FSARRAY.index["決算日"]]];
            
                // 財務諸表の要素
                bodyarray.push(fullComp[Comp.indexOf(account)][Def.index['element']]);
                    
                // 勘定科目
                bodyarray.push(account);
                
                // 金額
                let amount = FSARRAY.body[i][FSARRAY.index[account]];
                bodyarray.push(amount);
                
                // 比率の計算
                let basis;
                if (['assets', 'liabilities', 'equity'].indexOf(bodyarray[index['element']]) != -1) {
                    basis = FSARRAY.body[i][FSARRAY.index['assets']];
                } else if (FSARRAY.indexarray.indexOf('basis') != -1) {
                    basis = FSARRAY.body[i][FSARRAY.index['basis']] * (-1);
                } else {
                    basis = FSARRAY.body[i][FSARRAY.index['revenues']];
                }
                bodyarray.push(amount / basis)
                
                // bodyarray に出力
                body.push(bodyarray);
            }
        );
        
    }
    
    TIDY = {"index": index, "indexarray": colnames, "body": body};
}

// スプレッドシートに記録
function writeSpreadSheet () {
    // fs sheet
    let fsbody = [FSARRAY.indexarray].concat(FSARRAY.body);
    FSSHEET.clearContents();
    FSSHEET.getRange(1, 1, fsbody.length, fsbody[0].length).setValues(fsbody);
    
    // tidy sheet
    fsbody = [TIDY.indexarray].concat(TIDY.body);
    FSTIDY.clearContents();
    FSTIDY.getRange(1, 1, fsbody.length, fsbody[0].length).setValues(fsbody);
    
    /*
    // tag sheet
    fsbody = [['account', 'tag']].concat(TAGLIST);
    sheet = FSS.getSheetByName('tag');
    sheet.clearContents();
    sheet.getRange(1, 1, fsbody.length, fsbody[0].length).setValues(fsbody);
    */
}

// テーブルの作成 (FSARRAY を対象にする)
function mkTable () {
    // 何回ループするか決定する
    let n = 3;
    if (n > FSARRAY.body.length)
        n = FSARRAY.body.length;
    
    ['assets', 'liabilities', 'equity', 'revenues', 'expenses'].forEach(
        function (element) {
        
            // 変数定義
            let elemchr, basis;
            let earnings = '当期純利益';
            if (element == 'assets') {
                elemchr = '資産合計';
                basis = FSARRAY.index['assets'];
            } else if (element == 'liabilities') {
                elemchr = '負債合計';
                basis = FSARRAY.index['assets'];
            } else if (element == 'equity') {
                elemchr = '純資産合計';
                basis = FSARRAY.index['assets'];
            } else if (element == 'revenues') {
                elemchr = '収益合計';
                basis = FSARRAY.index['revenues'];
                if (FSARRAY.indexarray.indexOf('basis') != -1)
                    basis = FSARRAY.index['basis'];
            } else if (element == 'expenses') {
                elemchr = '費用合計';
                basis = FSARRAY.index['revenues'];
                if (FSARRAY.indexarray.indexOf('basis') != -1)
                    basis = FSARRAY.index['basis'];
            }
        
            // ヘッダ
            TABLELIST[element] = '<table><thead><tr><th />';
            for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                TABLELIST[element] = TABLELIST[element] + '<th>' +
                    FSARRAY.body[i][FSARRAY.index['決算月']] + '</th>';
            }
            TABLELIST[element] = TABLELIST[element] + '</tr></thead>';
            
            // フッタ
            TABLELIST[element] = TABLELIST[element] + 
                '<tfoot><tr><td class="left">' + elemchr + '</td>';
            for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                TABLELIST[element] = TABLELIST[element] + '<td class="right">' +
                    Number(FSARRAY.body[i][FSARRAY.index[element]]).toLocaleString() + '</td>';
            }
            TABLELIST[element] = TABLELIST[element] + '<tr><td />'
            for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                TABLELIST[element] = TABLELIST[element] + '<td class="right">(' +
                    Number(FSARRAY.body[i][FSARRAY.index[element]] / Math.abs(FSARRAY.body[i][basis]))
                        .toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2}) + ')</td>';
            }
            TABLELIST[element] = TABLELIST[element] + '</tr>';
            if (element == 'expenses') {
                TABLELIST[element] = TABLELIST[element] + '<tr><td class="left">当期純利益</td>';
                for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                    TABLELIST[element] = TABLELIST[element] + '<td class="right">' + 
                        Number(FSARRAY.body[i][FSARRAY.index[earnings]]).toLocaleString() + '</td>';
                }
                TABLELIST[element] = TABLELIST[element] + '<tr><td />'
                for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                    TABLELIST[element] = TABLELIST[element] + '<td class="right">(' +
                        Number(FSARRAY.body[i][FSARRAY.index[earnings]] / Math.abs(FSARRAY.body[i][basis]))
                            .toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2}) + ')</td>';
                }
                TABLELIST[element] = TABLELIST[element] + '</tr>';
            }
            TABLELIST[element] = TABLELIST[element] + '</tfoot>';
            
            // 本体
            TABLELIST[element] = TABLELIST[element] + '<tbody>';
            fullComp.filter(x => x[Def.index['element']] == element).map(x => x[Def.index['account']])
                .forEach(
                    function (account) {
                        TABLELIST[element] = TABLELIST[element] + 
                            '<tr><td class="left">' + account + '</td>';
                        for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                            TABLELIST[element] = TABLELIST[element] + '<td class="right">' +
                                Number(FSARRAY.body[i][FSARRAY.index[account]]).toLocaleString() + '</td>';
                        }
                        TABLELIST[element] = TABLELIST[element] + '<tr><td />'
                        for (let i = FSARRAY.body.length - 1; FSARRAY.body.length - i <= n; i--) {
                            TABLELIST[element] = TABLELIST[element] + '<td class="right">(' +
                                Number(FSARRAY.body[i][FSARRAY.index[account]] / Math.abs(FSARRAY.body[i][basis]))
                                    .toLocaleString(undefined, {style: 'percent', minimumFractionDigits: 2}) + ')</td>';
                        }
                        TABLELIST[element] = TABLELIST[element] + '</tr>';
                    }
                );
            TABLELIST[element] = TABLELIST[element] + '</tbody></table>';
        }
    );
}

/*
内部関数
*/

// 勘定科目の取得と金額の計算
function calcAccounts (monthdata, element) {
    let result = {};
    
    // 対象の要素に属する勘定科目名の取得
    let elembox = fullComp
        .filter(x => x[Def.index['element']] == element)
        .map(x => x[Def.index['account']]);
    
    // 各勘定の金額の計算
    // 貸方科目の場合
    if (['revenues', 'liabilities', 'equity'].indexOf(element) !== -1) {
        elembox.forEach(
            function (elem) {
                result[elem] = monthdata['body']
                    .filter(x => x[monthdata.index['account']] == elem)
                    .map(x => x[monthdata.index['amount']])
                    .reduce((sum, x) => sum + x, 0);
                if (result[elem] != 0) result[elem] = result[elem] * (-1);
            }
        );
    // 借方科目の場合
    } else {
        elembox.forEach(
            function (elem) {
                result[elem] = monthdata['body']
                    .filter(x => x[monthdata.index['account']] == elem)
                    .map(x => x[monthdata.index['amount']])
                    .reduce((sum, x) => sum + x, 0);
            }
        );
    }
    
    return result;
}

// 2 次元配列を {"index": [index の連想配列], "indexarray": [index], "body": [データの 2 次元配列]} に変換
// 1 次元配列が入力される場合エラーが出る
function convertArray (array) {
    let index = {};
    let body = array.slice(1);
    
    for (let i = 0; i < array[0].length; i++) {
        index[array[0][i]] = i;
    }
    
    return {"index": index, "indexarray": array[0], "body": body};
}

// Date を受け取ると、その時点でその月の何%が経過しているのかを計算する
function calcProgress (date) {
    day = date.getDate();
    lastday = new Date(date.getYear(), date.getMonth() + 1, 0).getDate();
    return day / lastday;
}

// 配列 A と配列 B の両方の要素を残しながら、新たな配列 C を作成する
// A を本体として、B にしか含まれない要素は A の末尾に追加していく形で実行する
// 財務諸表の colnames を更新するのに用いる (現行の勘定科目定義を A とする)
function joinArray (A, B) {
    let result = A;
    B.forEach(
        function (elem) {
            if (result.indexOf(elem) == -1) {
                result.push(elem);
            }
        }
    );
    return result;
}

// 2次元配列を並び替えする
function arraySort (ary, j) {
  ary.sort(
    function (a, b) {
      if (a[j] < b[j]) return -1;
      if (a[j] > b[j]) return 1;
      // 複数キーの場合、上 2 行と同様のコードをそのまま書けば良い
      // 降順の場合は不等号を逆にする
      return 0;
    }
  )
}

// 貸借対照表要素をストック額にする (当月のストック変化分に前月のストックを加える
function mkStocks (FS, YM) {
    // YM よりも小さな値であり、かつ、最も大きな値を抽出する
    let FSYM = FSARRAY.body.map(x => x[FSARRAY.index['決算月']])
        .filter(x => YM - x > 0);
        
    if (FSYM.length > 0) {
        FSYM = FSYM.reduce((a, b) => Math.max(a, b));
        let LASTDATA = FSARRAY.body.filter(x => x[FSARRAY.index['決算月']] == FSYM)[0];
        
        // 貸借対照表科目について適用
        ['assets', 'liabilities', 'equity'].forEach(
            function (element) {
                let accounts = Object.keys(FS[element]);
                accounts.forEach(
                    function (account) {
                        FS[element][account] = FS[element][account] + LASTDATA[FSARRAY.index[account]];
                    }
                );
            }
        );
    }
}

// 当月を YYYYMM という文字列で取得する
function thisMonth () {
    let date = new Date();
    let year = date.getFullYear();
    let month = date.getMonth() + 1;
    
    let toTwoDigits = function (num, digit) {
        num += '';
        if (num.length < digit) {
            num = '0' + num;
        };
        return num;
    };
    
    let yyyy = toTwoDigits(year, 4);
    let mm = toTwoDigits(month, 2);
    return yyyy + mm;
}

// tag list の作成
function mkTagList () {
    let taglist = [];
    for (let i = 0; i < fullComp.length; i++) {
        let tags = fullComp[i][Def.index['tag']].split(';');
        tags.forEach(
            function (tag) {
                taglist.push([Comp[i], tag]);
            }
        );
    }
    return taglist;
}

// 配列の重複を取り除く
function uniq(array) {
    const uniquedArray = [];
    for (const elem of array) {
        if (uniquedArray.indexOf(elem) < 0)
            uniquedArray.push(elem);
    }
    return uniquedArray;
}


function test() {
    // console.log(calcEarnings(refrectProgress(convertArray(getData(JNL.getSheetByName("202012"))), 2020, 11)));
    // console.log(calcProgress(today));
    // console.log(calcProgress(new Date("2020-12-31")));
    // console.log([1, 2, 3, 4, 5].filter(x => x > 5).reduce((sum, element) => sum + element, 0));
    // initFS(Comp);
    // let trg = convertArray(getData(JNL.getSheetByName("202012")));
    // let adjFS = calcFS(trg, 202011);
    // registerFS(adjFS, 202011);
    // reflectProgress(trg);
    // let pureFS = calcFS(trg, 202012);
    // console.log(pureFS);
    // registerFS(pureFS, 202012);
    // console.log(pureFS);
    // console.log(FSARRAY);
    // writeSpreadSheet();
    // initFS(Comp);
    // initFS(Comp);
    // writeSpreadSheet();
    // console.log(adjFS);
    // console.log(joinArray(["a", "b", "c"], ["a", "d", "c", "e"]));
    // console.log(fsbox);
    // console.log(registerFS(pureFS, fsbox, 202012));
    // console.log(MONTHFS);
    /*
    let Y = 202004;
    let monthjnl;
    monthjnl = JNL.getSheetByName(Y.toString()).getDataRange().getValues();
    while (Y <= 202012) {
        monthjnl = convertArray(JNL.getSheetByName(Y).getDataRange().getValues());
        MONTHFS = calcFS(monthjnl, Y);
        registerFS(MONTHFS, Y);
        writeSpreadSheet();
        console.log(Y);
        Y++;
    } 
    */
    /*
    let nodata = convertArray(FSS.getSheetByName('nodata').getDataRange().getValues());
    // let onlycolnames = convertArray(FSS.getSheetByName('onlycolnames').getDataRange().getValues());
    console.log(nodata);
    return;
    */
    mkTidy();
    console.log(TIDY);
}

