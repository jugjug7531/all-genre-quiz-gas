//-------------------------
//【暫定】各処理に必要なセル決め打ち
//問題一覧セル
const questionCells = "E2:K12"
//正解者の名前のプルダウンセル
const answererNameCell = "C14"
//問題ジャンルのセル
const correctAnswerKindCell = "C15"
//正解問題のセル
const correctAnswerCell = "C16"
//得点定義の一覧セル
const defScoreSells = "F15:F19"
//色の一覧セル
const colorCells = "B2:B12" //ラベル含む
const copyColorSeikaiCountCells = "B23:B33"
const copyColorScoreCells = "B36:B46"
//参加者名の一覧セル
const nameCells = "C2:C12" //ラベル含む
//各参加者のジャンルごとの正解数カウントセル
const seikaiCountCells = "E24:K33" //ラベル含まない
//得点表セル
const scoreCells = "E37:K46" //ラベル含まない
//残り問題数のセル
const QuestionRemainNum = "K15"
//-------------------------

//転置配列作成用ライブラリ
const _ = Underscore.load();

/* ゲームスタート */
function gameStart(){
  //現在のスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //現在のシートを取得
  const sheet = spreadsheet.getActiveSheet();

  //確認ダイアログ作成
  let ui = SpreadsheetApp.getUi();
  let response = ui.alert("確認","新しいゲームを開始しますか？\n得点はすべてリセットされます。", ui.ButtonSet.OK_CANCEL);
  //キャンセルなら終了
  if(response == "CANCEL"){
    return;
  }

  //参加者の背景色を得点表左側にコピー
  sheet.getRange(colorCells).copyTo(sheet.getRange(copyColorSeikaiCountCells));
  sheet.getRange(colorCells).copyTo(sheet.getRange(copyColorScoreCells));
  //得点&背景色リセット
  sheet.getRange(questionCells).setBackground(null);
  sheet.getRange(seikaiCountCells).clearContent();
  sheet.getRange(scoreCells).clearContent();
}

/* 正解者および正解問題の記録＆得点表更新 */
function Record(){
  //現在のスプレッドシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  //現在のシートを取得
  const sheet = spreadsheet.getActiveSheet();

  //正解問題を正解者の背景色で塗りつぶす
  RecordPersonAnswerdCorrectly(sheet);
  //得点更新
  UpdateScore(sheet);
  //問題終了ならダイアログ表示
  if(sheet.getRange(QuestionRemainNum).getValue() <= 0 ){
    let ui = SpreadsheetApp.getUi();
    let response = ui.alert("問題終了！","規定数の問題を読み上げました。\nお疲れさまでした！", ui.ButtonSet.OK);
  }
}

/* 正解数＆得点更新 */
function UpdateScore(sheet){
  /* 正解数を更新 */
  //正解数記入用の配列作成
  let arrSeikaiNum = sheet.getRange(seikaiCountCells).getValues();
  //色の配列を取得
  let arrColor = sheet.getRange(colorCells).getBackgrounds();
  //配列転置
  let arrColorTrans = _.zip.apply(_, arrColor);
  //参加者の背景色を取得
  let targetColors = arrColorTrans[0];
  targetColors.shift()
  //色の配列要素の黒色は削除
  for(let i=0; i<targetColors.length; i++){
    if(targetColors[i].toUpperCase() == "#FFFFFF"){
      targetColors.splice(i, targetColors.length - i);
      break;
    }
  }
  //カウント対象範囲の背景色のカラーコードを二次元配列で取得する
  let bgColors = sheet.getRange(questionCells).getBackgrounds();
  //配列転置
  let bgColorsTrans = _.zip.apply(_, bgColors);
  //参加者ごとの各ジャンル正解数カウント
  let count = 0;
  for(let p = 0; p < targetColors.length; p++){ //人数
    //各ジャンルの正解数カウント
    for(let r = 0; r < bgColorsTrans.length; r++){ //ジャンル数
      for(let c = 0; c < bgColorsTrans[r].length; c++) { //ジャンルごとの問題数
        //toUpperCaseメソッドを使ってどちらのカラーコードも大文字にする
        if(bgColorsTrans[r][c].toUpperCase() == targetColors[p].toUpperCase()){
          count++
        }
      }
      //ジャンル正解数を格納
      arrSeikaiNum[p][r] = count;
      count=0;
    }
  }
  //正解数を記入
  sheet.getRange(seikaiCountCells).setValues(arrSeikaiNum);

  /* 得点を更新 */
  //得点記入用の配列作成
  let arrScore = sheet.getRange(scoreCells).getValues();
  const arrDefScore = sheet.getRange(defScoreSells).getValues();
  //配列転置
  const tmpArrDefScoreTrans = _.zip.apply(_, arrDefScore);
  const arrDefScoreTrans = tmpArrDefScoreTrans[0];
  //参加者ごとの各ジャンル得点カウント
  let score = 0;
  for(let p = 0; p < targetColors.length; p++){ //人数
    //各ジャンルの得点カウント
    for(let r = 0; r < bgColorsTrans.length; r++){ //ジャンル数
      //正解数に応じて得点加算
      if(arrSeikaiNum[p][r] != 0){
        for(let n = 0; n < arrSeikaiNum[p][r] && n < arrDefScore.length; n++){
          score += arrDefScoreTrans[n];
        }
      }
      //得点を格納
      arrScore[p][r] = score;
      score=0;
    }
  }
  //得点を記入
  sheet.getRange(scoreCells).setValues(arrScore);
}

/* 正解問題を正解者の背景色で塗りつぶす */
function RecordPersonAnswerdCorrectly(sheet) {
  //正解者の背景色を取得
  let answererColor = getCorrectAnswererColor(sheet);
  //正解問題の行と列を取得
  let correctAnswerCell = getCorrectAnswerPos(sheet);
  //選択セルの背景色を変更する
  sheet.getRange(correctAnswerCell.getRow(),correctAnswerCell.getColumn()).setBackground(answererColor);
}

/* 正解問題の行と列を取得 */
function getCorrectAnswerPos(sheet){
  //正解問題を取得
  let correctAnswerKind = sheet.getRange(correctAnswerKindCell).getValue();
  //正解問題を取得
  let correctAnswer = sheet.getRange(correctAnswerCell).getValue();
  //問題一覧配列の作成
  let arrQuestion = sheet.getRange(questionCells).getValues();
  //配列転置
  let arrQuestionTrans = _.zip.apply(_, arrQuestion);
  //問題一覧配列の0行目から問題ジャンルの列番号取得
  let answererKindIndex = arrQuestion[0].indexOf(correctAnswerKind);
  //問題一覧の転置配列から正解問題のインデックス番号取得
  let answererIndex = arrQuestionTrans[answererKindIndex].indexOf(correctAnswer);
  //問題一覧セル内の正解問題のセルを返す（行と列はインデックス番号+1する）
  let cell = sheet.getRange(questionCells).getCell(answererIndex+1, answererKindIndex+1);

  return cell;
}

/* 正解者の背景色を取得 */
function getCorrectAnswererColor(sheet) {
  //正解者の名前を取得
  let answererName = sheet.getRange(answererNameCell).getValue();
  //色と名前の配列を取得
  let arrColor = sheet.getRange(colorCells).getBackgrounds();
  let arrName = sheet.getRange(nameCells).getValues();
  //配列転置
  let arrColorTrans = _.zip.apply(_, arrColor);
  let arrNameTrans = _.zip.apply(_, arrName);
  //名前配列から正解者のインデックス番号取得
  let answererIndex = arrNameTrans[0].indexOf(answererName);

  //正解者の背景色取得
  return arrColorTrans[0][answererIndex];
}
