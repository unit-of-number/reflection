

// ====== 設定項目：ここを先生が自由に調整してください ======
const REFLECTION_COLUMN = 7;  // G列: ふりかえりの内容
const SCORE_COLUMN = 8;       // H列: 総合評価
const KNOW_COLUMN = 9;        // I列: もっと知りたい
const LISTEN_COLUMN = 10;     // J列: もっと聴きたい
const TELL_COLUMN = 11;       // K列: もっと伝えたい
const COMMENT_COLUMN = 12;    // L列: コメント
// =======================================================

/**
 * --- ★★★ 変更点 ★★★ ---
 * メニューに「コメント返信」を追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('自動評価')
    .addItem('未評価の振り返りをすべて評価', 'evaluateUnevaluatedRows')
    .addSeparator() // 区切り線
    .addItem('選択行のコメントを児童に返信する', 'sendFeedbackToStudent') // コメント返信メニュー
    .addToUi();
 }
 
 
 /**
 * 未評価の行を自動で探し出して評価するメイン関数
 */
 function evaluateUnevaluatedRows() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = range.getValues();
 
  for (let i = 0; i < values.length; i++) {
    const reflectionText = values[i][REFLECTION_COLUMN - 1];
    const currentScore = values[i][SCORE_COLUMN - 1];
    const currentRow = i + 2;
 
    if (reflectionText && !currentScore) {
      const knowResult = calculateScore('KNOW', reflectionText);
      const listenResult = calculateScore('LISTEN', reflectionText);
      const tellResult = calculateScore('TELL', reflectionText);
      const totalScore = Math.round((knowResult.score + listenResult.score + tellResult.score) / 3);
      const comment = generateComment(knowResult, listenResult, tellResult, reflectionText);
 
      sheet.getRange(currentRow, SCORE_COLUMN).setValue(totalScore);
      sheet.getRange(currentRow, KNOW_COLUMN).setValue(knowResult.score);
      sheet.getRange(currentRow, LISTEN_COLUMN).setValue(listenResult.score);
      sheet.getRange(currentRow, TELL_COLUMN).setValue(tellResult.score);
      sheet.getRange(currentRow, COMMENT_COLUMN).setValue(comment);
    }
  }
  SpreadsheetApp.getUi().alert('未評価の振り返りの評価が完了しました。');
 }
 
 /**
 * 【改善版】評価パターン
 */
 const RUBRIC_PATTERNS = {
  KNOW: {
    LEVEL5: [/もし(.+)だったらどうなるだろうか？/i, /(.+)だけでなく(.+)についても知りたい/i, /自分なりに(.+)を調べてみたい/i, /次は(.+)を試したい/i, /別の方法はないか/i, /共通点は何だろう/i, /法則を見つけた/i],
    LEVEL4: [/なぜ(.+)なのだろうか？/i, /どうして(.+)なのか/i, /(.+)と(.+)の違いはなんだろう/i, /さらに詳しく/i, /原因は/i],
    LEVEL3: [/気になった/i, /わかったことは(.+)です/i, /調べてみたら/i, /面白いと思った/i],
    LEVEL2: [/〜って何？/i, /教えて/i, /難しかった/i]
  },
  LISTEN: {
    LEVEL5: [/みんなの意見を聞いて(.+)という新しい視点に気づいた/i, /(.+)さんの考えの面白いところは(.+)だ/i, /自分の考えと(.+)さんの考えを比べると/i, /対話を通して考えが深まった/i],
    LEVEL4: [/どうして(.+)さんはそう考えたのかな/i, /(.+)さんの意見に質問があります/i, /なるほど、そういう考え方もあるのか/i, /(.+)という意見が多かった/i],
    LEVEL3: [/(.+)さんの発表を聞いて/i, /グループの友達が/i, /みんなの意見は/i],
    LEVEL2: [/〇〇さんが言っていた/i, /発表を聞いた/i]
  },
  TELL: {
    LEVEL5: [/私の考えは(.+)です。なぜなら(.+)だからです/i, /(.+)という理由から(.+)と結論付けました/i, /図や表を使って説明したい/i, /みんなに(.+)ということを伝えたい/i],
    LEVEL4: [/私が一番伝えたいのは(.+)ということです/i, /理由は(.+)です/i, /具体例を挙げると/i, /つまり/i],
    LEVEL3: [/私の意見は(.+)です/i, /考えたことを発表した/i, /〜と思います/i],
    LEVEL2: [/〜だと思った/i, /〜でした/i]
  }
 };
 
 /**
 * 【改善版】評価ロジック
 */
 function calculateScore(aspect, text) {
  let score = 1;
  let matched = [];
  const patterns = RUBRIC_PATTERNS[aspect];
  for (let level = 5; level >= 2; level--) {
    const regexList = patterns[`LEVEL${level}`];
    for (const regex of regexList) {
      const matches = text.match(regex);
      if (matches) {
        score += (level - 3) * 0.75;
        matched.push(matches[0]);
      }
    }
  }
  const logicalConnectors = (text.match(/なぜなら|しかし|だから|さらに|つまり/g) || []).length;
  score += logicalConnectors * 0.5;
  score = Math.max(1, Math.min(score, 5));
  return { score: Math.round(score), matched: [...new Set(matched)] };
 }
 
 /**
 * 【改善版】コメント生成
 */
 function generateComment(know, listen, tell, text) {
  let comment = "";
  if (know.score >= 4) {
    comment += `「${know.matched.join('」「')}」といった記述から、自分の問いを立てて探求しようとする強い意欲が伝わってきます。素晴らしいですね！`;
    comment += `その疑問を解決するために、次は図書室で関連する本を探したり、違う条件で実験したりするなど、具体的な行動に繋げていきましょう。`;
  } else if (know.score >= 3) {
    comment += `学習内容に関心を持ち、「${know.matched.join('」「')}」と感じられたのは良い視点です。`;
    comment += `そこから「なぜそうなるのか？」「他にはどんな例があるか？」ともう一歩踏み込んで考えてみると、さらに学びが面白くなりますよ。`;
  } else {
    comment += `学習した内容を丁寧に振り返ることができています。まずは「一番心に残ったこと」や「少し不思議に思ったこと」を見つけることから始めると、自分の問いが見つかりやすくなります。`;
  }
  if (listen.score >= 4) {
    comment += " また、友達の意見から新しい視点を見つけ、自分の考えを深めようとする姿勢も素敵です。";
  }
  if (tell.score >= 4) {
     comment += " 理由や根拠を明確にして、自分の考えを分かりやすく伝えようとしている点も素晴らしいです。";
  }
  return comment.trim();
 }

// --- ★★★ ここから追記 ★★★ ---

/**
 * 選択されている行のコメントを、該当する児童のスプレッドシートに書き込む関数
 */
function sendFeedbackToStudent() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRow = sheet.getActiveRange().getRow();
  
  if (activeRow < 2) {
    SpreadsheetApp.getUi().alert("見出し行は選択できません。");
    return;
  }
  
  // 1. 選択行から必要な情報を取得
  // 144行目あたり
// B列の名前と、配布ファイル名を合体させて、探すファイル名を自動生成する
const studentFileName = sheet.getRange(activeRow, 2).getValue() + " - " + "【3年1組】児童用ふりかえり入力シート";
  const timestamp = sheet.getRange(activeRow, 1).getValue();   // A列: 提出日時
  const reflection = sheet.getRange(activeRow, 7).getValue();  // G列: ふりかえり
  const comment = sheet.getRange(activeRow, 12).getValue();    // L列: コメント
  
  if (!comment) {
    SpreadsheetApp.getUi().alert("この行には返信するコメントがありません。");
    return;
  }
  
  try {
    // 2. 児童のファイルを見つける
    const files = DriveApp.getFilesByName(studentFileName);
    if (!files.hasNext()) {
      throw new Error(`「${studentFileName}」が見つかりませんでした。ファイル名を確認してください。`);
    }
    const studentFile = files.next();
    const studentSpreadsheetId = studentFile.getId();
    
    // 3. 児童のスプレッドシートを直接操作してコメントを書き込む
const studentSpreadsheet = SpreadsheetApp.openById(studentSpreadsheetId);
const sheetName = "ふりかえり記録"; // 児童側の記録シート名
let sheet = studentSpreadsheet.getSheetByName(sheetName);

if (!sheet) {
  throw new Error(`児童のシート内に「${sheetName}」が見つかりませんでした。`);
}

// G列（7列目）に見出しがなければ追加
if (sheet.getRange("G1").getValue() !== "先生からのコメント") {
  sheet.getRange("G1").setValue("先生からのコメント");
}

const data = sheet.getDataRange().getValues();
const targetTimestamp = new Date(timestamp).getTime();

// 記録シートを1行ずつチェックして、該当するふりかえりを見つける
let found = false;
for (let i = 1; i < data.length; i++) {
  const rowTimestamp = new Date(data[i][0]).getTime();
  const rowReflection = data[i][5]; // F列のふりかえり内容

  // 提出日時とふりかえり内容が一致する行を探す
  if (rowTimestamp === targetTimestamp && rowReflection === reflection) {
    sheet.getRange(i + 1, 7).setValue(comment); // G列にコメントを書き込む
    found = true;
    break; // 見つかったらループを抜ける
  }
}

if (!found) {
  throw new Error("児童の記録シート内で、該当するふりかえりが見つかりませんでした。");
}
    SpreadsheetApp.getUi().alert(`「${sheet.getRange(activeRow, 2).getValue()}」さんにコメントを送信しました。`);

  } catch (e) {
    SpreadsheetApp.getUi().alert("エラーが発生しました: " + e.message);
    Logger.log(e);
  }
}
