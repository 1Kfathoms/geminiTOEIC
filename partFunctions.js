/*
  【自動配信仕様】
  ・毎日6〜7時に Part 5・Part 6・Part 7前半 のいずれかをランダムでメール送信する
    （runRandomPart5or6or7Single を利用）
  ・毎日17〜18時に Part 7後半 をメール送信する
    （runPart7Multi を利用）
  ※トリガー設定はGoogle Apps Scriptの時間主導型トリガーで行う
*/

// Part 5（短文穴埋め問題）を出題
function runPart5() {
  // 5問固定
  const problemNum = 5;
  sendGeminiPromptAndWriteToForm(5, problemNum);
}

// Part 6（長文穴埋め問題）を出題
function runPart6() {
  // 4問固定（TOEIC Part 6標準）
  const problemNum = 4;
  sendGeminiPromptAndWriteToForm(6, problemNum);
}

// Part 7 前半（単一文書読解）を出題
function runPart7Single() {
  // 2～4問のランダム
  const problemNum = Math.floor(Math.random() * 3) + 2;
  sendGeminiPromptAndWriteToForm(7, problemNum);
}

// Part 7 後半（複数文書読解）を出題
function runPart7Multi() {
  // 4または5問のランダム
  const problemNum = Math.random() < 0.5 ? 4 : 5;
  sendGeminiPromptAndWriteToForm(8, problemNum);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('study')
    .addItem('Part 5', 'runPart5')
    .addItem('Part 6', 'runPart6')
    .addItem('Part 7 前半', 'runPart7Single')
    .addItem('Part 7 後半', 'runPart7Multi')
    .addToUi();
}

// Part 5・Part 6・Part 7前半のいずれかをランダムで出題
function runRandomPart5or6or7Single() {
  const funcs = [runPart5, runPart6, runPart7Single];
  const idx = Math.floor(Math.random() * funcs.length);
  funcs[idx]();
}