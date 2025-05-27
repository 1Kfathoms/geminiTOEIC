function sendGeminiPromptAndWriteToForm(questionNumber, problemNum) {
  // settingシートのB1からAPIキー取得
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('B1').getValue();

  // ans[]をA, B, C, DからランダムにproblemNum個作成
  const choices = ['A', 'B', 'C', 'D'];
  let ans = [];
  for (let i = 0; i < problemNum; i++) {
    ans.push(choices[Math.floor(Math.random() * choices.length)]);
  }

  if (questionNumber === 8) {
    // プロンプト作成（出力例を明確に分離）
    let answerStr = '';
    for (let i = 0; i < problemNum; i++) {
      answerStr += `${ans[i]}`;
      if (i < problemNum - 1) answerStr += ', ';
    }
    const prompt = 
`指示（Instruction）:  
TOEIC Part 7の後半に相当する、複数文書を用いた長文読解問題を作成してください。複数の文書（Eメール＋通知、記事＋レビュー、広告＋チャットなど）を組み合わせて、情報を比較・照合する設問を作成してください。難易度は実際のTOEICよりやや高めにしてください。

【重要】
下記の「正解指定」に従って、各設問の正解を必ず指定された順番・内容にしてください。

【正解指定】
今回生成する${problemNum}問の設問の正解は、必ず
${answerStr}
の順番・内容になるようにしてください。
（例：3問の場合「D, C, D」と指定されたら、1問目の正解はD、2問目はC、3問目はDにしてください）

【出力形式】
必ず下記の出力例と全く同じ形式で出力してください。
余計な説明や補足は一切不要です。
---（半角ハイフン3つ）で区切り、
1つ目のブロックに複数文書（文書が複数の場合は文書間を「==============」で区切ること）、
2つ目以降のブロックに各設問（「1. 問題文」から「D.」まで）、
最後のブロックに正解（A/B/C/Dを1行ずつ）を出力してください。

【出力例（回答がD, C, Dの場合）】
Subject: Urgent Recall Notice - "AquaPure" Water Filter

Dear Valued Customer,

Due to a manufacturing defect affecting a limited batch of our AquaPure water filters (model APF-1200, serial numbers starting with X2347), we are initiating an urgent product recall.  Filters from this batch may contain microscopic particles that could potentially contaminate drinking water.  If you own an AquaPure APF-1200 filter with a serial number beginning with X2347, please cease using it immediately. You can find the serial number printed on the filter's base.

For a full refund or replacement, please visit our website at www.aquapurefilters.com/recall and complete the online form, providing your serial number and proof of purchase. You can also call our customer service hotline at 555-1212.  The recall process is expected to take no longer than 7 business days.

We sincerely apologize for any inconvenience this may cause and appreciate your prompt attention to this matter.

Sincerely,
The AquaPure Team

==============

AquaPure Water Filter Review: 5 Stars!

This filter is amazing!  The water tastes so much better than before. I've used it for over 6 months now, and it's still working like a charm. The flow rate is excellent, and it’s really easy to install. I highly recommend it. -Sarah J.

==============

AquaPure APF-1200 Water Filter - Unbeatable Value!

**Features:**

* Superior Filtration
* High Flow Rate
* Easy Installation
* Long-lasting Cartridge

Get yours today! www.aquapurefilters.com


---
1.  What is the primary reason for the AquaPure water filter recall?
A.  The filters are leaking.
B.  The filters are not compatible with all plumbing systems.
C.  The filters have a design flaw that reduces water flow.
D.  The filters may contain contaminants.
---
2.  According to the customer review, what is a positive aspect of the AquaPure APF-1200 water filter?
A.  It has a unique and stylish design.
B.  It is inexpensive compared to competitors.
C.  It provides superior water taste and flow.
D.  It is easy to clean and maintain.
---
3.  What information is missing from the recall notice that is mentioned in the product advertisement?
A.  The filter's dimensions and weight.
B.  The length of the filter’s warranty.
C.  The lifespan of the filter cartridge.
D. The retail price of the water filter.


---
D
C
D
---`;

    // Gemini APIリクエスト
    const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
    const headers = { 'Content-Type': 'application/json' };
    const payload = JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }]
    });
    const options = {
      method: 'post',
      headers: headers,
      payload: payload,
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());

    // Geminiの返答から本文のみ抽出
    let geminiText = '';
    try {
      geminiText = json.candidates[0].content.parts[0].text;
    } catch (e) {
      geminiText = response.getContentText();
    }

    // 改行コードを正しく反映
    geminiText = geminiText.replace(/\\n/g, '\n');

    // --- で分割
    const split = geminiText.split(/^-{3,}$/m);
    const docPart = split[0] ? split[0].trim() : '';
    // 設問部分は2番目から正解直前まで
    let questionBlocks = [];
    let answerLines = [];
    if (split.length >= 3) {
      // 設問部分を---で分割し、各設問ごとに抽出
      questionBlocks = split.slice(1, split.length - 1).map(s => s.trim()).filter((s, idx) => idx < problemNum);
      // 最後の---以降が正解
      answerLines = split[split.length - 1].trim().split('\n').filter(l => l.match(/^[A-D]$/));
    }

    // ★ Geminiの出力をコンソール（Logger）に表示
    Logger.log('Gemini出力:\n' + geminiText);

    // mailシートのB列からメールアドレス取得（2行目以降）
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('mail');
    const lastRow = sheet.getLastRow();
    let emails = [];
    if (lastRow > 1) {
      emails = sheet.getRange(2, 2, lastRow - 1, 1).getValues().flat();
    }

    // フォームIDをsettingシートのB2から取得（フォーム自体は変更しない）
    const formId = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('B2').getValue();
    const form = FormApp.openById(formId);
    const formUrl = form.getPublishedUrl();

    // 問題文・設問・選択肢をメール送信
    emails.forEach(function(email) {
      if (email) {
        // 複数文書の区切りをHTMLで装飾
        const docHtml = docPart
          .split('==============')
          .map((block, idx) => `<div style="border:1px solid #ccc; margin:1em 0; padding:1em;"><b>Document${idx+1}</b><br>${block.replace(/\n/g, '<br>')}</div>`)
          .join('');

        // 設問・選択肢をリスト化
        const questionsHtml = questionBlocks.map((q, idx) =>
          `<li style="margin-bottom:1em;"><b>Q${idx+1}</b><br>${q.replace(/\n/g, '<br>')}</li>`
        ).join('');

        const htmlBody = `
          <h2>TOEIC Part7 問題文・設問・選択肢</h2>
          <div><strong>【Question】</strong>${docHtml}</div>
          <div><strong>【Questions and Options】</strong><ul>${questionsHtml}</ul></div>
          <div><a href="${formUrl}"><b>【回答はこちら】</b></a></div>
        `;

        MailApp.sendEmail({
          to: email,
          subject: 'TOEIC Part7',
          body: `【問題文・設問・選択肢】\n\n${docPart}\n\n==============\n\n${questionBlocks.join('\n\n')}\n\n【回答はこちら】\n${formUrl}\n\n`,
          htmlBody: htmlBody
        });
      }
    });

    Logger.log('フォームURL: ' + formUrl);

    // ★ ans[]をanswersシートに保存（answerLinesではなくansを保存）
    const answersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
    answersSheet.appendRow([new Date(), ...ans]);

    // sendGeminiPromptAndWriteToForm内のメール送信後などに追加
    const formsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forms');
    // 設問（質問文＋選択肢）も保存する
    const questionsText = questionBlocks.join('\n\n');
    formsSheet.appendRow([new Date(), docPart, questionsText]);
  }
}

// onFormSubmitで自動採点・Geminiでフィードバック
function onFormSubmit(e) {
  // e.valuesの中身をLoggerで確認
  Logger.log(JSON.stringify(e.values));

  // 例: [タイムスタンプ, メール, Q1, Q2, Q3] の場合
  const email = e.values[1];
  const userAnswers = [e.values[2], e.values[3], e.values[4]];

  // 正解取得（answersシートの最新行＝ans[]）
  const answersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
  const lastRow = answersSheet.getLastRow();
  const correctAnswers = answersSheet.getRange(lastRow, 2, 1, userAnswers.length).getValues()[0];

  // 採点
  let score = 0;
  let resultText = '';
  for (let i = 0; i < correctAnswers.length; i++) {
    if (userAnswers[i] === correctAnswers[i]) score++;
    resultText += `Q${i+1}: あなたの回答:${userAnswers[i]} / 正解:${correctAnswers[i]}\n`;
  }
  resultText = `スコア: ${score} / ${correctAnswers.length}\n\n` + resultText;

  // 問題文・設問取得
  const formsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forms');
  const lastFormRow = formsSheet.getLastRow();
  const problemText = formsSheet.getRange(lastFormRow, 2).getValue(); // 本文
  const questionsText = formsSheet.getRange(lastFormRow, 3).getValue(); // 設問＋選択肢

  // Geminiへのプロンプト作成
  let feedbackPrompt = `以下はTOEIC Part7形式の問題文と設問、あなたの回答です。

【問題文】
${problemText}

【設問・選択肢】
${questionsText}

【あなたの回答】
`;
for (let i = 0; i < userAnswers.length; i++) {
  feedbackPrompt += `Q${i+1}: ${userAnswers[i]}（正解: ${correctAnswers[i]}）\n`;
}
feedbackPrompt += `
あなたの回答と正解を比較し、各設問ごとに
・正しい回答とその根拠（本文から導ける理由）
・本文の和訳（対訳形式で）
・難易度の高い語彙とその日本語解説
を日本語でまとめてください。
【重要】単なる解説ではなく、「あなたの回答」を分析した個別のアドバイスになるよう注意してください。
出力は必ずHTML形式（h2, strong, ul, li, p など）で返してください。マークダウンは使わないでください。`;

  // Gemini API呼び出し
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('B1').getValue();
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=${apiKey}`;
  const headers = { 'Content-Type': 'application/json' };
  const payload = JSON.stringify({
    contents: [{ parts: [{ text: feedbackPrompt }] }]
  });
  const options = {
    method: 'post',
    headers: headers,
    payload: payload,
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, options);
  let feedback = '';
  try {
    const json = JSON.parse(response.getContentText());
    feedback = json.candidates[0].content.parts[0].text;
  } catch (e) {
    feedback = response.getContentText();
  }

  // 解説メール送信
  MailApp.sendEmail({
    to: email,
    subject: 'TOEIC Part7 採点・フィードバック',
    body: resultText + '\n\n【Geminiによる解説】\n' + feedback, // テキスト形式
    htmlBody: resultText.replace(/\n/g, '<br>') + '<br><br><b>【Geminiによる解説】</b><br>' + feedback // HTML形式
  });
}

// 呼び出し例
//sendGeminiPromptAndWriteToForm(8, 3);
// 呼び出し例
//sendGeminiPromptAndWriteToForm(8, 3);
