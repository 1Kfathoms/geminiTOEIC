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
（例：問題数が5問で、正解が「C, C, D, B, D」と指定されたら、1問目の正解はC、2問目はC、3問目はD、4問目はB、5問目はDにしてください）

【出力形式】
必ず下記の出力例と全く同じ形式で出力してください。
余計な説明や補足は一切不要です。
---（半角ハイフン3つ）で区切り、
1つ目のブロックに複数文書（文書が複数の場合は必ず文書間を「==============」で区切ること）、
2つ目以降のブロックに各設問（「1. 問題文」から「D.」まで）、
最後のブロックに正解（A/B/C/Dを1行ずつ）を出力してください。
ブロックは必ず「---」で区切ること。

【出力例（回答がC, C, D, B, Dの場合）】
Subject: Follow-up on Vendor Fair Booth Setup
From: Clara Healy <clara.healy@harmonypublishing.com>
To: All Marketing Team Members
Date: June 7

Dear Team,

Thank you again for your efforts in preparing for next week’s Vendor Fair. Please ensure that all promotional materials, including the new line of eco-friendly bookmarks and sample copies of our summer releases, are shipped to the venue by Monday at noon. The event staff will begin booth setup at 2:00 p.m., and delays may result in losing our reserved location.

Also, for those volunteering to assist with setup, please arrive at the convention center by 1:30 p.m. sharp. A short orientation will be conducted in the lobby. Lastly, don’t forget that the revised schedule of author signings is now posted on the shared drive under “VendorFair2025.” Please review it and let me know if there are any conflicts with your assigned tasks.

Best,
Clara

==============
Title: Vendor Fair Kicks Off Next Week with a Focus on Sustainability
Published by: The City Herald
Date: June 7

The annual Vendor Fair at the Midtown Convention Center is set to begin next Tuesday, featuring over 120 local businesses. This year’s theme is “Sustainability in Business,” highlighting eco-conscious products and practices. Attendees can look forward to free samples, interactive exhibits, and author signings.

Event coordinator Jeremy Lin noted that several publishing companies are expected to present new titles and green-themed promotional items. “We’ve received confirmation from Harmony Publishing, among others, who are showcasing a new line of biodegradable bookmarks,” said Lin.

The event will open to the public at 10:00 a.m. on Tuesday, with setup allowed only on Monday afternoon. Author signings will be spread across all three days of the fair, with final timing updates posted on the event app this Friday.

---
1.  What is the most likely reason Clara asked the team to send materials by noon on Monday?
A.  Because she wants time to double-check the contents
B.  Because volunteers need them for the orientation session
C.  Because the venue starts allowing setup in the afternoon
D.  Because the author signings start right after setup
---
2.  According to both documents, what is true about Harmony Publishing’s participation?
A.  They will distribute a new set of printed catalogues
B.  They will be one of the event’s primary sponsors
C.  They will present items aligned with the fair’s theme
D.  They will offer early access to summer releases online
---
3.  Which task requires access to the shared drive?
A.  Sending promotional materials
B.  Confirming arrival times for volunteers
C.  Checking booth layout
D.  Reviewing the author signing schedule
---
4.  What can be inferred about the event’s schedule?
A.  Author signings will begin on Monday
B.  Setup is only permitted during a specific time
C.  Book sales will not take place on the first day
D.  The orientation will last several hours
---
5.  Based on the information provided, what is one possible risk for Harmony Publishing?
A.  Having their products removed by the event staff
B.  Having a smaller booth than planned
C.  Being unable to participate in author signings
D.  Being reassigned to a different booth space
---
C
C
D
B
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
      // 不足している場合は空文字で埋める
      while (questionBlocks.length < problemNum) questionBlocks.push('');
      // 最後の---以降が正解
      answerLines = split[split.length - 1].trim().split('\n').filter(l => l.match(/^[A-D]$/));
      // 不足している場合は空文字で埋める
      while (answerLines.length < problemNum) answerLines.push('');
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

  // 例: [タイムスタンプ, メール, Q1, Q2, ...] の場合
  const email = e.values[1];
  // 2列目（index=1）がメール、それ以降が解答
  const userAnswers = e.values.slice(2);

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
【重要】
・単なる解説ではなく、「あなたの回答」を分析した個別のアドバイスになるよう注意してください。
・出力は必ずHTML形式（h2, strong, ul, li, p など）で返してください。マークダウンは使わないでください。
・適切に改行を行い、解説が見やすいようにしてください。`;

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
