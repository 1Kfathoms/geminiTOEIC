function sendGeminiPromptAndWriteToForm(questionNumber, problemNum) {
  // settingシートのB1からAPIキー取得
  const apiKey = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('setting').getRange('B1').getValue();

  // ans[]をA, B, C, DからランダムにproblemNum個作成
  const choices = ['A', 'B', 'C', 'D'];
  let ans = [];
  for (let i = 0; i < problemNum; i++) {
    ans.push(choices[Math.floor(Math.random() * choices.length)]);
  }

  // ★ここでpartLabelを定義
  let partLabel = '';
  if (questionNumber === 5) {
    partLabel = 'Part 5';
  } else if (questionNumber === 6) {
    partLabel = 'Part 6';
  } else if (questionNumber === 7) {
    partLabel = 'Part 7 前半';
  } else if (questionNumber === 8) {
    partLabel = 'Part 7 後半';
  } else {
    partLabel = 'Part ?';
  }

  let prompt = '';
  if (questionNumber === 8) {
    // Part 7後半（複数文書）
    let answerStr = '';
    for (let i = 0; i < problemNum; i++) {
      answerStr += `${ans[i]}`;
      if (i < problemNum - 1) answerStr += ', ';
    }
    prompt = 
`指示（Instruction）:  
TOEIC Part 7の後半に相当する、複数文書(2個または3個)を用いた長文読解問題を作成してください。複数の文書（Eメール＋通知、記事＋レビュー、広告＋チャットなど）を組み合わせて、情報を比較・照合する設問を作成してください。難易度は実際のTOEICよりやや高めにしてください。

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

  } else if (questionNumber === 7) {
    // Part 7前半（単一文書）
    let answerStr = '';
    for (let i = 0; i < problemNum; i++) {
      answerStr += `${ans[i]}`;
      if (i < problemNum - 1) answerStr += ', ';
    }
    prompt = 
`指示（Instruction）:  
TOEIC Part 7の前半に相当する、単一文書（Eメール、通知、記事、広告、案内など）を用いた長文読解問題を作成してください。文書は1つのみとし、設問は本文の内容理解や情報検索を問うものにしてください。難易度は実際のTOEICよりやや高めにしてください。

【重要】
下記の「正解指定」に従って、各設問の正解を必ず指定された順番・内容にしてください。

【正解指定】
今回生成する${problemNum}問の設問の正解は、必ず
${answerStr}
の順番・内容になるようにしてください。
（例：問題数が4問で、正解が「A, D, B」と指定されたら、1問目の正解はA、2問目はD、3問目はB、4問目はCにしてください）

【出力形式】
必ず下記の出力例と全く同じ形式で出力してください。
余計な説明や補足は一切不要です。
---（半角ハイフン3つ）で区切り、
1つ目のブロックに単一文書、
2つ目以降のブロックに各設問（「1. 問題文」から「D.」まで）、
最後のブロックに正解（A/B/C/Dを1行ずつ）を出力してください。
ブロックは必ず「---」で区切ること。

【出力例（回答がA, D, B, Cの場合）】
Subject: Office Renovation Notice
From: Building Management <info@office.com>
Date: May 10

Dear Tenants,

Please be informed that the main lobby will undergo renovation from May 15 to May 20. During this period, access will be limited to the side entrance. We apologize for any inconvenience and appreciate your cooperation.

Best regards,
Building Management

---
1.  What is the main purpose of this notice?
A.  To announce a new security policy
B.  To inform about a maintenance schedule
C.  To request tenant feedback
D.  To introduce new staff
---
2.  When will the renovation take place?
A.  May 10–15
B.  May 12–17
C.  May 15–20
D.  May 20–25
---
3.  What should tenants do during the renovation?
A.  Use the main lobby as usual
B.  Use the side entrance
C.  Avoid the building
D.  Contact management for access
---
A
D
B
---`;
  } else if (questionNumber === 6) {
    // Part 6（長文穴埋め問題）
    let answerStr = '';
    for (let i = 0; i < problemNum; i++) {
      answerStr += `${ans[i]}`;
      if (i < problemNum - 1) answerStr += ', ';
    }
    prompt = 
`指示（Instruction）:  
TOEIC Part 6（長文穴埋め問題）に相当する英文メールや通知文を1つ作成し、文中に${problemNum}か所の空欄（___1.___ など）を設けてください。各空欄に対して設問を作成し、選択肢A～Dを用意してください。設問は空欄ごとに「What is the best choice to fill in the blank?」などの形式で出題してください。難易度は実際のTOEICよりやや高めにしてください。

【重要】
下記の「正解指定」に従って、各設問の正解を必ず指定された順番・内容にしてください。

【正解指定】
今回生成する${problemNum}問の設問の正解は、必ず
${answerStr}
の順番・内容になるようにしてください。
（例：問題数が4問で、正解が「C, C, B, B」と指定されたら、1問目の正解はC、2問目はC、3問目はB、4問目はBにしてください）

【出力形式】
必ず下記の出力例と全く同じ形式で出力してください。
余計な説明や補足は一切不要です。
---（半角ハイフン3つ）で区切り、
1つ目のブロックに長文（空欄付き）、
2つ目以降のブロックに各設問（「1. 問題文」から「D.」まで）、
最後のブロックに正解（A/B/C/Dを1行ずつ）を出力してください。
ブロックは必ず「---」で区切ること。

【出力例（回答がC, C, B, Bの場合）】
Subject: Update on Quarterly Staff Meeting Schedule

Dear Team,

As part of our ongoing effort to improve communication across departments, we will be adjusting the schedule of our quarterly staff meetings. These meetings are crucial to ensure that all team members are aligned with our company objectives and have the opportunity to raise questions or concerns.

Starting this quarter, the meetings will be held on the first Monday of every third month, rather than the last Friday. This change is intended to allow teams more time to __1.__ their monthly targets and share any early insights.

Additionally, we are pleased to introduce a new segment during the meeting where selected departments will present short case studies about their recent projects. This aims to __2.__ interdepartmental learning and foster a more collaborative environment.

Please be __3.__ to arrive at the meeting room at least ten minutes early so we can start promptly. If you are unable to attend in person, a video recording will be made available afterward.

We appreciate your continued engagement and look forward to your __4.__ in this revised format.

Best regards,
Samantha Lee
Corporate Communications Manager

---
1.  What is the best choice to fill in the blank?
A.  identify
B.  postpone
C.  evaluate
D.  oppose
---
2.  What word best completes the sentence in this context?
A.  hinder
B.  regulate
C.  promote
D.  disclose
---
3.  Which option best completes the sentence in terms of tone and formality?
A.  hesitant
B.  prepared
C.  reluctant
D.  eligible
---
4.  Choose the most appropriate word to end the email with a positive and professional tone.
A.  opposition
B.  participation
C.  dismissal
D.  interruption
---
C
C
B
B
---`;
  } else if (questionNumber === 5) {
    // Part 5（短文穴埋め問題）
    let answerStr = '';
    for (let i = 0; i < problemNum; i++) {
      answerStr += `${ans[i]}`;
      if (i < problemNum - 1) answerStr += ', ';
    }
    prompt =
`指示（Instruction）:  
TOEIC Part 5（短文穴埋め問題）に相当する設問を${problemNum}問作成してください。各設問は1文の英文で、空欄に入る最も適切な語句や表現を選ぶ形式にしてください。選択肢はA～Dの4つを用意し、実際のTOEICよりやや高めの難易度にしてください。

【重要】
下記の「正解指定」に従って、各設問の正解を必ず指定された順番・内容にしてください。

【正解指定】
今回生成する${problemNum}問の設問の正解は、必ず
${answerStr}
の順番・内容になるようにしてください。
（例：問題数が5問で、正解が「A, B, A, A, C」と指定されたら、1問目の正解はA、2問目はB、3問目はA、4問目はA、5問目はCにしてください）

【出力形式】
必ず下記の出力例と全く同じ形式で出力してください。
余計な説明や補足は一切不要です。
---（半角ハイフン3つ）で区切り、
【注意】最初のブロックは必ず空白のダミーブロック（何も書かない）として出力し、その後にQuestions and Options（設問＋選択肢）をまとめて記載してください。
最後のブロックに正解（A/B/C/Dを1行ずつ）を出力してください。
ブロックは必ず「---」で区切ること。

【出力例（回答がA, B, A, A, Cの場合）】

---

1.  The manager insisted that all reports be submitted no later than Friday, ______ ensuring timely review before the meeting.
A.  thus
B.  unless
C.  instead
D.  whereas
---
2.  The marketing team is developing a new campaign that ______ appeal to a broader demographic.
A.  ought
B.  might
C.  shall
D.  must
---
3.  The CEO expressed his appreciation for the employees’ dedication, particularly during the company's most ______ period.
A.  demanding
B.  demanded
C.  demands
D.  demand
---
4.  The updated software is compatible with most devices, ______ it may still experience glitches on older models.
A.  although
B.  provided
C.  whether
D.  because
---
5.  The vendor was asked to provide a ______ cost estimate before any purchases could be authorized.
A.  comprehensible
B.  comparable
C.  comprehensive
D.  competitive
---
A
B
A
A
C
---`;
  } else {
    throw new Error('questionNumberは5, 6, 7, 8のみ対応しています');
  }

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
      let htmlBody = '';
      // 設問ブロックの先頭番号（"1. "など）を除去
      const questionsHtml = questionBlocks.map((q, idx) => {
        // 先頭の「数字. 」を削除
        const cleaned = q.replace(/^\d+\.\s*/, '');
        return `<li style="margin-bottom:1em;"><b>Q${idx+1}</b><br>${cleaned.replace(/\n/g, '<br>')}</li>`;
      }).join('');

      if (questionNumber === 5) {
        htmlBody = `
          <h2>TOEIC ${partLabel} Questions and Options</h2>
          <div><strong>【Questions and Options】</strong><ul>${questionsHtml}</ul></div>
          <div><a href="${formUrl}"><b>【回答はこちら】</b></a></div>
        `;
      } else {
        const docHtml = docPart
          .split('==============')
          .map((block, idx) => `<div style="border:1px solid #ccc; margin:1em 0; padding:1em;"><b>Document${idx+1}</b><br>${block.replace(/\n/g, '<br>')}</div>`)
          .join('');
        htmlBody = `
          <h2>TOEIC ${partLabel} 問題文・設問・選択肢</h2>
          <div><strong>【Question】</strong>${docHtml}</div>
          <div><strong>【Questions and Options】</strong><ul>${questionsHtml}</ul></div>
          <div><a href="${formUrl}"><b>【回答はこちら】</b></a></div>
        `;
      }

      MailApp.sendEmail({
        to: email,
        subject: `TOEIC ${partLabel}`,
        body: `【問題文・設問・選択肢】\n\n${questionNumber === 5 ? questionBlocks.map((q, idx) => `Q${idx+1}\n${q.replace(/^\d+\.\s*/, '')}`).join('\n\n') : docPart + '\n\n==============\n\n' + questionBlocks.map((q, idx) => `Q${idx+1}\n${q.replace(/^\d+\.\s*/, '')}`).join('\n\n')}\n\n【回答はこちら】\n${formUrl}\n\n`,
        htmlBody: htmlBody
      });
    }
  });

  Logger.log('フォームURL: ' + formUrl);

  // sendGeminiPromptAndWriteToForm内
  // ★ ans[]をanswersシートに保存（answerLinesではなくansを保存）＋partLabelもG列に保存
  const answersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('answers');
  // G列（7列目）にpartLabelを追加
  answersSheet.appendRow([new Date(), ...ans, '', '', '', partLabel]); // 2列目～にans、G列にpartLabel

  // sendGeminiPromptAndWriteToForm内のメール送信後などに追加
  const formsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('forms');
  // 設問（質問文＋選択肢）も保存する
  const questionsText = questionBlocks.join('\n\n');
  formsSheet.appendRow([new Date(), docPart, questionsText]);
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

  // ★ partLabelをanswersシートG列から取得
  let partLabel = answersSheet.getRange(lastRow, 7).getValue();

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
  let feedbackPrompt = `以下はTOEIC ${partLabel}形式の問題文と設問、あなたの回答です。

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
    subject: `TOEIC ${partLabel} 採点・フィードバック`,
    body: resultText + '\n\n【Geminiによる解説】\n' + feedback, // テキスト形式
    htmlBody: resultText.replace(/\n/g, '<br>') + '<br><br><b>【Geminiによる解説】</b><br>' + feedback // HTML形式
  });
}

// 呼び出し例
//sendGeminiPromptAndWriteToForm(8, 3);
// 呼び出し例
//sendGeminiPromptAndWriteToForm(8, 3);
