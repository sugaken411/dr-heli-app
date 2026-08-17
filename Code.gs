const GAS_VERSION = "v3.1"; // メール+パスワード認証必須化（スタッフ用PIN廃止、管理者PINは非公開の緊急回復用のみ）
const PIN_ADMIN = "9999"; // 🌟 非公開の緊急回復用。画面上には一切表示・案内しない
const GAS_API_URL = "https://script.google.com/macros/s/AKfycbzdmgxL3GL-x7sANo05V4nujuZ9CzKTZIuQ-KMNJlawOAJdcMTMZH37c4S0xdSXRFnr/exec";
const LIBRARY_DB_ID = "17ejBS_Uq6cWxkagnFQknfycbMGnoaV2q7234U5Pwqnc";

// 🌟 認証まわりのヘルパー群
// トークン署名用の秘密鍵はコードに書かず、スクリプトプロパティ(非公開領域)に自動生成して保存する
function getAuthSecret() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('AUTH_SECRET');
  if (secret) return secret;
  // 🌟 初回生成時に複数リクエストが同時に走ると、それぞれ別のシークレットを生成して
  // 上書きし合い、片方で署名されたトークンが以後ずっと検証不能になる（2026-08-15、実際に発生）。
  // ロックで生成〜保存をアトミックにする。
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    secret = props.getProperty('AUTH_SECRET');
    if (!secret) {
      secret = Utilities.getUuid() + Utilities.getUuid();
      props.setProperty('AUTH_SECRET', secret);
    }
    return secret;
  } finally {
    lock.releaseLock();
  }
}

function bytesToHex(bytes) {
  return bytes.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function hashPassword(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(salt) + String(password));
  return bytesToHex(bytes);
}

function base64UrlEncodeStr(str) {
  return Utilities.base64EncodeWebSafe(str, Utilities.Charset.UTF_8).replace(/=+$/, '');
}

function base64UrlDecodeStr(str) {
  return Utilities.newBlob(Utilities.base64DecodeWebSafe(str)).getDataAsString();
}

function signToken(payloadObj) {
  const payloadB64 = base64UrlEncodeStr(JSON.stringify(payloadObj));
  const sigBytes = Utilities.computeHmacSha256Signature(payloadB64, getAuthSecret());
  const sigB64 = Utilities.base64EncodeWebSafe(sigBytes).replace(/=+$/, '');
  return payloadB64 + '.' + sigB64;
}

// 有効なトークンならペイロード(email/name/isAdmin等)を返す。無効・期限切れならnull
function verifyToken(token) {
  if (!token || token.indexOf('.') === -1) return null;
  const parts = token.split('.');
  if (parts.length !== 2) return null;
  const payloadB64 = parts[0], sigB64 = parts[1];
  const expectedSigBytes = Utilities.computeHmacSha256Signature(payloadB64, getAuthSecret());
  const expectedSigB64 = Utilities.base64EncodeWebSafe(expectedSigBytes).replace(/=+$/, '');
  if (expectedSigB64 !== sigB64) return null;
  try {
    const payload = JSON.parse(base64UrlDecodeStr(payloadB64));
    if (payload.exp && Date.now() > payload.exp) return null;
    return payload;
  } catch (e) {
    return null;
  }
}


const TYPE_TO_SHEET_MAP = {
 "始業点検": "日常",
 "終業点検": "日常",
 "ME機器": "ME",
 "ナースバック": "ナースバック",
 "ヘリ内": "ヘリ内",
 "ヘリバック": "ヘリバック",
 "待機物品": "待機物品",
 "定期点検": "定期"
};


function getSheetFlexible(ss, names) {
 for (let n of names) {
   const s = ss.getSheetByName(n);
   if (s) return s;
 }
 return null;
}


function getDbSheet() {
 return getSheetFlexible(SpreadsheetApp.getActiveSpreadsheet(), ["DB_事案", "事案データベース"]);
}


// 🌟 修正: 「メールアドレス」列ではなく「固定送信先」列から全員宛のリストを取得する
function getMailToList() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
 if (!msSheet) return "";
 const data = msSheet.getDataRange().getDisplayValues();
 const head = data[0].map(h => String(h).trim());
 const cMail = head.indexOf("固定送信先") !== -1 ? head.indexOf("固定送信先") : 3;
 const emails = [];
 for (let i = 1; i < data.length; i++) {
   const email = String(data[i][cMail]).trim();
   if (email && email.includes("@")) emails.push(email);
 }
 return emails.join(", ");
}


// 🌟 修正: 存在しない「管理者メール」列ではなく「システム管理者」列から取得する
function getAdminMailList() {
 const ss = SpreadsheetApp.getActiveSpreadsheet();
 const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
 if (!msSheet) return "";
 const data = msSheet.getDataRange().getDisplayValues();
 const head = data[0].map(h => String(h).trim());
 const cAdmin = head.indexOf("システム管理者") !== -1 ? head.indexOf("システム管理者") : 6;
 const emails = [];
 for (let i = 1; i < data.length; i++) {
   const email = String(data[i][cAdmin]).trim();
   if (email && email.includes("@")) emails.push(email);
 }
 return emails.join(", ");
}


function resolveSheetName(type, isMaster) {
 const base = TYPE_TO_SHEET_MAP[type];
 if (!base) throw new Error(`システムエラー：未定義の点検タイプ「${type}」が要求されました。`);
 return isMaster ? `点検マスタ_${base}` : `DB_点検_${base}`;
}


function gasNormalizeDate(dateStr) {
 if (!dateStr) return "";
 var clean = dateStr.split(" ")[0];
 var parts = clean.split(/[\/\-]/);
 if (parts.length === 3) {
   var y = parts[0];
   var m = parts[1].length === 1 ? "0" + parts[1] : parts[1];
   var d = parts[2].length === 1 ? "0" + parts[2] : parts[2];
   return y + "-" + m + "-" + d;
 }
 return clean.replace(/\//g, "-");
}


function sanitizeInput(val) {
 if (typeof val !== "string") return val;
 let str = val.trim();
 if (str.startsWith("=") || str.startsWith("+") || str.startsWith("-") || str.startsWith("@")) {
   return "'" + str;
 }
 return str;
}

// 🌟 index.html の buildPreview() と同じ集計ロジックをメール本文用に移植したもの。
// 処置内容がメールに一切表示されず、代行入力者がメールだけでは完結できないという報告への対応（2026-08-17）。
function buildProcedureSummaryText_(appData) {
 const costItems = [];
 if (appData["往診料"]) costItems.push(appData["往診料"]);
 if (appData["搬送料"]) costItems.push(appData["搬送料"]);
 if (appData["長時間加算"]) costItems.push("長時間加算");
 if (appData["モニター算定時間"]) costItems.push(`モニター(${appData["モニター算定時間"]})`);
 if (appData["経皮的酸素飽和度"]) costItems.push("SpO2");
 if (appData["EtCO2算定"]) costItems.push("EtCO2");
 if (appData["管理料_人工呼吸器"]) costItems.push("人工呼吸器");
 if (appData["管理料_除細動装置"]) costItems.push("除細動装置");
 if (appData["管理料_閉鎖式保育器"]) costItems.push("閉鎖式保育器");
 if (appData["管理料_補助循環装置"]) costItems.push("PCPS等(補助循環)");
 if (appData["管理料_人工心肺装置"]) costItems.push("人工心肺");

 const procs = [];
 ["末梢静脈路_20G", "末梢静脈路_22G", "末梢静脈路_24G", "末梢静脈路_18G", "骨髄路_EZ-IO", "骨髄路_BIG"].forEach(k => {
   if (appData[k]) procs.push(`${k.split('_')[1]}(${String(appData[k]).replace('◯ ', '')})`);
 });
 if (appData["ベニューラ"]) procs.push(`ベニューラ${appData["ベニューラタイプ"] ? "(" + appData["ベニューラタイプ"] + ")" : ""}(${String(appData["ベニューラ"]).replace('◯ ', '')})`);
 if (appData["血液ガス分析"]) procs.push(`血ガス(${appData["血液ガスタイプ"]})`);
 if (appData["気管挿管サイズ"]) procs.push(`気管挿管(ID${appData["気管挿管サイズ"]}mm, ${appData["気管挿管位置"]}cm${appData["カフ上吸引"] ? "カフ吸引あり" : ""})`);
 if (appData["NPPVマスクサイズ"]) procs.push(`NPPV(${appData["NPPVマスクサイズ"]}サイズ)`);
 if (appData["除細動回数"]) procs.push(`除細動(${appData["除細動回数"]}回${appData["除細動パッド"] ? " パッド使用" : ""})`);
 if (appData["経皮ペーシング"]) procs.push(`経皮ペーシング${appData["経皮ペーシングパッド"] ? "(パッド使用)" : ""}`);
 if (appData["胸腔ドレナージ_20Fr"]) procs.push(`胸腔ドレナージ(20Fr:${appData["胸腔ドレナージ_20Fr"]}本)`);
 if (appData["胸腔ドレナージ_28Fr"]) procs.push(`胸腔ドレナージ(28Fr:${appData["胸腔ドレナージ_28Fr"]}本)`);
 if (appData["心嚢穿刺タイプ"]) procs.push(`心嚢穿刺(${appData["心嚢穿刺タイプ"]})`);
 if (appData["胃管挿入サイズ"]) procs.push(`胃管挿入(${appData["胃管挿入サイズ"]}Fr)`);

 const echos = [];
 ["FAST", "心エコー", "大血管エコー", "消化器系エコー", "婦人女性生殖器系エコー", "泌尿器系エコー", "下肢静脈エコー"].forEach(k => {
   if (appData[k]) echos.push(k.replace('エコー', ''));
 });
 if (appData["頸動脈エコー"]) echos.push(`頸動脈(${appData["頸動脈エコー"]})`);
 if (appData["その他エコー"]) echos.push(`その他(${appData["その他エコー"]})`);

 const others = [];
 ["12誘導心電図", "喀痰吸引", "血糖測定", "輪状甲状靭帯", "心膜開窓術", "開胸心マッサージ", "胸腔開放", "止血帯", "サムスリング", "シーネ固定", "バックボード", "頸椎カラー", "保温"].forEach(k => {
   if (appData[k]) others.push(k + (String(appData[k]).startsWith("◯ ") ? String(appData[k]).replace("◯ ", "(") + ")" : ""));
 });
 if (appData["その他外科処置"]) others.push(String(appData["その他外科処置"]).replace(/^◯\s*/, ''));

 let out = `■ 処置・医療機器\n ・ 確保ライン: ${procs.length > 0 ? procs.join(", ") : "なし"}\n ・ エコー: ${echos.length > 0 ? echos.join(", ") : "なし"}\n ・ その他処置: ${others.length > 0 ? others.join(", ") : "なし"}\n ・ 算定・コスト情報: ${costItems.length > 0 ? costItems.join(" / ") : "なし"}\n`;

 const drugList = String(appData["使用薬剤リスト"] || "").split(",").map(s => s.trim()).filter(Boolean);
 if (drugList.length > 0) out += `\n■ 使用薬剤\n ・ ${drugList.join(", ")}\n`;

 return out;
}

// 🌟 submit(新規登録)・send_email(再送)・update_record(上書き保存時の任意送信)で共通のメール本文を組み立てて送信する。
// 3箇所に同じテンプレート文字列を重複させないための共通化（2026-08-17）。
function sendCaseNotificationEmail_(appData, yoseiId, headerLine, subjectSuffix) {
 const bikoSection = appData["備考"] ? `\n■ 備考\n${appData["備考"]}\n` : "";
 const emailBody = `${headerLine}\n-----------------------------------------\n\n■ 基本情報\n ・ 要請番号 : No.${yoseiId}\n ・ 日付 : ${appData["日付"] || ""}\n ・ 出場先 : ${appData["出場先"] || ""}\n ・ 要請区分 : ${appData["要請区分"] || "未選択"}\n ・ Ｄｒ : ${appData["フライトドクター"] || ""} / Ｎｓ : ${appData["フライトナース"] || ""}\n ・ スキーム : ${appData["スキーム選択"] || "未選択"}\n ・ キーワード : ${appData["キーワード"] || "なし"}\n${bikoSection}\n■ 事案概要\n${appData["事案概要"] || "記述なし"}\n\n■ タイムライン\n ・ 要請/依頼: ${appData["要請時刻・施設間搬送依頼時刻"]||"--:--"}  離陸: ${appData["初期離陸時間"]||"--:--"}  着陸: ${appData["最終着陸時間"]||"--:--"}\n ・ 接触: ${appData["接触"]||"--:--"}  病着: ${appData["病着"]||"--:--"}  終了: ${appData["終了"]||"--:--"}\n\n${buildProcedureSummaryText_(appData)}\n■ カルテ完了報告 (本人 or 代行入力の完了報告はこちら)\n${GAS_API_URL}?id=${encodeURIComponent(yoseiId)}\n\n`;

 let mailTo = getMailToList(); let allEmails = [];
 if (mailTo) allEmails = mailTo.split(",").map(e => e.trim());
 if (appData["追加送信先"]) allEmails = allEmails.concat(String(appData["追加送信先"]).split(","));
 const uniqueEmails = [...new Set(allEmails)].filter(e => e).join(",");
 if (!uniqueEmails) return { sent: false, message: "送信先メールアドレスが設定されていません" };

 try {
   MailApp.sendEmail({ to: uniqueEmails, subject: `【要請 No.${yoseiId}】AW109 ${subjectSuffix}`, body: emailBody });
   return { sent: true };
 } catch (mailErr) {
   return { sent: false, message: "メール送信に失敗しました: " + mailErr.message };
 }
}


function findRowIndexBySysId(data, headers, targetId) {
 const idIdx = headers.indexOf("要請番号");
 const sysIdx = headers.indexOf("システムデータ");
  for (let i = 1; i < data.length; i++) {
   if (sysIdx !== -1 && data[i][sysIdx]) {
     try {
       const sysObj = JSON.parse(data[i][sysIdx]);
       if (sysObj.sysId === targetId) return i + 1;
     } catch(e) {}
   }
   if (idIdx !== -1 && String(data[i][idIdx]).trim() === String(targetId).trim()) {
     return i + 1;
   }
 }
 return -1;
}


function doGet(e) {
 const id = e.parameter.id;
 const resetToken = e.parameter.resetToken;
 if (resetToken) return renderResetPasswordPage_(resetToken);
 if (!id) return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Redirect OK" })).setMimeType(ContentService.MimeType.JSON);
  let html = `<!DOCTYPE html>
 <html lang="ja">
 <head>
   <meta charset="utf-8">
   <meta name="viewport" content="width=device-width,initial-scale=1">
   <title>完了報告</title>
   <style>
     body{font-family:sans-serif;background:#f2f2f7;color:#1c1c1e;padding:20px;}
     .card{background:#fff;border-radius:12px;padding:20px;max-width:400px;margin:0 auto;}
     button{width:100%;padding:14px;background:#007aff;color:#fff;border:none;border-radius:8px;font-weight:bold;font-size:16px;margin-top:20px;cursor:pointer;}
     input[type="text"]{width:100%;padding:10px;border:1px solid #ccc;border-radius:6px;margin-top:8px;}
     label{display:block;margin-top:12px;font-weight:bold;}
   </style>
 </head>
 <body>
   <div class="card">
     <h2 style="margin-top:0;">No.${id}<br>カルテ完了報告</h2>
     <label><input type="radio" name="who" value="本人" checked onchange="toggle()"> 本人入力</label>
     <label><input type="radio" name="who" value="代行" onchange="toggle()"> 代行入力</label>
     <div id="kana-box" style="display:none;">
       <input type="text" id="kana" placeholder="代行入力者名">
     </div>
     <button type="button" id="btn" onclick="send()">ステータス更新</button>
     <div id="msg" style="margin-top:16px; font-weight:bold; text-align:center; color:#065f46;"></div>
   </div>
   <script>
     function toggle() {
       document.getElementById("kana-box").style.display = document.querySelector('input[name="who"]:checked').value === "代行" ? "block" : "none";
     }
     function send() {
       const val = document.querySelector('input[name="who"]:checked').value;
       let who = val;
       if(val==="代行"){
         const kana=document.getElementById("kana").value.trim();
         if(!kana) return alert("名前を入力してください");
         who="代行: "+kana;
       }
       document.getElementById("btn").disabled = true;
       document.getElementById("btn").innerText = "送信中...";
       fetch("${GAS_API_URL}", {
         method:"POST",
         headers:{"Content-Type":"text/plain;charset=utf-8"},
         body:JSON.stringify({action:"update_status", password:"${PIN_ADMIN}", id:"${id}", who:who}),
         redirect:"follow"
       }).then(r=>r.text()).then(()=>{
         document.getElementById("btn").style.display="none";
         document.getElementById("msg").innerText="更新完了。閉じてください。";
       }).catch(e=>alert("通信エラー"));
     }
   </script>
 </body>
 </html>`;
 return HtmlService.createHtmlOutput(html);
}

// 🌟 パスワード再設定リンク（メール内URL）から開かれるフォーム。トークンの有効性判定はauth_reset_password側（doPost）で行うため、
// ここではトークンをそのままJSに埋め込んで渡すだけ。confirm()を使わないため、どのブラウザでも動作する。
function renderResetPasswordPage_(resetToken) {
  const html = `<!DOCTYPE html>
 <html lang="ja">
 <head>
   <meta charset="utf-8">
   <meta name="viewport" content="width=device-width,initial-scale=1">
   <title>パスワード再設定</title>
   <style>
     body{font-family:sans-serif;background:#f2f2f7;color:#1c1c1e;padding:20px;}
     .card{background:#fff;border-radius:12px;padding:20px;max-width:400px;margin:0 auto;}
     button{width:100%;padding:14px;background:#007aff;color:#fff;border:none;border-radius:8px;font-weight:bold;font-size:16px;margin-top:20px;cursor:pointer;}
     button:disabled{opacity:0.5;}
     input[type="password"]{width:100%;padding:10px;border:1px solid #ccc;border-radius:6px;margin-top:8px;box-sizing:border-box;}
     label{display:block;margin-top:12px;font-weight:bold;}
   </style>
 </head>
 <body>
   <div class="card">
     <h2 style="margin-top:0;">🔑 パスワード再設定</h2>
     <label>新しいパスワード (4文字以上)<input type="password" id="pw1"></label>
     <label>新しいパスワード (確認)<input type="password" id="pw2"></label>
     <button type="button" id="btn" onclick="send()">パスワードを変更する</button>
     <div id="msg" style="margin-top:16px; font-weight:bold; text-align:center;"></div>
   </div>
   <script>
     function send() {
       const pw1 = document.getElementById("pw1").value;
       const pw2 = document.getElementById("pw2").value;
       const msgEl = document.getElementById("msg");
       msgEl.style.color = "#d32f2f";
       if (pw1.length < 4) { msgEl.innerText = "4文字以上のパスワードを入力してください"; return; }
       if (pw1 !== pw2) { msgEl.innerText = "パスワードが一致しません"; return; }
       document.getElementById("btn").disabled = true;
       document.getElementById("btn").innerText = "送信中...";
       msgEl.innerText = "";
       fetch("${GAS_API_URL}", {
         method: "POST",
         headers: {"Content-Type": "text/plain;charset=utf-8"},
         body: JSON.stringify({ action: "auth_reset_password", token: ${JSON.stringify(resetToken)}, newPassword: pw1 }),
         redirect: "follow"
       }).then(r => r.json()).then(json => {
         if (json.status === "success") {
           document.getElementById("btn").style.display = "none";
           msgEl.style.color = "#2e7d32";
           msgEl.innerText = "パスワードを変更しました。このページを閉じて、新しいパスワードでログインしてください。";
         } else {
           document.getElementById("btn").disabled = false;
           document.getElementById("btn").innerText = "パスワードを変更する";
           msgEl.innerText = json.message || "変更に失敗しました。";
         }
       }).catch(e => {
         document.getElementById("btn").disabled = false;
         document.getElementById("btn").innerText = "パスワードを変更する";
         msgEl.innerText = "通信エラーが発生しました。";
       });
     }
   </script>
 </body>
 </html>`;
  return HtmlService.createHtmlOutput(html);
}

function doPost(e) {
 try {
   const requestData = JSON.parse(e.postData.contents);
   const pass = String(requestData.password);
   const action = requestData.action;
  
   const allowed = [
     "submit", "error_log", "fetch_init", "fetch_daily_report", "save_daily_report",
     "fetch_all_reports", "fetch_library", "fetch_all",
     "fetch_recent_cases", "fetch_debriefings", "submit_debriefing", "update_review_status", "update_status",
     "update_record", "delete_record", "send_email", "add_master", "submit_question",
     "answer_question", "fetch_checklist", "submit_checklist", "fetch_checklist_history",
     "fetch_checklist_status", "delete_checklist_record", "manage_news", "manage_manual", "manage_qa_full",
     "update_library_record", "auth_register", "auth_login", "set_admin_flag",
     "auth_request_reset", "auth_reset_password"
   ];

   if (!allowed.includes(action)) {
     return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "不明なアクション" })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "error_log") {
     const cache = CacheService.getScriptCache();
     const dedupKey = "errlog_" + Utilities.base64Encode(Utilities.newBlob(String(requestData.errorMsg||"") + String(requestData.source||"")).getBytes()).substring(0, 40);
     if (cache.get(dedupKey)) {
       return ContentService.createTextOutput(JSON.stringify({ status: "success", skipped: true })).setMimeType(ContentService.MimeType.JSON);
     }
     cache.put(dedupKey, "1", 180);


     const emails = getAdminMailList();
     if (emails) {
       const nowStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss");
       const whoStr = requestData.authName || "不明(PIN利用/未ログイン)";
       MailApp.sendEmail({
         to: emails,
         subject: `【AW109 EMS】システムエラー自動報告`,
         body: `現場の端末でシステムエラーが発生しました。\n\n・発生日時: ${nowStr}\n・画面: ${requestData.source || "不明"}\n・操作者: ${whoStr}\n・内容: ${requestData.errorMsg || "不明なエラー"}\n・ユーザー環境: ${requestData.userAgent || "不明"}\n\nシステムの確認をお願いします。`
       });
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "auth_register") {
     const email = String(requestData.email || "").trim().toLowerCase();
     const newPassword = String(requestData.newPassword || "");
     if (!email || newPassword.length < 4) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "メールアドレスと4文字以上のパスワードを入力してください" })).setMimeType(ContentService.MimeType.JSON);
     }
     const msSheet = getSheetFlexible(SpreadsheetApp.getActiveSpreadsheet(), ["マスタ_基本設定", "マスタデータ"]);
     if (!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     const data = msSheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => String(h).trim());
     const cMail = headers.indexOf("メールアドレス");
     const cHash = headers.indexOf("パスワードハッシュ");
     const cSalt = headers.indexOf("パスワードソルト");
     if (cMail === -1 || cHash === -1 || cSalt === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタ_基本設定に必要な列（メールアドレス/パスワードハッシュ/パスワードソルト）がありません" })).setMimeType(ContentService.MimeType.JSON);
     }
     let rowIndex = -1;
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][cMail]).trim().toLowerCase() === email) { rowIndex = i + 1; break; }
     }
     if (rowIndex === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "登録されていないメールアドレスです。管理者にマスタ登録を依頼してください" })).setMimeType(ContentService.MimeType.JSON);
     }
     if (data[rowIndex - 1][cHash]) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "このメールアドレスは既に登録済みです。ログインしてください" })).setMimeType(ContentService.MimeType.JSON);
     }
     const salt = Utilities.getUuid();
     const hash = hashPassword(newPassword, salt);
     msSheet.getRange(rowIndex, cSalt + 1).setValue(salt);
     msSheet.getRange(rowIndex, cHash + 1).setValue(hash);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "auth_login") {
     const email = String(requestData.email || "").trim().toLowerCase();
     const password = String(requestData.loginPassword || "");
     const msSheet = getSheetFlexible(SpreadsheetApp.getActiveSpreadsheet(), ["マスタ_基本設定", "マスタデータ"]);
     if (!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     const data = msSheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => String(h).trim());
     const cMail = headers.indexOf("メールアドレス");
     const cHash = headers.indexOf("パスワードハッシュ");
     const cSalt = headers.indexOf("パスワードソルト");
     const cName = headers.indexOf("宛先名");
     const cAdmin = headers.indexOf("システム管理者");
     if (cMail === -1 || cHash === -1 || cSalt === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタ_基本設定に必要な列がありません" })).setMimeType(ContentService.MimeType.JSON);
     }
     let found = null;
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][cMail]).trim().toLowerCase() === email) { found = data[i]; break; }
     }
     if (!found || !found[cHash]) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "メールアドレスまたはパスワードが違います" })).setMimeType(ContentService.MimeType.JSON);
     }
     if (hashPassword(password, found[cSalt]) !== found[cHash]) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "メールアドレスまたはパスワードが違います" })).setMimeType(ContentService.MimeType.JSON);
     }
     // 🌟「システム管理者」欄はメールアドレスと同じ行とは限らず、シート内のどこか1行にまとめて入力されていることがあるため、全行から探す
     let isAdmin = false;
     if (cAdmin !== -1) {
       for (let i = 1; i < data.length; i++) {
         if (String(data[i][cAdmin]).split(",").map(s => s.trim().toLowerCase()).includes(email)) { isAdmin = true; break; }
       }
     }
     const name = cName !== -1 && found[cName] ? found[cName] : email;
     const token = signToken({ email: email, name: name, isAdmin: isAdmin, iat: Date.now(), exp: Date.now() + 90 * 24 * 60 * 60 * 1000 });
     return ContentService.createTextOutput(JSON.stringify({ status: "success", token: token, name: name, isAdmin: isAdmin })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "auth_request_reset") {
     const email = String(requestData.email || "").trim().toLowerCase();
     if (!email) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "メールアドレスを入力してください" })).setMimeType(ContentService.MimeType.JSON);
     const msSheet = getSheetFlexible(SpreadsheetApp.getActiveSpreadsheet(), ["マスタ_基本設定", "マスタデータ"]);
     if (!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     const data = msSheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => String(h).trim());
     const cMail = headers.indexOf("メールアドレス");
     const cHash = headers.indexOf("パスワードハッシュ");
     const cName = headers.indexOf("宛先名");
     if (cMail === -1 || cHash === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタ_基本設定に必要な列がありません" })).setMimeType(ContentService.MimeType.JSON);
     }
     let found = null;
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][cMail]).trim().toLowerCase() === email) { found = data[i]; break; }
     }
     if (!found) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "登録されていないメールアドレスです。管理者にマスタ登録を依頼してください" })).setMimeType(ContentService.MimeType.JSON);
     }
     if (!found[cHash]) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "このメールアドレスはまだパスワード未設定です。「初めての方」から設定してください" })).setMimeType(ContentService.MimeType.JSON);
     }
     const name = cName !== -1 && found[cName] ? found[cName] : email;
     // 🌟 再設定リンクは30分のみ有効。signToken/verifyTokenの既存インフラを再利用し、専用の秘密情報を新設しない
     const resetToken = signToken({ email: email, purpose: "pwreset", exp: Date.now() + 30 * 60 * 1000 });
     const resetUrl = `${GAS_API_URL}?resetToken=${encodeURIComponent(resetToken)}`;
     MailApp.sendEmail({
       to: email,
       subject: "【AW109 EMS】パスワード再設定のご案内",
       body: `${name} 様\n\nパスワード再設定のリクエストを受け付けました。\n以下のリンクから30分以内に新しいパスワードを設定してください。\n\n${resetUrl}\n\n※このメールに心当たりがない場合は、無視してください。パスワードは変更されません。`
     });
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "auth_reset_password") {
     const token = String(requestData.token || "");
     const newPassword = String(requestData.newPassword || "");
     const payload = verifyToken(token);
     if (!payload || payload.purpose !== "pwreset") {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "リンクが無効か有効期限が切れています。もう一度パスワード再設定をリクエストしてください。" })).setMimeType(ContentService.MimeType.JSON);
     }
     if (newPassword.length < 4) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "4文字以上のパスワードを入力してください" })).setMimeType(ContentService.MimeType.JSON);
     }
     const email = payload.email;
     const msSheet = getSheetFlexible(SpreadsheetApp.getActiveSpreadsheet(), ["マスタ_基本設定", "マスタデータ"]);
     if (!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     const data = msSheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => String(h).trim());
     const cMail = headers.indexOf("メールアドレス");
     const cHash = headers.indexOf("パスワードハッシュ");
     const cSalt = headers.indexOf("パスワードソルト");
     let rowIndex = -1;
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][cMail]).trim().toLowerCase() === email) { rowIndex = i + 1; break; }
     }
     if (rowIndex === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象のアカウントが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     }
     const salt = Utilities.getUuid();
     const hash = hashPassword(newPassword, salt);
     msSheet.getRange(rowIndex, cSalt + 1).setValue(salt);
     msSheet.getRange(rowIndex, cHash + 1).setValue(hash);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }

   // 🌟 認証: スタッフ用PINは廃止。管理者用PIN(非公開の緊急回復用)か、ログイントークンのみ受け付ける
   let isAdminUser = false;
   let authName = "";
   if (pass === PIN_ADMIN) {
     isAdminUser = true;
   } else {
     const tokenPayload = verifyToken(pass);
     if (!tokenPayload) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "認証エラー" })).setMimeType(ContentService.MimeType.JSON);
     }
     isAdminUser = !!tokenPayload.isAdmin;
     authName = tokenPayload.name || tokenPayload.email || "";
   }
   if (authName) requestData.authName = authName;

   const ss = SpreadsheetApp.getActiveSpreadsheet();


   if (action === "fetch_daily_report") {
     const targetDate = gasNormalizeDate(requestData.date);
     const forceSync = requestData.force_sync;
    
     let reportData = { recorder: "", doctor: "", hamamachiReason: "", events: [] };
    
     const hSheet = getSheetFlexible(ss, ["DB_日報", "浜町日誌"]);
     if (hSheet) {
       const d = hSheet.getDataRange().getDisplayValues();
       for (let i = 1; i < d.length; i++) {
         if (gasNormalizeDate(d[i][0]) === targetDate) {
           reportData.recorder = d[i][1];
           reportData.doctor = d[i][2];
           reportData.hamamachiReason = d[i][3];
           try { reportData.events = JSON.parse(d[i][4]); } catch(e) { reportData.events = []; }
           break;
         }
       }
     }
    
     if (forceSync || reportData.events.length === 0) {
       const dbSheet = getDbSheet();
       if (dbSheet) {
         const d = dbSheet.getDataRange().getDisplayValues();
         const head = d[0].map(h => h.trim());
         const idIdx = head.indexOf("要請番号");
         const dateIdx = head.indexOf("日付");
         const destIdx = head.indexOf("出場先");
         const drIdx = head.indexOf("フライトドクター");
         const nsIdx = head.indexOf("フライトナース");
         const reqTimeIdx = head.indexOf("要請時刻・施設間搬送依頼時刻");
         const takeoffIdx = head.indexOf("初期離陸時間");
         const landIdx = head.indexOf("最終着陸時間");
         const schemeIdx = head.indexOf("スキーム選択");
         const reqTypeIdx = head.indexOf("要請区分");
         const diagIdx = head.indexOf("現場診断");
        
         let autoEvents = [];
         for (let i = 1; i < d.length; i++) {
           if (gasNormalizeDate(d[i][dateIdx]) === targetDate) {
             const caseId = d[i][idIdx] || "未定義";
             const scheme = d[i][schemeIdx] || "";
             const isCancel = scheme.includes("キャンセル");
             const startTime = d[i][takeoffIdx] || d[i][reqTimeIdx] || "08:30";
             const endTime = d[i][landIdx] || "09:30";
             const reqTypeStr = reqTypeIdx !== -1 ? d[i][reqTypeIdx] : "";
             const diagStr = diagIdx !== -1 ? d[i][diagIdx] : "";
            
             const descStr = `【${isCancel ? 'キャンセル' : '実施'}】[${reqTypeStr}] ${d[i][destIdx] || ''} 診断:${diagStr} (Dr:${d[i][drIdx] || '-'}/Ns:${d[i][nsIdx] || '-'})`;
            
             autoEvents.push({
               id: "auto-" + caseId,
               cat: isCancel ? "cancel" : "mission",
               no: caseId,
               start: startTime,
               end: endTime,
               desc: descStr,
               cont: false
             });
            
             if (!reportData.doctor && d[i][drIdx]) reportData.doctor = d[i][drIdx];
             if (!reportData.recorder && d[i][nsIdx]) reportData.recorder = d[i][nsIdx];
           }
         }
        
         if (forceSync) {
           const manualEvents = reportData.events.filter(e => !String(e.id).startsWith("auto-"));
           reportData.events = [...manualEvents, ...autoEvents];
         } else if (reportData.events.length === 0) {
           reportData.events = autoEvents;
         }
       }
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: reportData })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "save_daily_report") {
     const hSheet = getSheetFlexible(ss, ["DB_日報", "浜町日誌"]);
     if (!hSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "日報データベースのシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     const targetDate = gasNormalizeDate(requestData.date);
     const d = hSheet.getDataRange().getDisplayValues();
     let rowIndex = -1;
     for (let i = 1; i < d.length; i++) {
       if (gasNormalizeDate(d[i][0]) === targetDate) {
         rowIndex = i + 1;
         break;
       }
     }
    
     const writeRow = [
       requestData.date,
       requestData.recorder,
       requestData.doctor,
       requestData.hamamachiReason,
       JSON.stringify(requestData.events)
     ].map(sanitizeInput);
    
     if (rowIndex !== -1) {
       hSheet.getRange(rowIndex, 1, 1, writeRow.length).setValues([writeRow]);
     } else {
       hSheet.appendRow(writeRow);
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }
// 🌟 統括・業務管理コンソール用：日報データベースの全件取得
   if (action === "fetch_all_reports") {
     const hSheet = getSheetFlexible(ss, ["DB_日報", "浜町日誌"]);
     if (!hSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "日報データベースのシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     const d = hSheet.getDataRange().getDisplayValues();
     const result = [];
     for (let i = 1; i < d.length; i++) {
       if (!d[i][0]) continue;
       let events = [];
       try { events = JSON.parse(d[i][4]); } catch(e) { events = []; }
       result.push({
         date: d[i][0],
         recorder: d[i][1],
         doctor: d[i][2],
         reason: d[i][3],
         events: events
       });
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result })).setMimeType(ContentService.MimeType.JSON);
   }
   if (action === "fetch_init") {
     let mastersData = { dest: [], dr: [], ns: [], mailOptions: [], phrases: [], diagnosis: [], keyword: [], extAffil: [], extStaff: [], pilot: [], mechanic: [], cs: [], drugs: [] };

     const drugSheet = getSheetFlexible(ss, ["マスタ_処置項目"]);
     if (drugSheet) {
       const dd = drugSheet.getDataRange().getDisplayValues();
       for (let i = 1; i < dd.length; i++) {
         const name = String(dd[i][0]).trim();
         if (name) mastersData.drugs.push(name);
       }
     }

     const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
     if (msSheet) {
       const d = msSheet.getDataRange().getDisplayValues();
       const hd = d[0].map(h => String(h).trim());
      
       const col = (name, fallback) => hd.indexOf(name) !== -1 ? hd.indexOf(name) : fallback;
      
       // 🌟 ご提示いただいたマスタの項目並びに完全一致するフォールバック
       const cDest = col("出場先", 0);
       const cDr = col("フライトドクター", 1);
       const cNs = col("フライトナース", 2);
       const cMailName = col("宛先名", 4);
       const cMail = col("メールアドレス", 5);
       const cPhrase = col("キャンセル理由", 8);
       const cExtAffil = col("外部所属", 9);
       const cExtStaff = col("外部スタッフ", 10);
       const cPilot = col("機長", 11);
       const cMech = col("整備士", 12);
       const cCs = col("CS", 13);
       const cDiag = col("現場診断", 14);
       const cKey = col("キーワード", 15);


       for(let i=1; i<d.length; i++) {
         if (d[i][cDest]) mastersData.dest.push(String(d[i][cDest]).trim());
         if (d[i][cDr]) mastersData.dr.push(String(d[i][cDr]).trim());
         if (d[i][cNs]) mastersData.ns.push(String(d[i][cNs]).trim());
         if (d[i][cMailName] && d[i][cMail]) mastersData.mailOptions.push({ name: String(d[i][cMailName]).trim(), email: String(d[i][cMail]).trim() });
         if (d[i][cPhrase]) mastersData.phrases.push(String(d[i][cPhrase]).trim());
         if (d[i][cExtAffil]) mastersData.extAffil.push(String(d[i][cExtAffil]).trim());
         if (d[i][cExtStaff]) mastersData.extStaff.push(String(d[i][cExtStaff]).trim());
         if (d[i][cPilot]) mastersData.pilot.push(String(d[i][cPilot]).trim());
         if (d[i][cMech]) mastersData.mechanic.push(String(d[i][cMech]).trim());
         if (d[i][cCs]) mastersData.cs.push(String(d[i][cCs]).trim());
         if (d[i][cDiag]) mastersData.diagnosis.push(String(d[i][cDiag]).trim());
         if (d[i][cKey]) mastersData.keyword.push(String(d[i][cKey]).trim());
       }
      
       mastersData.extAffil = [...new Set(mastersData.extAffil)];
       mastersData.extStaff = [...new Set(mastersData.extStaff)];
       mastersData.pilot = [...new Set(mastersData.pilot)];
       mastersData.mechanic = [...new Set(mastersData.mechanic)];
       mastersData.cs = [...new Set(mastersData.cs)];
       mastersData.diagnosis = [...new Set(mastersData.diagnosis)];
       mastersData.keyword = [...new Set(mastersData.keyword)];
     }
    
     let qaData = [], manuals = [], alerts = [], newsData = [];
     const appName = requestData.app_name || "不明";


     const newsSheet = getSheetFlexible(ss, ["DB_お知らせ", "お知らせデータベース"]);
     if (newsSheet) {
       const nd = newsSheet.getDataRange().getDisplayValues();
       for (let i = 1; i < nd.length; i++) {
         if (nd[i][0]) {
           if (appName === "管理" || String(nd[i][4]).trim() !== "非公開") {
             newsData.push({
               rowIdx: i + 1,
               date: String(nd[i][0]).trim(), category: String(nd[i][1]).trim(),
               title: String(nd[i][2]).trim(), content: String(nd[i][3]).trim(),
               status: String(nd[i][4]).trim()
             });
           }
         }
       }
       if (appName !== "管理") newsData.reverse();
     }


     const qaSheet = getSheetFlexible(ss, ["DB_QA", "Q&Aデータベース"]);
     if (qaSheet) {
       const d = qaSheet.getDataRange().getDisplayValues();
       for(let i=1; i<d.length; i++) {
         if (d[i][3] === "◯" || appName === "管理") {
           qaData.push({ rowIdx: i + 1, date: String(d[i][0]).trim(), q: String(d[i][1]).trim(), a: String(d[i][2]).trim(), status: String(d[i][3]).trim(), source: String(d[i][4]||"不明").trim(), target: String(d[i][5]||"").trim() });
         }
       }
     }
    
     const mSheet = getSheetFlexible(ss, ["システム_取扱説明書", "マスタ_取扱説明書", "取扱説明書_現場", "取扱説明書_検索"]);
     if (mSheet) {
       const d = mSheet.getDataRange().getDisplayValues();
       for(let i=1; i<d.length; i++) {
         if (String(d[i][1]).trim() !== "") {
           const target = String(d[i][4] || "全て").trim();
           if (target === "" || target.includes("全て") || target.includes(appName) || appName === "管理") {
             manuals.push({ rowIdx: i + 1, title: String(d[i][1]).trim(), text: String(d[i][2]).trim(), target: target });
           }
         }
       }
     }
    
     if (appName === "ポータル") {
       const dbSheet = getDbSheet();
       if (dbSheet) {
         const d = dbSheet.getDataRange().getDisplayValues();
         const statusIdx = d[0].indexOf("ステータス");
         let uncompletedCount = 0;
         for(let i=1; i<d.length; i++) { if(d[i][statusIdx] && d[i][statusIdx].includes("未完了")) uncompletedCount++; }
         if (uncompletedCount > 0) alerts.push({ type: "warning", text: `⚠️ カルテ未完了の事案が ${uncompletedCount} 件あります。[ココをクリックして処理]`, count: uncompletedCount });
       }
      
       const now = new Date(); const date = now.getDate(); const dayOfWeek = now.getDay();
       if (dayOfWeek === 2 && date <= 7) alerts.push({ type: "warning", text: `🗓️ 【本日は第1火曜です】モナール点検・シリンジポンプ交換を実施してください。` });
       if (dayOfWeek === 3 && date >= 8 && date <= 14) alerts.push({ type: "warning", text: `🗓️ 【本日は第2水曜です】除細動器(フィリップス小型)の点検を実施してください。` });


       const limit30 = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000);
       const allSheets = ss.getSheets();
       allSheets.forEach(sheet => {
         const sName = sheet.getName();
         if(sName.startsWith("点検マスタ_") || sName.includes("点検マスタ")) {
           const d = sheet.getDataRange().getDisplayValues();
           if(d.length > 0) {
             const expIdx = d[0].indexOf("有効期限");
             if(expIdx !== -1) {
               for(let i=1; i<d.length; i++) {
                 let expStr = String(d[i][expIdx]).trim();
                 if(expStr) {
                   let expDate = new Date(expStr.replace(/[\/\.-]/g, '/'));
                   if(!isNaN(expDate.getTime())) {
                     if (expDate < now) alerts.push({ type: "error", text: `🚨 【期限切れ】${sName.replace('点検マスタ_','')}の「${d[i][1]}」の有効期限が切れています！(${expStr})` });
                     else if (expDate <= limit30) alerts.push({ type: "warning", text: `⚠️ 【期限間近】${sName.replace('点検マスタ_','')}の「${d[i][1]}」の期限が近づいています(${expStr}まで)` });
                   }
                 }
               }
             }
           }
         }
       });
     }
    
     return ContentService.createTextOutput(JSON.stringify({ status: "success", masters: mastersData, qa: qaData, manuals: manuals, alerts: alerts, news: newsData, gasVersion: GAS_VERSION })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "manage_news") {
     if (!isAdminUser) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "管理者権限が必要です" })).setMimeType(ContentService.MimeType.JSON);
     const sheet = getSheetFlexible(ss, ["DB_お知らせ", "お知らせデータベース"]);
     if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     const { sub_action, rowIdx, data } = requestData;
     if (sub_action === "add") {
       sheet.appendRow([data.date, data.category, data.title, data.content, data.status].map(sanitizeInput));
     } else if (sub_action === "update") {
       sheet.getRange(rowIdx, 1, 1, 5).setValues([[data.date, data.category, data.title, data.content, data.status].map(sanitizeInput)]);
     } else if (sub_action === "delete") {
       sheet.deleteRow(rowIdx);
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "manage_manual") {
     if (!isAdminUser) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "管理者権限が必要です" })).setMimeType(ContentService.MimeType.JSON);
     const sheet = getSheetFlexible(ss, ["システム_取扱説明書", "マスタ_取扱説明書", "取扱説明書_現場", "取扱説明書_検索"]);
     if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     const { sub_action, rowIdx, data } = requestData;
     if (sub_action === "add") {
       sheet.appendRow(["", data.title, data.text, "", data.target].map(sanitizeInput));
     } else if (sub_action === "update") {
       const currentVals = sheet.getRange(rowIdx, 1, 1, sheet.getLastColumn()).getValues()[0];
       currentVals[1] = sanitizeInput(data.title);
       currentVals[2] = sanitizeInput(data.text);
       currentVals[4] = sanitizeInput(data.target);
       sheet.getRange(rowIdx, 1, 1, currentVals.length).setValues([currentVals]);
     } else if (sub_action === "delete") {
       sheet.deleteRow(rowIdx);
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "manage_qa_full") {
     if (!isAdminUser) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "管理者権限が必要です" })).setMimeType(ContentService.MimeType.JSON);
     const sheet = getSheetFlexible(ss, ["DB_QA", "Q&Aデータベース"]);
     if (!sheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     const { sub_action, rowIdx, data } = requestData;
     if (sub_action === "update") {
       sheet.getRange(rowIdx, 1, 1, 6).setValues([[data.date, data.q, data.a, data.status, data.source, data.target].map(sanitizeInput)]);
     } else if (sub_action === "delete") {
       sheet.deleteRow(rowIdx);
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "add_master") {
     const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
     if(!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが存在しません" })).setMimeType(ContentService.MimeType.JSON);
    
     const type = requestData.type; const val1 = requestData.val1; const val2 = requestData.val2 || "";
    
     const head = msSheet.getRange(1, 1, 1, msSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
     const col = (name, fallback) => head.indexOf(name) !== -1 ? head.indexOf(name) + 1 : fallback;


     // 🌟 ご提示いただいたマスタの項目並びに完全一致するフォールバック
     let colIdx = 1; let writeVals = [val1];
     if (type === 'dest') colIdx = col("出場先", 1);
     else if (type === 'dr') colIdx = col("フライトドクター", 2);
     else if (type === 'ns') colIdx = col("フライトナース", 3);
     else if (type === 'mail') { colIdx = col("宛先名", 5); writeVals = [val1, val2]; }
     else if (type === 'ext_affil') colIdx = col("外部所属", 10);
     else if (type === 'ext_staff') colIdx = col("外部スタッフ", 11);
     else if (type === 'pilot') colIdx = col("機長", 12);
     else if (type === 'mechanic') colIdx = col("整備士", 13);
     else if (type === 'cs') colIdx = col("CS", 14);
     else if (type === 'diagnosis') colIdx = col("現場診断", 15);
     else if (type === 'keyword') colIdx = col("キーワード", 16);
    
     const colData = msSheet.getRange(1, colIdx, msSheet.getLastRow() || 1).getDisplayValues();
     let insertRow = 2;
     for(let i=colData.length-1; i>=0; i--) {
       if(colData[i][0] !== "") { insertRow = i+2; break; }
     }
     msSheet.getRange(insertRow, colIdx, 1, writeVals.length).setValues([writeVals.map(sanitizeInput)]);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "set_admin_flag") {
     // 🌟 管理者権限の付与・剥奪。誤操作の影響が大きいため管理者本人のみ実行可能にする
     if (!isAdminUser) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "管理者権限が必要です" })).setMimeType(ContentService.MimeType.JSON);
     const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
     if (!msSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);

     const email = String(requestData.email || "").trim().toLowerCase();
     const makeAdmin = !!requestData.isAdmin;
     if (!email) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "メールアドレスを指定してください" })).setMimeType(ContentService.MimeType.JSON);

     const d = msSheet.getDataRange().getDisplayValues();
     const head = d[0].map(h => String(h).trim());
     const cMail = head.indexOf("メールアドレス");
     const cAdmin = head.indexOf("システム管理者");
     if (cMail === -1 || cAdmin === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタ_基本設定に必要な列（メールアドレス/システム管理者）がありません" })).setMimeType(ContentService.MimeType.JSON);

     let rowIndex = -1;
     for (let i = 1; i < d.length; i++) {
       if (String(d[i][cMail]).trim().toLowerCase() === email) { rowIndex = i + 1; break; }
     }
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "登録されていないメールアドレスです。先に「メール送信先」マスタへの追加が必要です" })).setMimeType(ContentService.MimeType.JSON);

     // 🌟 このメールアドレス自身の行の「システム管理者」列に自分のメールを書く/消すだけにする。
     // auth_login側は全行のシステム管理者列をカンマ区切りで走査するため、この行の値だけ変えれば十分。
     msSheet.getRange(rowIndex, cAdmin + 1).setValue(makeAdmin ? email : "");
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "fetch_checklist_status") {
     const types = Object.keys(TYPE_TO_SHEET_MAP);
     let targetDateStr = requestData.date;
     if (!targetDateStr) { targetDateStr = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd"); }
     const targetDash = targetDateStr.replace(/\//g, "-");
     const targetSlash = targetDateStr.replace(/-/g, "/");
    
     let statusMap = {};
     types.forEach(t => {
       statusMap[t] = false;
       let sheetName = resolveSheetName(t, false);
       let dbSheet = ss.getSheetByName(sheetName);
       if(dbSheet) {
         const d = dbSheet.getDataRange().getDisplayValues();
         const head = d[0].map(h => String(h).trim());
         const dateIdx = head.indexOf("点検日");
         const jsonIdx = head.indexOf("詳細データJSON");


         if(dateIdx !== -1 && jsonIdx !== -1) {
           for(let i = d.length - 1; i > 0; i--) {
             if(d[i][dateIdx] === targetDash || d[i][dateIdx] === targetSlash) {
               if (t === "始業点検" || t === "終業点検") {
                 if (String(d[i][jsonIdx]).includes(t)) { statusMap[t] = true; break; }
               } else {
                 statusMap[t] = true; break;
               }
             }
           }
         }
       }
     });
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: statusMap })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "delete_checklist_record") {
     const typeName = requestData.type || "ME機器";
     let sheetName = resolveSheetName(typeName, false);
     let dbSheet = ss.getSheetByName(sheetName);
     if (!dbSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `データ保存シート(${sheetName})が見つかりません。` })).setMimeType(ContentService.MimeType.JSON);
    
     const targetId = requestData.id;
     const data = dbSheet.getDataRange().getDisplayValues();
     const head = data[0].map(h => String(h).trim());
     const idIdx = head.indexOf("点検ID");
    
     if(idIdx === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートに「点検ID」列が存在しません。" })).setMimeType(ContentService.MimeType.JSON);


     let rowIndex = -1;
     for (let i = 1; i < data.length; i++) {
       if (String(data[i][idIdx]).trim() === String(targetId).trim()) { rowIndex = i + 1; break; }
     }
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象レコードが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     dbSheet.deleteRow(rowIndex);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "fetch_checklist") {
     const typeName = requestData.type || "ME機器";
     let sheetName = resolveSheetName(typeName, true);
     let clSheet = ss.getSheetByName(sheetName);
     if (!clSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `マスタシート(${sheetName})が見つかりません。` })).setMimeType(ContentService.MimeType.JSON);
    
     const d = clSheet.getDataRange().getDisplayValues();
     let items = [];
     if(d.length > 1) {
       const head = d[0].map(h => String(h).trim());
       const catIdx = head.indexOf("大分類");
       const subIdx = head.indexOf("中分類");
       const nameIdx = head.indexOf("点検項目");
       const constIdx = head.indexOf("定数");
       const typeIdx = head.indexOf("入力タイプ");
       const reqIdx = head.indexOf("必須");
       const orderIdx = head.indexOf("表示順");


       if(catIdx === -1 || nameIdx === -1) {
         return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "マスタに必須列（大分類、点検項目）が存在しません。" })).setMimeType(ContentService.MimeType.JSON);
       }


       for(let i=1; i<d.length; i++) {
         if(d[i][nameIdx]) {
           items.push({
             category: d[i][catIdx] || "",
             subCategory: subIdx !== -1 ? d[i][subIdx] : "",
             name: d[i][nameIdx] || "",
             constant: constIdx !== -1 ? d[i][constIdx] : "",
             inputType: typeIdx !== -1 ? d[i][typeIdx] || "◯/✕" : "◯/✕",
             required: reqIdx !== -1 ? String(d[i][reqIdx]).toUpperCase() === "TRUE" : false,
             order: orderIdx !== -1 ? (Number(d[i][orderIdx]) || 99) : 99
           });
         }
       }
     }
     items.sort((a, b) => a.order - b.order);
    
     let nsList = [];
     const msSheet = getSheetFlexible(ss, ["マスタ_基本設定", "マスタデータ"]);
     if (msSheet) {
       const md = msSheet.getDataRange().getDisplayValues();
       const mHead = md[0].map(h => String(h).trim());
       const cNs = mHead.indexOf("フライトナース") !== -1 ? mHead.indexOf("フライトナース") : 2;
       for(let i=1; i<md.length; i++) { if(md[i][cNs]) nsList.push(md[i][cNs]); }
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", items: items, nsList: nsList })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "submit_checklist") {
     const typeName = requestData.type || "ME機器";
     let sheetName = resolveSheetName(typeName, false);
     let dbSheet = ss.getSheetByName(sheetName);
     if (!dbSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `データ保存シート(${sheetName})が見つかりません。` })).setMimeType(ContentService.MimeType.JSON);
    
     const p = requestData.data;
     const newId = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMdd-HHmm");
    
     const head = dbSheet.getRange(1, 1, 1, dbSheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
     let newRow = new Array(head.length).fill("");
    
     const setVal = (colName, val) => { const idx = head.indexOf(colName); if(idx !== -1) newRow[idx] = sanitizeInput(val); };
    
     setVal("点検ID", newId);
     setVal("点検日", p.date);
     setVal("実施時間", p.time);
     setVal("点検者", p.staff);
     setVal("総合判定", p.judgment);
     setVal("異常項目リスト", p.errorItems);
     setVal("詳細データJSON", JSON.stringify(p.rawJson));


     dbSheet.appendRow(newRow);
    
     if(p.judgment === "異常あり" || p.errorItems !== "なし") {
       const emails = getAdminMailList();
       if (emails) {
         MailApp.sendEmail({
           to: emails,
           subject: `【AW109 EMS】⚠️ ${typeName}点検で異常が報告されました`,
           body: `【点検報告アラート】\n\n・日付: ${p.date} ${p.time}\n・点検者: ${p.staff}\n・異常項目: ${p.errorItems}\n\nシステムで詳細を確認し、対応をお願いします。`
         });
       }
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "fetch_checklist_history") {
     const typeName = requestData.type || "ME機器";
     let sheetName = resolveSheetName(typeName, false);
     let dbSheet = ss.getSheetByName(sheetName);
     if (!dbSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: `データシートが見つかりません` })).setMimeType(ContentService.MimeType.JSON);
    
     const d = dbSheet.getDataRange().getDisplayValues();
     const head = d[0].map(h => String(h).trim());
    
     const cId = head.indexOf("点検ID");
     const cDate = head.indexOf("点検日");
     const cTime = head.indexOf("実施時間");
     const cStaff = head.indexOf("点検者");
     const cJudg = head.indexOf("総合判定");
     const cErr = head.indexOf("異常項目リスト");
     const cJson = head.indexOf("詳細データJSON");


     if(cId === -1 || cDate === -1 || cJson === -1) {
       return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "データシートに必須列が存在しません" })).setMimeType(ContentService.MimeType.JSON);
     }


     let history = [];
     for(let i = d.length - 1; i > 0; i--) {
       if(d[i][cId]) {
         let rawData = {};
         try { rawData = JSON.parse(d[i][cJson]); } catch(e){}
         history.push({
           id: d[i][cId],
           date: d[i][cDate],
           time: cTime !== -1 ? d[i][cTime] : "",
           staff: cStaff !== -1 ? d[i][cStaff] : "",
           judgment: cJudg !== -1 ? d[i][cJudg] : "",
           errorItems: cErr !== -1 ? d[i][cErr] : "",
           rawJson: rawData
         });
       }
       if(history.length >= 60) break;
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", history: history })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "submit_question") {
     const qSheet = getSheetFlexible(ss, ["DB_QA", "Q&Aデータベース"]);
     if(qSheet) qSheet.appendRow([Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm"), requestData.question, "", "未", requestData.source, ""].map(sanitizeInput));
     const emails = getAdminMailList();
     if (emails) {
       MailApp.sendEmail({
         to: emails,
         subject: `【AW109 EMS】新しいQ&A・要望が届きました`,
         body: `現場から新しい要望・質問が送信されました。\n\n・送信元画面: ${requestData.source || "不明"}\n・内容:\n${requestData.question}\n\nシステム管理コンソール(admin.html)から回答を行ってください。`
       });
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }
  
   // 🌟 物理ヘッダーに完全対応した事案データフェッチ
   if (action === "fetch_recent_cases" || action === "fetch_all") {
     const sheet = getDbSheet();
     const data = sheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => String(h).trim());
     const result = [];
     for(let i = 1; i < data.length; i++) {
       let obj = {};
       headers.forEach((h, idx) => { obj[h] = String(data[i][idx]); });
      
       let sysId = obj["要請番号"];
       if(obj["システムデータ"]) {
         try {
           const sData = JSON.parse(obj["システムデータ"]);
           if(sData.sysId) sysId = sData.sysId;
         } catch(e) {}
       }
      
       result.push({
         id: sysId,
         displayId: obj["要請番号"],
         date: obj["日付"],
         dest: obj["出場先"],
         scheme: obj["スキーム選択"],
         rawData: obj
       });
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: result })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "fetch_debriefings") {
     const dSheet = getSheetFlexible(ss, ["DB_デブリーフィング", "デブリーフィングデータベース"]);
     if (!dSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが存在しません" })).setMimeType(ContentService.MimeType.JSON);
     const d = dSheet.getDataRange().getDisplayValues();
     const headers = d[0].map(h => String(h).trim());
     const res = [];
     for(let i=1; i<d.length; i++) {
       let obj = {};
       headers.forEach((h, idx) => { obj[h] = String(d[i][idx]); });
       res.push(obj);
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: res })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "delete_record") {
     const sheet = getDbSheet();
     const data = sheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => h.trim());
     const rowIndex = findRowIndexBySysId(data, headers, requestData.id);
    
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象レコードが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
     sheet.deleteRow(rowIndex);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "update_record") {
     const sheet = getDbSheet();
     const appData = requestData.data;
     const oldId = requestData.old_id;
     const editor = requestData.editor;
    
     const data = sheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => h.trim());
     const rowIndex = findRowIndexBySysId(data, headers, oldId);
    
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象レコードが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
    
     let newRow = [...data[rowIndex - 1]];
     headers.forEach((header, idx) => {
       if (header === "更新履歴") {
         newRow[idx] = `${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm")} : 修正保存 (${editor})\n` + String(data[rowIndex-1][idx] || "");
       } else if (header === "システムデータ") {
         let sysDataObj = {};
         try { sysDataObj = JSON.parse(newRow[idx] || "{}"); } catch(e) {}
         if(appData["システムデータ"]) {
           try {
             const reqSys = JSON.parse(appData["システムデータ"]);
             sysDataObj = { ...sysDataObj, ...reqSys };
           } catch(e) {}
         }
         if(!sysDataObj.sysId) sysDataObj.sysId = oldId;
         newRow[idx] = JSON.stringify(sysDataObj);
       } else if (header === "日付" && appData["日付"] !== undefined) {
         newRow[idx] = "'" + String(appData["日付"]).replace(/-/g, '/');
       } else if (appData[header] !== undefined) {
         newRow[idx] = sanitizeInput(appData[header]);
       }
     });
     sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);

     // 🌟 上書き保存後の確認メール送信を同一リクエストにまとめる（2026-08-17）。
     // 従来はクライアント側でupdate_record→confirm→send_emailと2回GAS呼び出しを直列に行っており、
     // 通信タイムアウトが頻発する一因になっていた。newRowは既にメモリ上にあるため再読込不要で送信できる。
     let emailResult = null;
     if (requestData.sendEmail) {
       let updatedAppData = {};
       headers.forEach((h, idx) => { updatedAppData[h] = newRow[idx]; });
       const yoseiId = updatedAppData["要請番号"];
       emailResult = sendCaseNotificationEmail_(updatedAppData, yoseiId, "【AW109 EMS 事案記録 (修正後再送)】", "事案記録 (修正後再送)");
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", emailSent: emailResult ? emailResult.sent : undefined, emailMessage: emailResult ? emailResult.message : undefined })).setMimeType(ContentService.MimeType.JSON);
   }

   if (action === "submit") {
     const sheet = getDbSheet();
     const appData = requestData.data;
     const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
     const newRow = [];
    
     const generatedSysId = Utilities.getUuid();
    
     headers.forEach(header => {
       if (header === "ステータス") {
         newRow.push(appData["ステータス"] || "⚠️ 未完了(現場入力)");
       } else if (header === "更新履歴") {
         newRow.push(`${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm")} : 新規登録\n`);
       } else if (header === "日付" && appData["日付"] !== undefined) {
         newRow.push("'" + String(appData["日付"]).replace(/-/g, '/'));
       } else if (header === "システムデータ") {
         let sysDataObj = {};
         if (appData["システムデータ"]) {
           try { sysDataObj = JSON.parse(appData["システムデータ"]); } catch(e) {}
         }
         sysDataObj.sysId = generatedSysId;
         newRow.push(JSON.stringify(sysDataObj));
       } else if (appData[header] !== undefined) {
         newRow.push(sanitizeInput(appData[header]));
       } else {
         newRow.push("");
       }
     });
     sheet.appendRow(newRow);
    
     const yoseiId = appData["要請番号"];
     sendCaseNotificationEmail_(appData, yoseiId, "【AW109 EMS 新規事案登録】\n現場アプリから事案が登録されました。", "新規事案記録");
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "update_status") {
     const sheet = getDbSheet();
     const data = sheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => h.trim());
    
     const targetId = String(requestData.id).trim();
     const idIdx = headers.indexOf("要請番号");
     const statusIdx = headers.indexOf("ステータス");
     const histIdx = headers.indexOf("更新履歴");
     let found = false;
    
     for (let i = 1; i < data.length; i++) {
       if (idIdx !== -1 && String(data[i][idIdx]).trim() === targetId) {
         sheet.getRange(i + 1, statusIdx + 1).setValue(`◯ カルテ完了 (${requestData.who})`);
         if (histIdx !== -1) sheet.getRange(i + 1, histIdx + 1).setValue(`${Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm")} : カルテ完了処理 (${requestData.who})\n` + String(data[i][histIdx] || ""));
         found = true; break;
       }
     }
     if (found) return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
     else return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "IDが見つかりません" })).setMimeType(ContentService.MimeType.JSON);
   }
  
   // 🌟 メール再送アクション
   if (action === "send_email") {
     const sheet = getDbSheet();
     const data = sheet.getDataRange().getDisplayValues();
     const headers = data[0].map(h => h.trim());
     const rowIndex = findRowIndexBySysId(data, headers, requestData.id);
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象レコードが見つかりません" })).setMimeType(ContentService.MimeType.JSON);


     let appData = {};
     headers.forEach((h, idx) => { appData[h] = data[rowIndex - 1][idx]; });


     const yoseiId = appData["要請番号"];
     const result = sendCaseNotificationEmail_(appData, yoseiId, "【AW109 EMS 事案記録 再送】", "事案記録 (再送)");
     if (!result.sent) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: result.message })).setMimeType(ContentService.MimeType.JSON);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }
  
   if (action === "update_review_status") {
     const dSheet = getSheetFlexible(ss, ["DB_デブリーフィング", "デブリーフィングデータベース"]);
     if (!dSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが存在しません" })).setMimeType(ContentService.MimeType.JSON);
     const d = dSheet.getDataRange().getDisplayValues();
     const head = d[0].map(h => h.trim());
     const idIdx = head.indexOf("要請番号");
     const statusIdx = head.indexOf("レビューステータス");
     if (statusIdx === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "レビューステータス列が見つかりません" })).setMimeType(ContentService.MimeType.JSON);

     let rowIndex = -1;
     for (let i = 1; i < d.length; i++) {
       if (idIdx !== -1 && String(d[i][idIdx]).trim() === String(requestData.caseId).trim()) { rowIndex = i + 1; break; }
     }
     if (rowIndex === -1) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象のデブリーフィングが見つかりません" })).setMimeType(ContentService.MimeType.JSON);

     dSheet.getRange(rowIndex, statusIdx + 1).setValue(sanitizeInput(requestData.reviewStatus || "未確認"));
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "submit_debriefing") {
     const dSheet = getSheetFlexible(ss, ["DB_デブリーフィング", "デブリーフィングデータベース"]);
     if (!dSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "シートが存在しません" })).setMimeType(ContentService.MimeType.JSON);
     const p = requestData.data;
     const d = dSheet.getDataRange().getDisplayValues();
     const head = d[0].map(h => h.trim());
     const idIdx = head.indexOf("要請番号");
    
     let rowIndex = -1;
     for (let i = 1; i < d.length; i++) {
       if (idIdx !== -1 && String(d[i][idIdx]).trim() === String(p.caseId).trim()) { rowIndex = i + 1; break; }
     }


     let newRow = rowIndex !== -1 ? [...d[rowIndex - 1]] : new Array(head.length).fill("");
     while (newRow.length < head.length) newRow.push("");


     const setVal = (colName, val) => { const idx = head.indexOf(colName); if (idx !== -1) newRow[idx] = sanitizeInput(val); };


     setVal("要請番号", p.caseId);
     setVal("登録日時", Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm"));
     setVal("実施日", p.dDate || "");
     setVal("開始時間", p.dTime || "");
     setVal("終了時間", p.dDuration || "");
     setVal("実施場所", p.dPlace || "");
     setVal("実施方法", p.method || "");
     setVal("記録者", p.recorder || "");
     setVal("固定参加者", p.participants || "");
     setVal("外部参加者", p.extPt || "");
     setVal("OJT参加者", p.ojtPt || "");
     setVal("事前準備ギャップ", p.gap || "");
     setVal("チーム連携評価", p.team || "");
     setVal("マスタ・システム改善要望", p.masterGap || "");
     setVal("総括・次回アクション", p.action || "");
     setVal("タイムライン評価JSON", JSON.stringify(p.evals || []));
     setVal("多職種フィードバックJSON", JSON.stringify(p.feedbacks || []));
     setVal("レビューステータス", rowIndex !== -1 ? (d[rowIndex-1][head.indexOf("レビューステータス")] || "未確認") : "未確認");
     setVal("参加者一覧", p.allParticipants || "");
     setVal("システムデータ", JSON.stringify(p.systemData || {}));


     if (rowIndex !== -1) dSheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
     else dSheet.appendRow(newRow);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "fetch_library") {
     const libSs = SpreadsheetApp.openById(LIBRARY_DB_ID);
     const libSheet = getSheetFlexible(libSs, ["資料データベース"]);
     if (!libSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "資料データベースシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);

     const data = libSheet.getDataRange().getDisplayValues();
     const head = data[0].map(h => h.trim());
     const cTitle = head.indexOf("資料名");
     const cUrl = head.indexOf("ドライブURL");
     const cLarge = head.indexOf("【種類】（大分類）");
     const cMedium = head.indexOf("【対象領域】（中分類）");
     const cTags = head.indexOf("検索用キーワード（タグ）");
     const cDesc = head.indexOf("概要・特記事項");
     const cDate = head.indexOf("更新日");
     const cStatus = head.indexOf("公開ステータス");

     const docs = [];
     for (let i = 1; i < data.length; i++) {
       if (data[i].every(v => v === "")) continue;
       docs.push({
         rowIdx: i + 1,
         title: cTitle !== -1 ? data[i][cTitle] : "",
         url: cUrl !== -1 ? data[i][cUrl] : "",
         category: cLarge !== -1 ? data[i][cLarge] : "",
         domain: cMedium !== -1 ? data[i][cMedium] : "",
         tags: cTags !== -1 ? data[i][cTags] : "",
         desc: cDesc !== -1 ? data[i][cDesc] : "",
         date: cDate !== -1 ? data[i][cDate] : "",
         status: cStatus !== -1 ? data[i][cStatus] : ""
       });
     }
     return ContentService.createTextOutput(JSON.stringify({ status: "success", data: docs })).setMimeType(ContentService.MimeType.JSON);
   }


   if (action === "update_library_record") {
     if (!isAdminUser) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "管理者権限が必要です" })).setMimeType(ContentService.MimeType.JSON);
     const libSs = SpreadsheetApp.openById(LIBRARY_DB_ID);
     const libSheet = getSheetFlexible(libSs, ["資料データベース"]);
     if (!libSheet) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "資料データベースシートが見つかりません" })).setMimeType(ContentService.MimeType.JSON);

     const rowIdx = Number(requestData.rowIdx);
     const d = requestData.editData || {};
     if (!rowIdx || rowIdx < 2) return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "対象の行が不正です" })).setMimeType(ContentService.MimeType.JSON);

     const head = libSheet.getRange(1, 1, 1, libSheet.getLastColumn()).getDisplayValues()[0].map(h => h.trim());
     const currentVals = libSheet.getRange(rowIdx, 1, 1, head.length).getValues()[0];
     const setCol = (colName, val) => { const idx = head.indexOf(colName); if (idx !== -1) currentVals[idx] = sanitizeInput(val); };

     setCol("資料名", d.title || "");
     setCol("更新日", Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd"));
     setCol("概要・特記事項", d.desc || "");
     setCol("【種類】（大分類）", d.categoryRaw || "");
     setCol("【対象領域】（中分類）", d.domain || "");
     setCol("検索用キーワード（タグ）", d.tags || "");
     setCol("ドライブURL", d.url || "");
     setCol("公開ステータス", d.status || "");

     libSheet.getRange(rowIdx, 1, 1, currentVals.length).setValues([currentVals]);
     return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
   }


 } catch (err) {
   return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message })).setMimeType(ContentService.MimeType.JSON);
 }
}