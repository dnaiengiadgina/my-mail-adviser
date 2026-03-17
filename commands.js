/*
 * commands.js
 * 送信イベントを捕捉し、確認ダイアログを表示する
 */

Office.initialize = function () {};

// messageSending イベントのハンドラ
Office.actions.associate("onSendHandler", onSendHandler);

function onSendHandler(event) {
  const item = Office.context.mailbox.item;

  // 宛先・件名・添付ファイルを並列取得
  Promise.all([
    getRecipients(item),
    getSubject(item),
    getAttachments(item)
  ]).then(function ([recipients, subject, attachments]) {

    const COMPANY_DOMAIN = "fujilogi.co.jp";
    const externalRecipients = recipients.filter(function (r) {
      return !r.emailAddress.toLowerCase().endsWith("@" + COMPANY_DOMAIN);
    });
    const hasExternal = externalRecipients.length > 0;
    const hasAttachments = attachments.length > 0;
    const noSubject = !subject || subject.trim() === "";

    // 確認が不要なケース（社内のみ・添付なし・件名あり）は素通り
    if (!hasExternal && !hasAttachments && !noSubject) {
      event.completed({ allowEvent: true });
      return;
    }

    // 確認ダイアログを表示
    const dialogUrl = "https://dnaiengiadgina.github.io/my-mail-adviser/dialog.html"
      + "?external=" + encodeURIComponent(JSON.stringify(externalRecipients.map(r => r.emailAddress)))
      + "&attachments=" + encodeURIComponent(JSON.stringify(attachments.map(a => a.name)))
      + "&noSubject=" + noSubject
      + "&total=" + recipients.length;

    Office.context.ui.displayDialogAsync(
      dialogUrl,
      { height: 60, width: 40, displayInIframe: false },
      function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          // ダイアログが開けない場合は送信を許可
          event.completed({ allowEvent: true });
          return;
        }

        const dialog = asyncResult.value;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          function (messageEvent) {
            dialog.close();
            const message = messageEvent.message;
            if (message === "send") {
              event.completed({ allowEvent: true });
            } else if (message === "cancel") {
              event.completed({ allowEvent: false });
            } else if (message.startsWith("delay:")) {
              // 送信遅延: キャンセルしてタスクパネルに通知
              const seconds = parseInt(message.split(":")[1], 10);
              event.completed({ allowEvent: false });
              // 遅延送信はタスクパネル側で管理
              Office.context.ui.messageParent("delay:" + seconds);
            }
          }
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          function () {
            // ダイアログが閉じられた場合はキャンセル
            event.completed({ allowEvent: false });
          }
        );
      }
    );
  }).catch(function () {
    event.completed({ allowEvent: true });
  });
}

// --- ヘルパー関数 ---

function getRecipients(item) {
  return new Promise(function (resolve) {
    const all = [];
    let pending = 3;

    function done(recipients) {
      all.push(...recipients);
      if (--pending === 0) resolve(all);
    }

    item.to.getAsync(function (r) { done(r.status === "succeeded" ? r.value : []); });
    item.cc.getAsync(function (r) { done(r.status === "succeeded" ? r.value : []); });
    item.bcc.getAsync(function (r) { done(r.status === "succeeded" ? r.value : []); });
  });
}

function getSubject(item) {
  return new Promise(function (resolve) {
    item.subject.getAsync(function (r) {
      resolve(r.status === "succeeded" ? r.value : "");
    });
  });
}

function getAttachments(item) {
  return new Promise(function (resolve) {
    item.attachments ? resolve(item.attachments) : resolve([]);
  });
}
