function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  handleNotificationEdit(e, sheet, sheetName, row, col);

  if (sheetName !== "業務履歴" || col !== 2 || row < 2) {
    return;
  }

  const targetCell = sheet.getRange(row, 3);
  targetCell.clearDataValidations();

  const keyword = e.range.getValue();
  if (!keyword) {
    targetCell.clearContent();
    return;
  }

  const masterSheet = e.source.getSheetByName("マスタ");
  if (!masterSheet) {
    return;
  }

  const lastRow = masterSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }

  const masterValues = masterSheet.getRange(2, 1, lastRow - 1, 2).getValues();
  const matched = Array.from(new Set(
    masterValues
      .filter(([code, name]) => name === keyword && code !== "")
      .map(([code]) => code)
  ));

  if (matched.length === 0) {
    targetCell.clearContent();
    return;
  }

  const dvRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(matched, true)
    .setAllowInvalid(false)
    .build();
  targetCell.setDataValidation(dvRule);
}

function handleNotificationEdit(e, sheet, sheetName, row, col) {
  const targetSheets = ["開示", "下書き", "サンプル"];
  if (!targetSheets.includes(sheetName)) {
    return;
  }

  if (col === 6) {
    const props = PropertiesService.getScriptProperties();
    const pending = props.getProperty("pendingRows");
    let pendingRows = [];

    try {
      const parsed = pending ? JSON.parse(pending) : [];
      pendingRows = Array.isArray(parsed) ? parsed : [];
    } catch (error) {
      console.error("pendingRows の読み込みに失敗したため初期化します。", error);
      pendingRows = [];
    }

    const alreadyQueued = pendingRows.some(item => item.sheetName === sheetName && item.row === row);
    if (!alreadyQueued) {
      pendingRows.push({ sheetName, row });
      props.setProperty("pendingRows", JSON.stringify(pendingRows));
      ScriptApp.newTrigger("generateChatMessages")
        .timeBased()
        .after(30 * 60 * 1000)
        .create();
    }
  }

  if (col === 7) {
    const status = sheet.getRange(row, 7).getValue();
    if (typeof status === "string" && status.trim() === "提出済み") {
      sheet.getRange(row, 12).clearContent();
    }
  }
}

function generateChatMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const master = ss.getSheetByName("チャットグループマスタ");
  if (!master) {
    console.log("チャットグループマスタが見つかりませんでした。");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  let rows;
  try {
    const parsed = JSON.parse(props.getProperty("pendingRows") || "[]");
    rows = Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    console.error("pendingRows のパースに失敗しました。", error);
    rows = [];
  }

  if (!rows.length) {
    props.deleteProperty("pendingRows");
    console.log("処理対象の行が存在しません。");
    return;
  }

  const masterData = master.getDataRange().getValues();

  const rowInfoList = [];

  rows.forEach(({ sheetName, row }) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      console.log(`❌ シート "${sheetName}" が見つからないためスキップします。`);
      return;
    }

    const isSampleSheet = sheetName === "サンプル";
    const rawStatus = sheet.getRange(row, 7).getValue();
    const status = typeof rawStatus === "string" ? rawStatus.trim() : "";

    if (isSampleSheet) {
      if (status === "提出済み" || status === "戻し済み") {
        console.log(`Row ${row} (${sheetName}) は提出済みまたは戻し済みのためスキップ\n`);
        return;
      }
    } else {
      if (status !== "") {
        console.log(`Row ${row} (${sheetName}) はステータスが空白ではないためスキップ\n`);
        return;
      }
    }

    const rawColumnF = sheet.getRange(row, 6).getValue();
    const normalizedColumnF =
      typeof rawColumnF === "string" ? rawColumnF.trim() : rawColumnF;
    const hasColumnFValue = !isBlankValue(normalizedColumnF);

    const rawColumnH = sheet.getRange(row, 8).getValue();
    const rawColumnI = sheet.getRange(row, 9).getValue();
    const isColumnHBlank = isBlankValue(rawColumnH);
    const isColumnIBlank = isBlankValue(rawColumnI);

    const client = sheet.getRange(row, 4).getValue()?.toString().trim();
    const project = sheet.getRange(row, 5).getValue()?.toString().trim();

    if (!isSampleSheet && !hasColumnFValue) {
      console.log(`Row ${row} (${sheetName}) はF列が空白のためスキップ\n`);
      return;
    }

    if (sheetName === "下書き" && !isColumnHBlank) {
      console.log(`Row ${row} (${sheetName}) はH列が空白ではないためスキップ\n`);
      return;
    }

    if (isSampleSheet) {
      if (isBlankValue(project)) {
        console.log(`Row ${row} (${sheetName}) はE列が空白のためスキップ\n`);
        return;
      }

      if (!isColumnIBlank) {
        console.log(`Row ${row} (${sheetName}) はI列が空白ではないためスキップ\n`);
        return;
      }
    }

    let url = "";
    if (!isSampleSheet) {
      if (typeof normalizedColumnF === "string") {
        url = normalizedColumnF;
      } else {
        url = String(normalizedColumnF).trim();
      }
    }
    const columnFValueForSample =
      isSampleSheet && typeof rawColumnF === "string"
        ? rawColumnF.trim()
        : isSampleSheet
        ? rawColumnF
        : "";
    const rawQuantity = sheetName === "サンプル" ? rawStatus : null;

    console.log(`--- ${sheetName} Row ${row} 処理開始 ---`);
    console.log(`クライアント: ${client}, 案件: ${project}, ステータス: ${status}`);

    const {
      mentions,
      chatGroupUrl,
      matched,
      fallbackMatched
    } = resolveChatInfo(masterData, client, project);

    if (!fallbackMatched) {
      console.log(`❌ クライアント名 "${client}" に該当する行がマスタに存在しません。通知スキップ。\n`);
      return;
    }

    if (!matched) {
      console.log(`⚠️ 案件名 "${project}" はマスタに見つからず、クライアント名一致の最初の行を使用します。`);
    }

    console.log(`✅ 使用するメンション:\n${mentions.join("\n")}`);
    console.log(`✅ 使用するチャットワークURL: ${chatGroupUrl || "(なし)"}`);

    rowInfoList.push({
      sheetName,
      row,
      client,
      project,
      url: sheetName === "サンプル" ? "" : url,
      variant: sheetName === "サンプル" ? columnFValueForSample : "",
      quantity:
        sheetName === "サンプル" && status !== "提出済み" && status !== "戻し済み"
          ? rawQuantity
          : "",
      mentions,
      chatGroupUrl,
      status
    });

    console.log(`--- ${sheetName} Row ${row} 処理終了 ---\n`);
  });

  const groupedBySheet = rowInfoList.reduce((acc, info) => {
    if (!acc[info.sheetName]) {
      acc[info.sheetName] = [];
    }
    acc[info.sheetName].push(info);
    return acc;
  }, {});

  Object.keys(groupedBySheet).forEach(sheetName => {
    const items = groupedBySheet[sheetName];

    if (sheetName === "開示") {
      const groupedByClient = items.reduce((acc, info) => {
        const key = info.client || "";
        if (!acc[key]) {
          acc[key] = [];
        }
        acc[key].push(info);
        return acc;
      }, {});

      Object.values(groupedByClient).forEach(groupItems => {
        if (groupItems.length > 1) {
          const message = buildGroupedDisclosureMessage(groupItems);
          if (!message) {
            console.log("⚠️ メッセージが生成できませんでした。スキップします。\n");
            return;
          }

          console.log("📤 最終送信メッセージ:\n" + message);
          sendToGoogleChat(message);
        } else {
          const singleMessage = buildChatMessage(groupItems[0]);
          if (!singleMessage) {
            console.log("⚠️ メッセージが生成できませんでした。スキップします。\n");
            return;
          }

          console.log("📤 最終送信メッセージ:\n" + singleMessage);
          sendToGoogleChat(singleMessage);
        }
      });
    } else {
      items.forEach(item => {
        const message = buildChatMessage(item);
        if (!message) {
          console.log("⚠️ メッセージが生成できませんでした。スキップします。\n");
          return;
        }

        console.log("📤 最終送信メッセージ:\n" + message);
        sendToGoogleChat(message);
      });
    }
  });

  props.deleteProperty("pendingRows");
}

function isBlankValue(value) {
  if (value === null || value === undefined) {
    return true;
  }

  if (typeof value === "string") {
    return value.trim() === "";
  }

  return false;
}

function resolveChatInfo(masterData, client, project) {
  let mentions = [];
  let chatGroupUrl = "";
  let fallbackMentions = [];
  let fallbackUrl = "";
  let matched = false;
  let fallbackMatched = false;

  for (let i = 1; i < masterData.length; i++) {
    const mClient = masterData[i][0]?.toString().trim();
    const mProject = masterData[i][1]?.toString().trim();

    if (mClient === client) {
      if (!fallbackMatched) {
        fallbackUrl = masterData[i][3]?.toString().trim();
        for (let j = 4; j < masterData[i].length; j++) {
          const mention = masterData[i][j];
          if (mention && mention.toString().trim() !== "") {
            fallbackMentions.push(mention.toString().trim());
          }
        }
        fallbackMatched = true;
      }

      if (mProject === project) {
        chatGroupUrl = masterData[i][3]?.toString().trim();
        for (let j = 4; j < masterData[i].length; j++) {
          const mention = masterData[i][j];
          if (mention && mention.toString().trim() !== "") {
            mentions.push(mention.toString().trim());
          }
        }
        matched = true;
        break;
      }
    }
  }

  if (!matched) {
    mentions = fallbackMentions;
    chatGroupUrl = fallbackUrl;
  }

  return { mentions, chatGroupUrl, matched, fallbackMatched };
}

function buildChatMessage({ sheetName, project, url, variant, quantity, mentions, chatGroupUrl }) {
  if (!project) {
    return "";
  }

  const lines = [];

  if (mentions && mentions.length > 0) {
    lines.push(...mentions);
  }

  lines.push("お世話になっております。");

  if (sheetName === "開示") {
    lines.push(`${project}の実施希望者をリストに追加させていただきました！`);
    lines.push("可否確認のほどよろしくお願いいたします！");
    if (url) {
      lines.push(url);
    }
  } else if (sheetName === "下書き") {
    lines.push(`${project}の下書きを提出させていただきます！`);
    lines.push("ご確認のほどよろしくお願いいたします！");
    if (url) {
      lines.push(url);
    }
  } else if (sheetName === "サンプル") {
    const trimmedVariant = variant && variant.toString().trim() !== "" ? variant.toString().trim() : "";
    const detail = trimmedVariant ? `の${trimmedVariant}を` : "を";
    let quantityText = "";
    if (quantity !== null && quantity !== undefined && quantity !== "") {
      const qStr = quantity.toString().trim();
      if (qStr !== "" && qStr !== "提出済み" && qStr !== "戻し済み") {
        quantityText = `${qStr}個`;
      }
    }
    lines.push(`${project}${detail}${quantityText}発送お願いできますでしょうか！`);
    lines.push("よろしくお願いいたします！");
  }

  if (sheetName !== "サンプル" && url === "") {
    // no-op, url already handled above
  }

  if (chatGroupUrl) {
    lines.push(`チャットワークグループ:${chatGroupUrl}`);
  }

  return lines.join("\n");
}

function buildGroupedDisclosureMessage(items) {
  if (!items.length) {
    return "";
  }

  const validItems = items.filter(item => item.project);
  if (!validItems.length) {
    return "";
  }

  const count = validItems.length;
  const uniqueMentions = Array.from(
    new Set(
      validItems.reduce((acc, item) => {
        if (item.mentions && item.mentions.length) {
          acc.push(...item.mentions);
        }
        return acc;
      }, [])
    )
  );

  const chatGroupUrls = Array.from(
    new Set(validItems.map(item => item.chatGroupUrl).filter(Boolean))
  );

  if (chatGroupUrls.length > 1) {
    console.log("⚠️ 同一クライアントで異なるチャットワークURLが検出されました。最初のURLを使用します。");
  }

  const chatGroupUrl = chatGroupUrls[0] || "";
  const lines = [];

  if (uniqueMentions.length) {
    lines.push(...uniqueMentions);
  }

  lines.push("お世話になっております。");
  lines.push(`下記${count}件の実施希望者をリストに追加させていただきました！`);
  lines.push("可否確認のほどよろしくお願いいたします！");
  lines.push("");

  validItems.forEach((item, index) => {
    const marker = getCircledNumber(index + 1);
    lines.push(`${marker}${item.project}`);
    if (item.url) {
      lines.push(item.url);
    }
  });

  if (chatGroupUrl) {
    lines.push("");
    lines.push(`チャットワークグループ:${chatGroupUrl}`);
  }

  return lines.join("\n");
}

function getCircledNumber(index) {
  const circledNumbers = [
    "①",
    "②",
    "③",
    "④",
    "⑤",
    "⑥",
    "⑦",
    "⑧",
    "⑨",
    "⑩",
    "⑪",
    "⑫",
    "⑬",
    "⑭",
    "⑮",
    "⑯",
    "⑰",
    "⑱",
    "⑲",
    "⑳"
  ];

  if (index >= 1 && index <= circledNumbers.length) {
    return circledNumbers[index - 1];
  }

  return `${index}.`;
}

function sendToGoogleChat(message) {
  const webhookUrl = "https://chat.googleapis.com/v1/spaces/AAQAIKpx4ug/messages?key=AIzaSyDdI0hCZtE6vySjMm-WEfRq3CPzqKqqsHI&token=sZjeMOls7aB6jji8CdvXjQlXMlX-jPDmyplFk1FxQAg";
  const payload = {
    text: message
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    const response = UrlFetchApp.fetch(webhookUrl, options);
    const code = response.getResponseCode();
    const body = response.getContentText();

    if (code >= 200 && code < 300) {
      console.log("✅ Google Chatへの通知成功");
    } else {
      console.error(`❌ 通知失敗（ステータス: ${code}）\nレスポンス: ${body}`);
    }
  } catch (error) {
    console.error("❗ 送信中にエラー発生:", error);
  }
}
