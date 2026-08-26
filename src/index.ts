import "dotenv/config";

import { getAccessToken } from "./auth.js";

import {
  listOfficeRequests,
  markAsRead,
  replyWithAttachment,
} from "./graph.js";

import {
  createPresentation,
} from "./agent.js";

const pollInterval =
  Number(process.env.POLL_INTERVAL_MS) || 10_000;

console.log("========================================");
console.log("Office Mail Agent");
console.log("========================================");
console.log();
console.log("Outlookの監視を開始します。");
console.log(
  `確認間隔: ${pollInterval / 1000}秒`,
);
console.log();

async function processMessages(): Promise<void> {
  const accessToken = await getAccessToken();

  const messages =
    await listOfficeRequests(accessToken);

  if (messages.length === 0) {
    return;
  }

  console.log(
    `${messages.length}件の依頼メールを検出しました。`,
  );

  for (const message of messages) {
    console.log();
    console.log("----------------------------------------");
    console.log(`件名: ${message.subject}`);
    console.log(
      `送信者: ${
        message.from?.emailAddress?.address ?? "不明"
      }`,
    );
    console.log("----------------------------------------");

    try {
      /*
       * 同じメールを繰り返しAIへ送らないよう、
       * 処理開始時点で既読にします。
       *
       * 処理に失敗した場合はOutlookから
       * 手動で未読へ戻すことで再実行できます。
       */
      await markAsRead(
        accessToken,
        message.id,
      );

      console.log("AIエージェントへ依頼します...");

      const outputPath =
        await createPresentation(
          message.body.content,
        );

      console.log();
      console.log(
        `PowerPointを作成しました: ${outputPath}`,
      );

      /*
       * Office資料生成中に時間が経過する可能性があるため、
       * 返信前に最新のaccess tokenを取得します。
       */
      const sendToken =
        await getAccessToken();

      console.log("返信メールを作成します...");

      await replyWithAttachment(
        sendToken,
        message.id,
        outputPath,
      );

      console.log("返信メールを送信しました。");
    } catch (error) {
      console.error();
      console.error(
        "メールの処理中にエラーが発生しました。",
      );
      console.error(error);
      console.error();
      console.error(
        "再実行する場合はOutlookで対象メールを未読へ戻してください。",
      );
    }
  }
}

while (true) {
  try {
    await processMessages();
  } catch (error) {
    console.error(
      "メール確認中にエラーが発生しました。",
    );
    console.error(error);
  }

  await new Promise((resolvePromise) =>
    setTimeout(
      resolvePromise,
      pollInterval,
    ),
  );
}
