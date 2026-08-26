import { getAccessToken } from "./auth.js";
import { listOfficeRequests } from "./graph.js";

const accessToken = await getAccessToken();

const messages = await listOfficeRequests(accessToken);

console.log();
console.log(`対象メール: ${messages.length}件`);
console.log();

for (const message of messages) {
  console.log("----------------------------------------");
  console.log(`件名: ${message.subject}`);
  console.log(
    `送信者: ${message.from?.emailAddress?.address ?? "不明"}`,
  );
  console.log();
  console.log(message.body.content);
  console.log("----------------------------------------");
}
