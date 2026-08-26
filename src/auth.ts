import "dotenv/config";

import {
  PublicClientApplication,
  type AccountInfo,
} from "@azure/msal-node";

const clientId = process.env.MICROSOFT_CLIENT_ID;
const authority =
  process.env.MICROSOFT_AUTHORITY ??
  "https://login.microsoftonline.com/consumers";

if (!clientId) {
  throw new Error(
    "MICROSOFT_CLIENT_ID が設定されていません。npm run setup:azure を実行してください。",
  );
}

const scopes = [
  "https://graph.microsoft.com/Mail.ReadWrite",
  "https://graph.microsoft.com/Mail.Send",
];

const client = new PublicClientApplication({
  auth: {
    clientId,
    authority,
  },
});

let account: AccountInfo | null = null;

async function loginByDeviceCode(): Promise<string> {
  const result = await client.acquireTokenByDeviceCode({
    scopes,
    deviceCodeCallback: (response) => {
      console.log();
      console.log("Microsoftアカウントへログインしてください");
      console.log(response.message);
      console.log();
    },
  });

  if (!result) {
    throw new Error("Microsoftアカウントへのログインに失敗しました。");
  }

  if (!result.account) {
    throw new Error("Microsoftアカウント情報を取得できませんでした。");
  }

  account = result.account;

  return result.accessToken;
}

export async function getAccessToken(): Promise<string> {
  if (!account) {
    return loginByDeviceCode();
  }

  try {
    const result = await client.acquireTokenSilent({
      account,
      scopes,
    });

    return result.accessToken;
  } catch {
    return loginByDeviceCode();
  }
}
