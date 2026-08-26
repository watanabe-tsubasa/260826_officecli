import "dotenv/config";

import { readFile } from "node:fs/promises";
import { basename } from "node:path";

const GRAPH_BASE_URL = "https://graph.microsoft.com/v1.0";

export type GraphMessage = {
  id: string;
  subject: string;
  receivedDateTime: string;
  body: {
    contentType: string;
    content: string;
  };
  from?: {
    emailAddress?: {
      name?: string;
      address?: string;
    };
  };
};

type GraphMessageList = {
  value: GraphMessage[];
};

async function graphRequest<T>(
  accessToken: string,
  path: string,
  init: RequestInit = {},
): Promise<T> {
  const headers = new Headers(init.headers);

  headers.set("Authorization", `Bearer ${accessToken}`);

  if (init.body && !headers.has("Content-Type")) {
    headers.set("Content-Type", "application/json");
  }

  const response = await fetch(`${GRAPH_BASE_URL}${path}`, {
    ...init,
    headers,
  });

  if (!response.ok) {
    const body = await response.text();

    throw new Error(
      `Microsoft Graph API error: ${response.status} ${response.statusText}\n${body}`,
    );
  }

  if (response.status === 202 || response.status === 204) {
    return undefined as T;
  }

  const body = await response.text();

  if (!body) {
    return undefined as T;
  }

  return JSON.parse(body) as T;
}

export async function listOfficeRequests(
  accessToken: string,
): Promise<GraphMessage[]> {
  const params = new URLSearchParams({
    "$filter": "isRead eq false",
    "$select": "id,subject,receivedDateTime,body,from",
    "$top": "50",
  });

  const result = await graphRequest<GraphMessageList>(
    accessToken,
    `/me/mailFolders/inbox/messages?${params.toString()}`,
    {
      headers: {
        Prefer: 'outlook.body-content-type="text"',
      },
    },
  );

  const prefix = process.env.MAIL_SUBJECT_PREFIX ?? "[OFFICE-AI]";

  return result.value.filter((message) =>
    message.subject.startsWith(prefix),
  );
}

export async function markAsRead(
  accessToken: string,
  messageId: string,
): Promise<void> {
  await graphRequest(
    accessToken,
    `/me/messages/${encodeURIComponent(messageId)}`,
    {
      method: "PATCH",
      body: JSON.stringify({
        isRead: true,
      }),
    },
  );
}

export async function replyWithAttachment(
  accessToken: string,
  messageId: string,
  filePath: string,
): Promise<void> {
  const draft = await graphRequest<GraphMessage>(
    accessToken,
    `/me/messages/${encodeURIComponent(messageId)}/createReply`,
    {
      method: "POST",
      body: JSON.stringify({
        comment:
          "Office CLIを利用して資料を作成しました。添付ファイルをご確認ください。",
      }),
    },
  );

  const file = await readFile(filePath);

  const maxSize = Math.floor(2.8 * 1024 * 1024);

  if (file.byteLength > maxSize) {
    throw new Error(
      `添付ファイルが大きすぎます: ${file.byteLength} bytes\n` +
        "基本課題では約2.8MB以下のPowerPointを利用してください。",
    );
  }

  await graphRequest(
    accessToken,
    `/me/messages/${encodeURIComponent(draft.id)}/attachments`,
    {
      method: "POST",
      body: JSON.stringify({
        "@odata.type": "#microsoft.graph.fileAttachment",
        name: basename(filePath),
        contentType:
          "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        contentBytes: file.toString("base64"),
      }),
    },
  );

  await graphRequest(
    accessToken,
    `/me/messages/${encodeURIComponent(draft.id)}/send`,
    {
      method: "POST",
    },
  );
}
