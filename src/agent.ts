import { spawn } from "node:child_process";
import { mkdir, access } from "node:fs/promises";
import { resolve } from "node:path";

export async function createPresentation(
  request: string,
): Promise<string> {
  const jobId = Date.now().toString();

  const outputDir = resolve("output", jobId);
  const outputPath = resolve(outputDir, "result.pptx");

  await mkdir(outputDir, {
    recursive: true,
  });

  const prompt = `
あなたはOffice資料作成専用のAIエージェントです。

ユーザーから届いた以下の依頼をもとに、
PowerPointファイルを作成してください。

必ずofficecliのMCPツールを利用してください。

出力先は必ず次のパスにしてください。

${outputPath}

基本ルール:
- 日本語で作成する
- 5枚程度を目安にする
- シンプルで読みやすいビジネス資料にする
- 外部画像は利用しない
- 添付メールで送信できるようファイルサイズを小さくする
- 指定した出力先以外へファイルを作成しない

ユーザーからの依頼:
---
${request}
---
`;

  const opencodePath = resolve(
    "node_modules",
    ".bin",
    "opencode",
  );

  await new Promise<void>((resolvePromise, rejectPromise) => {
    const child = spawn(
      opencodePath,
      [
        "run",
        prompt,
      ],
      {
        cwd: process.cwd(),
        stdio: "inherit",
        env: process.env,
      },
    );

    child.on("error", rejectPromise);

    child.on("exit", (code) => {
      if (code === 0) {
        resolvePromise();
      } else {
        rejectPromise(
          new Error(
            `OpenCodeが異常終了しました。exit code: ${code}`,
          ),
        );
      }
    });
  });

  try {
    await access(outputPath);
  } catch {
    throw new Error(
      `PowerPointが生成されませんでした: ${outputPath}`,
    );
  }

  return outputPath;
}
