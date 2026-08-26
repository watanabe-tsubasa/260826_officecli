# Office Mail Agent Sample

Outlookにメールで依頼すると、GitHub Codespaces上のAIエージェントが内容を読み取り、Office CLIを使ってPowerPointを作成し、完成したファイルを返信メールに添付して返すサンプルです。

プロトアウトスタジオ体験会

**「メールでOffice資料が届く！Office CLIで自分用AIエージェント開発 #ProtoOut」**

で利用する完成サンプルとして作成しています。

## このサンプルでできること

次のような流れを体験できます。

```text
Outlookへメールを送信
        ↓
Microsoft Graphでメールを取得
        ↓
AIエージェントが依頼内容を読み取る
        ↓
Office CLIでPowerPointを生成
        ↓
Microsoft Graphで返信
        ↓
PowerPointをメールで受け取る
```

たとえば、次のようなメールを送信します。

```text
件名:
[OFFICE-AI] 新商品の企画資料

本文:
新しいスマートウォッチの商品企画資料を
PowerPointで作成してください。

想定読者:
社内の商品企画会議

含めたい内容:
・商品のコンセプト
・想定ターゲット
・主な機能
・競合との差別化
・発売までのスケジュール

5枚程度のシンプルな資料にしてください。
```

AIエージェントがPowerPointを作成し、元のメールへの返信として `.pptx` ファイルを送信します。

## 使用技術

* GitHub Codespaces
* Node.js / TypeScript
* npm
* Microsoft Azure
* Microsoft Entra ID
* Microsoft Graph
* Azure CLI
* Outlook.com
* OpenCode
* Office CLI
* MCP
* ClaudeなどOpenCodeから利用できるAIモデル

## 必要なもの

利用前に以下を準備してください。

* GitHubアカウント
* Microsoftアカウント
* Outlook.comのメールアドレス
* ActiveなAzure Subscription
* OpenCodeから利用するAIサービス
* AIサービスのAPI Keyなど

体験会参加者向けの詳細な事前準備手順は、別途配布する資料を参照してください。

## セットアップ

### 1. GitHub Codespacesを起動

このリポジトリからGitHub Codespacesを起動します。

ターミナルが開いたら、依存パッケージをインストールします。

```bash
curl -sL https://aka.ms/InstallAzureCLIDeb | sudo bash
npm ci
```

### 2. 各ツールを確認

```bash
node --version
npm --version
az version
npx opencode --version
npx officecli --version
```

それぞれバージョン情報が表示されれば問題ありません。

## Azure / Microsoft Graphの準備

### 1. Azure CLIでログイン

```bash
az login --use-device-code
```

ターミナルに表示された案内に従い、Azureアカウントでログインします。

ログイン後、利用中のSubscriptionを確認します。

```bash
az account show -o table
```

複数のSubscriptionがある場合は、次のコマンドで確認できます。

```bash
az account list -o table
```

必要に応じて利用するSubscriptionを指定します。

```bash
az account set --subscription "SUBSCRIPTION_ID"
```

### 2. Microsoft Graph用アプリを作成

```bash
npm run setup:azure
```

このスクリプトでは、自分のMicrosoft Entra環境に

```text
ProtoOut Office Mail Agent
```

というアプリを登録します。

主に次の権限を利用します。

* `Mail.ReadWrite`
* `Mail.Send`

正常に完了すると `.env` が作成されます。

```text
MICROSOFT_CLIENT_ID=...
MICROSOFT_AUTHORITY=https://login.microsoftonline.com/consumers
MAIL_SUBJECT_PREFIX=[OFFICE-AI]
POLL_INTERVAL_MS=10000
```

`.env` はGitHubへコミットしないでください。

## Outlookとの接続確認

まず、AIエージェントを動かす前にMicrosoft Graphからメールを取得できるか確認します。

別のメールアドレスから、自分のOutlook.comアドレスへ次のようなメールを送信します。

```text
件名:
[OFFICE-AI] テスト

本文:
メール取得テストです。
```

メールは未読のままにしておきます。

続いて次を実行します。

```bash
npm run check:mail
```

初回実行時にはDevice Code認証が表示されます。

案内に従ってMicrosoftアカウントへログインし、

* メールの読み取り・書き込み
* メールの送信

に関する権限を許可します。

ターミナルにメールの件名や本文が表示されれば、Microsoft Graphとの接続は成功です。

## OpenCodeの準備

OpenCodeを起動します。

```bash
npx opencode
```

OpenCode上で

```text
/connect
```

を実行し、利用するAI Providerを設定します。

Claudeを利用する場合はAnthropicを選択し、事前に取得したAPI Keyを設定します。

モデルも必要に応じて選択してください。

接続確認として、次のようなコマンドを実行できます。

```bash
npx opencode run "1+1の答えだけを返してください。"
```

回答が返ればAIサービスとの接続は成功です。

## Office CLIとの接続確認

このリポジトリでは、OpenCodeからOffice CLIをMCP経由で利用する設定を `opencode.jsonc` に記載しています。

接続状況を確認します。

```bash
npx opencode mcp list
```

`officecli` が接続済みとして表示されれば成功です。

## AIエージェントを起動

すべての準備が完了したら、次を実行します。

```bash
npm start
```

次のような表示が出れば起動完了です。

```text
========================================
Office Mail Agent
========================================

Outlookの監視を開始します。
確認間隔: 10秒
```

この状態では、Outlookの受信箱を定期的に確認しています。

## PowerPointを作ってみる

別のメールアドレスから、自分のOutlook.comアドレスへメールを送信します。

### 件名

```text
[OFFICE-AI] 新商品の企画資料
```

### 本文例

```text
新しいスマートウォッチの商品企画資料を
PowerPointで作成してください。

想定読者:
社内の商品企画会議

含めたい内容:
・商品のコンセプト
・想定ターゲット
・主な機能
・競合との差別化
・発売までのスケジュール

5枚程度のシンプルな資料にしてください。
```

しばらくするとCodespaces上でメールが検出され、OpenCodeがOffice CLIを利用してPowerPointを生成します。

処理が完了すると、元のメールへの返信として

```text
result.pptx
```

が添付されます。

## npm scripts

主に次のコマンドを利用します。

```bash
npm run setup:azure
```

Microsoft Graphを利用するためのEntraアプリを作成します。

```bash
npm run check:mail
```

Outlookから対象メールを取得できるか確認します。

```bash
npm start
```

Outlookを監視し、Office資料を作成するAIエージェントを起動します。

## プロジェクト構成

```text
.
├── config/
│   └── graph-permissions.json
├── scripts/
│   └── setup-azure.sh
├── src/
│   ├── auth.ts
│   ├── graph.ts
│   ├── check-mail.ts
│   ├── agent.ts
│   └── index.ts
├── output/
├── opencode.jsonc
├── .env.example
├── .gitignore
├── package.json
├── package-lock.json
├── tsconfig.json
└── README.md
```

### `src/auth.ts`

Microsoft GraphへアクセスするためのOAuth認証を担当します。

Device Code Flowを利用してMicrosoftアカウントへログインし、アクセストークンを取得します。

### `src/graph.ts`

Microsoft Graphとの通信を担当します。

主に次の処理があります。

* Outlookの未読メールを取得
* メールを既読に変更
* 返信メールを作成
* Officeファイルを添付
* メールを送信

### `src/check-mail.ts`

Microsoft Graphとの接続確認用プログラムです。

AIエージェントを起動せず、Outlookからメールを取得できるか確認できます。

### `src/agent.ts`

AIエージェント部分です。

メール本文をOpenCodeへ渡し、Office CLIを利用してPowerPointを作成します。

AIエージェントへ与える役割やOffice資料の作成ルールも、このファイルで設定しています。

### `src/index.ts`

システム全体をつなぐメインプログラムです。

```text
Outlook
↓
Microsoft Graph
↓
AI Agent
↓
Office CLI
↓
Microsoft Graph
↓
Outlook
```

という一連の処理を実行します。

### `scripts/setup-azure.sh`

Azure CLIを利用してMicrosoft Entra IDへアプリを登録します。

### `config/graph-permissions.json`

Microsoft Graphで利用する権限を定義しています。

### `opencode.jsonc`

OpenCodeとOffice CLI MCPを接続する設定です。

## このサンプルを読んでみる

本リポジトリは完成済みのサンプルとして提供しています。

コードを最初からすべて入力する必要はありません。

まず実際に動かして、

```text
メールを送る
↓
PowerPointが返ってくる
```

ことを確認してみましょう。

その後、コードを読みながら

* Outlookのメールはどこで取得しているのか
* AIへ何を渡しているのか
* OpenCodeはどこで呼び出されているのか
* Office CLIとはどのようにつながっているのか
* 作成したファイルはどこでメールへ添付されているのか

を確認してみてください。

AIコーディングツールへ、

```text
src/graph.tsが何をしているか初心者向けに説明してください。
コードは変更しないでください。
```

のように質問しながらコードを読む方法もおすすめです。

## 改造してみる

基本動作を確認できたら、自分用のAIエージェントへ改造してみましょう。

たとえば、

* PowerPointの作成ルールを変更する
* Wordを作成できるようにする
* Excelを作成できるようにする
* CSV添付ファイルを読み取る
* 既存Officeファイルを修正する
* メールの返信から追加修正を受け付ける
* 自分の業務に合わせた資料テンプレートを利用する

といった拡張が考えられます。

AIコーディングツールを利用して、

```text
現在はPowerPointだけを生成しています。

メール本文に「Word」と書かれている場合は
Wordファイルを作成できるようにしたいです。

まず現在のコードを確認し、
変更方針を説明してください。
まだコードは変更しないでください。
```

のように相談してみるのもよいでしょう。

## 注意事項

* API Keyやアクセストークンなどの認証情報をGitHubへコミットしないでください
* `.env`はGitHubへコミットしないでください
* メールへ機密情報や個人情報を送信しないでください
* 本サンプルはハンズオン・学習用途を想定しています
* 実サービスとして公開する場合は、認証・認可・ジョブ管理・監査・エラー処理・レート制限などを追加してください
* 現在の基本実装では、メール添付するPowerPointは小さなファイルを想定しています
* Codespacesを停止するとメール監視も停止します

## License

tbd
