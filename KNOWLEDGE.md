# 36協定書 PDF生成 KNOWLEDGE.md

> **NEVER-REPEAT 原則**: このファイルに記録されたバグ・落とし穴は、二度と同じ間違いをしない。
> 新しいバグを発見・修正したら **必ず即追記する**。修正後も削除しない。

---

## 📊 現在のスコア推移

| 日付 | 対象 | テンプレ全体平均 | 条文平均 | 備考 |
|------|------|----------------|----------|------|
| 2026-04-11 | test_all7forms（テストデータ） | 〜62% | 62.7% | ベースライン |
| 2026-04-11 | test_with_original_data（原本データ） | 72.4% | 81.6% | 原本データ使用後 |
| 2026-04-12 | test_with_original_data（S1〜S5修正後） | 89.2% | 90.9% | 様式9_3特別条項・氏名修正等 |
| 2026-04-12 | test_with_original_data（全ホワイトスペース正規化） | 89.2% | 94.2% | mask_variables()改行→スペース統合 |
| 2026-04-12 | test_with_original_data（BUG-009〜012修正後） | 92.4% | **96.0%** | 目標達成 ✅ |
| 2026-04-12 | test_with_original_data（BUG-013修正後） | **99.6%** | **99.8%** | 全様式99%台達成 🎉 |
| 2026-04-13 | test_with_original_data（S30〜S37追加）| **100.0%** | **100.0%** | **平均100%達成 🏆** |

**目標**: 条文平均 **95%以上** → **達成済み (100.0%)**

---

## 🐛 既知バグ一覧（NEVER REPEAT）

### BUG-001: 比較スクリプトで全角条番号が検出されない
- **症状**: 生成PDFの条文が「1条のみ」と表示される（第２条以降が未検出）
- **原因**: `split_arts()` の正規表現が半角数字のみ対応。weasyprintが `第１条`（全角）で出力する
- **修正**: `normalize_art_num()` で全角→半角変換を先に実行してから分割
- **発見日**: 2026-04-11
- **ファイル**: `compare_with_originals.py`, `test_with_original_data.py`

### BUG-002: マスク処理を先に実行すると条番号が破壊される
- **症状**: `第１条　甲は` → NAME マスクで `第１NAME` になり、条文ヘッダーが消える
- **原因**: `mask_variables()` 内の人名パターン `[ぁ-ん一-龥]{1,4}　[ぁ-ん一-龥]{1,4}` が `条　甲` にマッチ
- **修正**: **処理順を「正規化 → 分割 → マスク」に固定**。マスクを分割前に実行してはいけない
- **発見日**: 2026-04-11
- **ファイル**: `compare_with_originals.py`, `test_with_original_data.py`

### BUG-003: 内部参照の条番号で誤分割される
- **症状**: `第14条の期間における初日を起算日とする` の `第14条` が新しい条文ヘッダーとして認識される
- **原因**: 行頭判定なし + 助詞除外なし
- **修正**: `(?:^|\n)` で行頭限定 + `(?![のにをはがでもへとからまでよりについて])` で内部参照除外
- **発見日**: 2026-04-11
- **ファイル**: `compare_with_originals.py`, `test_with_original_data.py`

### BUG-004: `第8 条`（スペース付き）が正規化されない
- **症状**: C社原本PDF の `第8 条` が `normalize_art_num()` の正規表現にマッチせず未検出
- **原因**: PDFテキスト抽出ノイズで番号と「条」の間にスペースが挿入される
- **修正**: `normalize_art_num()` の冒頭に `re.sub(r'第(\d+)\s+条', r'第\1条', text)` を追加
- **発見日**: 2026-04-11 → **修正済み 2026-04-12 (S1)**
- **ファイル**: `compare_with_originals.py`, `test_with_original_data.py`

### BUG-005: 健康確保措置の文言が原本に存在しない
- **症状**: `_special_para()` が `健康確保措置：医師による面接指導。` を条文末尾に付加するが、原本B社/C社PDFにはこの文言がない → 第7条の類似度が大幅低下
- **原因**: 生成コードが独自に健康確保措置テキストを追加していた
- **修正**: `_special_para()` から `健康確保措置：{措置}。` を削除
- **発見日**: 2026-04-11 → **修正済み 2026-04-12 (S3)**
- **ファイル**: `pdf_generator.py` L240

### BUG-006: 署名欄「職種」ラベルが二重になる
- **症状**: 静的ラベル `（乙）労働者代表　職種` + 動的フィールド `{代表職}="職種"` → `（乙）労働者代表　職種職種` と出力される
- **原因**: HTMLテンプレートが `職種` をラベルとして持ちつつ、データフィールドにも `職種` が入る
- **修正**: 静的ラベルを `（乙）労働者代表` に変更し、動的 `{代表職}` がラベルを担う構造に
- **発見日**: 2026-04-11 → **修正済み 2026-04-12 (S4)**
- **ファイル**: `pdf_generator.py` 旧L322, L609, L638（3箇所を `replace_all=True` で一括修正）

### BUG-007: `_start_date()` が日付を `1日` 固定にする
- **症状**: C社の起算日が `令和8年2月21日` なのに `令和8年2月1日` と出力される
- **原因**: `_start_date()` が `月1日` をハードコード
- **修正**: `起算日_日` フィールドを追加（デフォルト `"1"`）。C社レコードに `"起算日_日": "21"` を追加
- **発見日**: 2026-04-11 → **修正済み 2026-04-12 (S5)**
- **ファイル**: `pdf_generator.py` `_start_date()`, `test_with_original_data.py` C社レコード

### BUG-008: PDF抽出テキストの空白・改行ノイズで類似度が下がる
- **症状**: weasyprintと原本PDFでテキスト抽出結果の改行/空白パターンが異なり、実質同じ内容でも類似度が低く計算される
- **原因**: PDFライブラリの差異（weasyprint生成 vs Adobe等で作成の原本）
- **修正**: `mask_variables()` の先頭で `[ \t　]+ → ' '` と `\n{2,} → \n` で正規化
- **発見日**: 2026-04-11 → **修正済み 2026-04-12 (S2)**
- **ファイル**: `compare_with_originals.py`, `test_with_original_data.py`

### BUG-009: `休日労働_日数` フィールドに二重プレフィックス付加

- **症状**: 生成PDFに「1か月に1か月に2日日」と出力される
- **原因**: `_holiday_table()` の `f"1か月に{日数}日"` でプレフィックス/サフィックスを付加しているが、フィールド値が既に「1か月に2日」という完全文字列を持つ
- **修正**: テンプレートを `f"{日数}<br>{時刻_html}"` に変更（プレフィックス/サフィックス削除）。`_driver_holiday_table()` の単行版も同様に修正
- **発見日**: 2026-04-12 → **修正済み 2026-04-12**
- **ファイル**: `pdf_generator.py` L235, L548

### BUG-010: `_driver_special_table()` に①②行構造がない

- **症状**: 生成PDFのドライバー特別条項テーブルが単行（flat）で出力されるが、原本は「①下記②に該当しない労働者（空行）」+「②自動車の運転の業務に従事する労働者（データ行）」の2行構造
- **原因**: `_driver_special_table()` が単純な1データ行しか持っていなかった
- **修正**: ①空行と②データ行を持つ2行構造に変更。`特別_延長時間_1日` フィールドを追加（通常OTと別値）
- **発見日**: 2026-04-12 → **修正済み 2026-04-12**
- **ファイル**: `pdf_generator.py` `_driver_special_table()`

### BUG-011: 様式10_2（F社）の `始業終業時刻` が 第4条休日テーブルと第10条で共有されていた

- **症状**: 第4条の休日時刻（運行予定表形式）と第10条の所定始業時刻（8時30分〜17時30分）を同一フィールドで管理できない
- **原因**: `_build_html_1nen_driver()` が `始業終業時刻` を第10条の始業終業に使っており、フィールドが競合
- **修正**: `始業終業時刻_10条` 専用フィールドを追加。`_build_html_1nen_driver()` でフォールバック付きで参照
- **発見日**: 2026-04-12 → **修正済み 2026-04-12**
- **ファイル**: `pdf_generator.py` `_build_html_1nen_driver()`, `test_with_original_data.py`

### BUG-013: 会社名マスクの `\S+` がS21後のCJK連結テキストで過剰マッチ

- **症状**: 様式10_2（F社）テンプレ全体類似度が99%台→59.3%に急落
- **原因**: `mask_variables()` のS21（CJK間スペース削除）後に `\S+(?:株式会社|...)` の `\S+` が文書途中から末尾まで2000文字以上にマッチ。スペースで区切られていたCJKテキストがS21で連結され、`\S+`が止まらなくなる
- **修正**: `\S+` → 上限付き `\S{0,25}` / `\S{0,20}` に変更
- **発見日**: 2026-04-12
- **ファイル**: `test_with_original_data.py`, `compare_with_originals.py` L380, L114

### BUG-014: 特別条項テーブルの列順序がPDF抽出で異なる（残存、構造的差異）

- **症状**: ドライバー様式（9_4/10_2）の特別条項テーブル条文（第4条/第5条）が97〜98%止まり
- **原因**: 原本PDFでは`満NUM歳以上の者NUMNUMNUM年延長する...延長する...（重複）`の列順で抽出されるが、生成PDFでは`満NUM歳以上の者延長する...NUMNUMNUM年NUMNUM年`の順で抽出される。PDF生成エンジン（weasyprint vs Adobe等）によるテーブルテキスト抽出順序の差異
- **修正**: マスクでの完全吸収は困難。97%以上のスコアで許容（構造的限界）
- **発見日**: 2026-04-12
- **ファイル**: `pdf_generator.py` `_driver_special_table()`, `_special_table()`

### BUG-012: テストデータが原本PDFの実際の値と大きく乖離していた

- **症状**: E社・F社のテストデータが汎用プレースホルダ値（"運送業務", "臨時の対応" 等）であり、原本との比較で低スコア
- **原因**: 初期テストデータ作成時に原本PDFの実際の値を転記していなかった
- **修正**: E社・F社・D社の各社テストデータを原本PDFから抽出した実際の値に更新
  - E社: 時間外_事由を多行テキストに更新、特別_延長時間_1日="7"、休日3行対応
  - F社: 起算日_月/日="7"/"16"、時間外1日="5"、1か月="42"、特別月="50"/年="400"
  - D社: 休日4行（営業/取付工事/製造/事務）追加
- **発見日**: 2026-04-12 → **修正済み 2026-04-12**
- **ファイル**: `test_with_original_data.py`

---

## ⚠️ 残存低スコア条文（96.0%達成後）

### 残存課題（改善余地あり）

| 様式 | 条文 | スコア | 主因 |
|------|------|--------|------|
| 様式9 第6条 | A社 | 85.8% | 休日テーブル列ヘッダー差異 |
| 様式9_2 第6条 | B社 | 80.5% | 休日テーブル複数行・列ヘッダー差異 |
| 様式9_4 第2条 | E社 | 87.0% | 時間外①行が多業種（自動車/事務/倉庫）なのに1行のみ対応 |
| 様式9_4 第3条 | E社 | 87.3% | 休日テーブルの期間列有無差異（原本は期間列なし） |
| 様式10 第9条 | D社 | 72.3% | 対象期間条文の文言差異 |
| 様式10_2 第3条 | F社 | 83.2% | 時間外①行の期間・起算日列が原本は異なるレイアウト |
| 様式10_2 第10条 | F社 | 76.3% | 第2章 始業終業・休憩時間の複数行表現差異 |

---

## 📐 設計原則（NEVER VIOLATE）

1. **処理順は必ず**: `normalize_art_num()` → `split_arts_raw()` → `mask_variables()` の順
   - マスクを先に実行すると条番号が破壊される（BUG-002）

2. **条文ヘッダー認識は行頭のみ**: `(?:^|\n)` アンカー必須
   - 内部参照 `第14条の期間における` をヘッダーと誤認する（BUG-003）

3. **全角・スペース付き条番号を正規化してから比較**:
   - `第８条`, `第8 条`, `第八条` はすべて `第8条` に統一済み

4. **_start_date() の日付はフィールドから取得**:
   - `起算日_日` フィールドがない場合のデフォルトは `"1"` だが、Excelデータ確認必須

5. **_special_para() の文言は原本PDFに合わせて維持**:
   - `健康確保措置：{措置}。` は原本B社/C社には存在しない → 削除済み（BUG-005）

---

## 🔧 テスト手順

### 類似度計測（推奨）
```bash
cd /home/kosuke_igarashi/my-project/yasuda-36kyotei
python3 test_with_original_data.py
```

### 全7様式生成テスト
```bash
cd /home/kosuke_igarashi/my-project/yasuda-36kyotei
python3 test_all7forms_accuracy.py
```

### 原本PDF直接比較
```bash
cd /home/kosuke_igarashi/my-project/yasuda-36kyotei
python3 compare_with_originals.py
```
（要: `test_output_all7/` に生成PDFが存在すること）

---

## 📧 メール下書き（mail_drafter.py）の落とし穴

### BUG-MAIL-001: Yahoo Mail で日本語PDF添付が「Untitled」表示される
- **症状**: `mail_drafter.save_draft()` で保存した下書きを Yahoo Mail webUI で開くと、PDF添付ファイルが `Untitled` または `noname` と表示される。本物のファイル名（例: `36協定書_株式会社サンプル商事_9_1.pdf`）が消える
- **原因**:
  1. Python の `email` モジュールは `add_header("Content-Disposition", "attachment", filename=...)` で日本語ファイル名を渡すと **RFC 2231 形式（`filename*=UTF-8''<urlencoded>`）でしか出力しない**
  2. **Yahoo Japan webUI は RFC 2231 をパースできない**。Content-Type の `name=` パラメータを優先的に見るが、ここに値が無いと "Untitled" にフォールバックする
  3. 旧コードでは `MIMEBase("application", "pdf", name=pdf_filename)` と日本語をそのまま渡していたため、Content-Type は文字化け、Disposition は RFC 2231 のみ → Yahoo がパース失敗
- **修正**: 「**RFC 2047 B-encoding ＋ RFC 2231 ＋ ASCIIフォールバック**」の3形式を**手動で同時出力**する。Pythonの自動エンコードはバイパスする
  ```python
  # NG: Python が自動で RFC 2231 のみ出力 → Yahoo Untitled
  part = MIMEBase("application", "pdf", name=pdf_filename)
  part.add_header("Content-Disposition", "attachment", filename=pdf_filename)

  # OK: 手動でヘッダー文字列を組み立て3形式同時出力
  encoded = f"=?UTF-8?B?{base64.b64encode(pdf_filename.encode('utf-8')).decode('ascii')}?="
  url_enc = quote(pdf_filename, safe="")
  del part["Content-Type"]
  del part["Content-Disposition"]
  part["Content-Type"] = f'application/pdf; name="{encoded}"'
  part["Content-Disposition"] = (
      f'attachment; filename="{encoded}"; filename*=UTF-8\'\'{url_enc}'
  )
  ```
- **検証方法**: 修正後は **必ず実IMAP接続でテスト送信し Yahoo Mail webUI で目視確認**する。MIMEヘッダーの単体テストだけでは検知できない（Yahoo側のパーサ挙動に依存）
- **検証コマンド**:
  ```bash
  cd ~/my-project/yasuda-36kyotei
  python3 -c "
  from mail_drafter import save_draft
  from excel_reader import read_excel
  from pdf_generator import generate_pdf
  secrets = {}
  with open('.streamlit/secrets.toml','r',encoding='utf-8') as f:
      for line in f:
          if '=' in line and not line.strip().startswith('#'):
              k,v = line.split('=',1); secrets[k.strip()] = v.strip().strip('\"')
  records = read_excel('demo_data/demo_36kyotei.xlsx')
  pdf = generate_pdf(records[0])
  res = save_draft(
      to_address=secrets['yahoo_user'],
      subject='【テスト】添付ファイル名検証',
      body='検証用テスト。削除可。',
      pdf_bytes=pdf,
      pdf_filename='36協定書_株式会社サンプル商事_TEST.pdf',
      imap_user=secrets['yahoo_user'],
      imap_password=secrets['yahoo_password'],
      from_address=secrets['yahoo_user'],
      idx=99,
  )
  print(res)
  "
  ```
- **発見日**: 2026-04-14
- **修正日**: 2026-04-14
- **ファイル**: `mail_drafter.py`
- **NEVER REPEAT**:
  - 日本語ファイル名を `MIMEBase(name=...)` のキーワードに直接渡してはいけない
  - `add_header("Content-Disposition", "attachment", filename=...)` だけに頼ってはいけない（RFC 2231のみで Yahoo 非対応）
  - メール添付を扱う変更を加えたら、**必ず実IMAP→Yahoo webUI目視確認** をリリース前チェックリストに入れる

### BUG-MAIL-002: IMAP接続が一時的失敗で1件落ちる
- **症状**: 複数件の下書き保存ループで、1件だけネットワーク瞬断・タイムアウトで失敗してスキップされる
- **原因**: `mail_drafter.save_draft()` がリトライ未実装で、1回失敗 = 即失敗扱い
- **修正**: 最大3回リトライ・2秒待機・**認証エラーは即時打ち切り**（無駄なリトライ防止）・接続エラーとIMAPエラーを別捕捉
- **発見日**: 2026-04-14（防御的に追加）
- **ファイル**: `mail_drafter.py`
- **NEVER REPEAT**: IMAP/SMTPなどネットワークI/Oは必ずリトライロジックを入れる。ただし認証失敗にはリトライしない

### BUG-MAIL-003: Yahoo Japan IMAP が海外IPからの接続を「海外からのアクセス制限」で拒否
- **症状**: ローカル（日本IP）では IMAP ログイン成功するのに、Streamlit Cloud（米国IP）にデプロイすると `IMAP認証/APPENDエラー: b'[AUTHENTICATIONFAILED] Incorrect username or password.'` で失敗する。パスワードは正しい。Yahoo側からセキュリティ警告メールも届かない（サイレント拒否）
- **原因**: Yahoo Japan は2025年8月から **「海外からのアクセス制限」** という独立した設定を導入し、全アカウントで順次「有効」化を進めている。これが有効だと **海外IPからのIMAP/POP/SMTP接続を一律ブロック** する。**「IMAP/POP/SMTPアクセスを許可」設定とは別の項目**で、こちらが「許可」でも海外制限が「有効」だと弾かれる
- **重要**: **Yahoo Japan には Gmail のような「アプリパスワード」「アクセスキー」機能は存在しない**（米国Yahoo!とは別運用）。通常のログインパスワードがそのまま使われる
- **解決策**: Yahoo!メール の **「海外からのアクセス制限」を「無効」に切り替える**
  1. https://mail.yahoo.co.jp にログイン
  2. 右上の歯車アイコン → 設定 → 左メニューから **「海外からのアクセス制限」** を選択
  3. 「有効」→ **「無効」** に切り替えて保存
  4. 即時反映 → 海外（米国IP含む）からのIMAP接続が通るようになる
  5. **同じパスワードのまま** Streamlit Cloud などの海外ホスティングからアクセス可能
- **二重設定の関係**:
  | 設定項目 | 場所 | 役割 |
  |---------|------|------|
  | IMAP/POP/SMTPアクセス | 設定 → IMAP/POP/SMTPアクセス | 外部メールソフト全般の許可（国内含む） |
  | **海外からのアクセス制限** | 設定 → 海外からのアクセス制限 | **海外IPだけを別レイヤーでブロック** |
  - **両方ON必須**（IMAP/POP/SMTPアクセス=許可 ＋ 海外からのアクセス制限=無効）
- **セキュリティ補強の推奨**: 「海外からのアクセス制限」を無効化すると全世界からアクセス可能になるため、Yahoo!ログインパスワードを強固なもの（16桁以上・記号含む）に変更することを推奨
- **検証方法**:
  ```bash
  # ローカルで動くか先に確認（地理的制限で本番と挙動が変わる罠を減らす）
  python3 -c "
  import imaplib
  m = imaplib.IMAP4_SSL('imap.mail.yahoo.co.jp', 993)
  m.login('USER@yahoo.co.jp', 'PASSWORD')
  print('login OK')
  m.logout()
  "
  ```
- **発見日**: 2026-04-14（安田さん36協定ツールのStreamlit Cloudデプロイ時）
- **解決日**: 2026-04-14（同日、設定変更で即解決）
- **ファイル**: Yahoo!メール設定（コード変更は不要）
- **NEVER REPEAT**:
  - Yahoo Japan メール連携機能の案件では **クライアントヒヤリング時に必ず「海外からのアクセス制限の無効化」を依頼**する
  - ローカル開発では日本IPで動くため気付かず、**本番（海外クラウド）に上げて初めて発覚**する典型パターン
  - Yahoo Japan には「アプリパスワード」「アクセスキー」機能はない（Gmail/米国Yahoo!とは異なる）。通常のログインパスワードを使い、海外制限を無効化するのが唯一の正解
  - 2025年8月以降、Yahoo!は **対応期日後にシステム側で自動的に「海外からのアクセス制限=有効」に変更** していく方針なので、過去動いていた連携が突然動かなくなる事象も発生する。デプロイ後に動かなくなったら最初にこの設定を疑う
  - Gmail は逆に「アプリパスワード」必須（2段階認証ON前提）。**メールサービスごとの認証ポリシーは全部違うので、ヒヤリング時に必ず確認する**

### BUG-DEPLOY-001: Streamlit Cloud Secrets で日本語キーが "Invalid format: please enter valid TOML." エラー
- **症状**: Streamlit Cloud の Advanced settings → Secrets に下記を貼ると弾かれる
  ```toml
  差出人名 = "あさひ労務管理センター"
  ```
- **原因**: TOML仕様では **非ASCII文字をキーに使う場合はクォーテッドキー（ダブルクォート囲み）必須**。Streamlit Cloudのバリデータは厳密にTOML仕様準拠
- **修正**: 日本語キーをすべて `"..."` で囲む
  ```toml
  # NG
  差出人名 = "あさひ労務管理センター"

  # OK
  "差出人名" = "あさひ労務管理センター"
  ```
- **正しいSecretsテンプレ（yasuda-36kyotei）**:
  ```toml
  password = "asahi"
  yahoo_user = "asahiroumu@yahoo.co.jp"
  yahoo_password = "bd19960605!"
  "差出人名" = "あさひ労務管理センター"
  "差出人所属" = "社会保険労務士法人あさひ労務管理センター"
  "差出人電話" = "029-8370-209"
  ```
- **発見日**: 2026-04-14
- **ファイル**: `.streamlit/secrets.toml.example`、Streamlit Cloud Advanced settings
- **NEVER REPEAT**:
  - 日本語（ASCII以外）キーは必ず `"..."` で囲む
  - Pythonの `tomllib` も実は厳密で、クォートなし日本語キーは `TOMLDecodeError` を返す
  - ローカル `secrets.toml` と Cloud Secrets は同じTOMLルールで統一する
  - 多言語キーが必要なら **ASCIIキー（例: `sender_name`）に統一する方が事故が少ない**（将来の改善案）

### メール添付改修時の必須チェックリスト

メール下書き機能（`mail_drafter.py`）に変更を加えるときは以下を**全て**実施してから完了とする：

| # | チェック項目 | コマンド/確認方法 |
|---|------------|----------------|
| 1 | MIMEヘッダー生成のユニットテスト | `python3 -c "from mail_drafter import _build_message; print(_build_message(...).decode())"` で `name="=?UTF-8?B?` と `filename="=?UTF-8?B?` の両方が出ることを目視 |
| 2 | 実IMAP接続テスト（自分宛） | 上記「検証コマンド」を実行し `下書き保存成功 attempts=1` を確認 |
| 3 | **Yahoo Mail webUIで添付ファイル名を目視確認** | https://mail.yahoo.co.jp で下書きフォルダを開き、添付名が `36協定書_◯◯.pdf` 形式で表示されることを確認 |
| 4 | PDFを開いて中身も確認 | クリックしてダウンロード→開いて内容が正しいことを確認 |
| 5 | 複数件（3件以上）の連続下書き保存テスト | リトライ動作・全件成功を確認 |

---

## 📁 関連ファイル

| ファイル | 役割 |
|---------|------|
| `pdf_generator.py` | HTML→PDF生成ロジック |
| `compare_with_originals.py` | 原本PDF vs 生成PDF 類似度計算 |
| `test_with_original_data.py` | 原本データで生成→即比較 |
| `test_all7forms_accuracy.py` | 全7様式の生成精度テスト |
| `PDF_TRACE_KNOWLEDGE.md` | PDFトレース手法・weasyprintノウハウ |
| `KNOWLEDGE.md` | **本ファイル**: バグ蓄積・NEVER-REPEAT |
| `mail_drafter.py` | Yahoo IMAP下書き保存・PDF添付（RFC2047三重エンコード） |
