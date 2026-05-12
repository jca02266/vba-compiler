# VBA リファクタリング具体例：TaskScheduler マクロ

## 概要

このドキュメントは、**実際のレガシーコード（TaskScheduler_v1.vba）をリファクタリングして（TaskScheduler.vba + TaskScheduler_Core.vba）どのように改善されたか、具体的に示します**。

### 対象コード

- **リファクタリング前**: `sample/src/vba_legacy/TaskScheduler_v1.vba`（393 行の巨大 Sub）
- **リファクタリング後**: `sample/src/vba/TaskScheduler.vba` + `sample/src/vba/TaskScheduler_Core.vba`

---

## 問題点：リファクタリング前のコード

### 問題 1: 単一の巨大 Sub（393 行）

```vba
' TaskScheduler_v1.vba
Sub AutoScheduleTasks()
    ' Const 定義 (50 行)
    Const COL_CALENDAR_START_IDX As Long = 24
    Const ROW_HEADER_SHEET As Long = 3
    ' ... 15+ の定数定義
    
    ' 変数宣言 (30 行)
    Dim ws As Worksheet
    Dim screenUpdateState As Boolean
    ' ... 複数の変数
    
    ' Phase 1: Scan Locked Rows （30 行）
    For taskRow = 1 To numRows
        ' ネストされた If/For ループ
        ' 状態変数を更新しながら処理
    Next
    
    ' Phase 2: Schedule & Calculate Dependencies (280 行)
    For taskRow = 1 To numRows
        ' 階層ロジック
        ' 依存関係計算
        ' スケジューリング
        ' Level Max 更新
        ' ... すべてが混在
    Next
    
    ' Phase 3: Write Back
    rangeGrid.Value = gridData
End Sub
```

### 問題 2: 読み方が一方向（上から下へ）

レガシーコードを理解するには、**すべての行を順番に読む必要がある**：

```
1行目 ~ 50行目: Const 定義を暗記
51行目 ~ 85行目: 変数宣言を追跡
86行目: Range 計算開始
100行目: あ、ここで配列初期化か
...
250行目: あ、ここで Locked Row の処理だ
280行目: ここから本メインロジック？
...
```

**結果**: 
- 全 393 行を読んでようやく全体像がわかる
- 「ここは何をしているのか」が不明確
- 変更時に影響範囲が不明（どこが依存しているのか？）
- **デバッグが困難**（どの部分で状態が狂ったのか？）

### 問題 3: 状態管理の混在

```vba
' Phase 1 と Phase 2 で状態が共有される
Dim levelMaxFinish(0 To 100) As Long        ' Level ごとの最終日
Dim personUsage As Object                   ' 担当者ごとの工数使用量
Dim gridData As Variant                     ' スケジュール配列

' Phase 1 で以下を更新
For taskRow = 1 To numRows
    ' personUsage を更新
    personUsage(assigneeName) = newAllocArray
Next

' Phase 2 で上記の personUsage に依存しながら処理
For taskRow = 1 To numRows
    ' personUsage を読み取りながら
    capacity = maxDailyLoad - newAllocArray(dayIdx)
    ' ...同時に gridData も更新
Next
```

**問題**: 
- 複数の Variant 配列と Dictionary が共有される
- どこで何が変更されるのか追跡困難
- バグの原因を特定しづらい

---

## 改善策：リファクタリング後のコード

### 改善 1: 関心の分離と構造化（UDT 導入）

**リファクタリング前**:
```vba
Const COL_LEVEL_IDX As Long = 8
Const COL_OFFSET_IDX As Long = 9
Const COL_LOCK_IDX As Long = 13
Const COL_DURATION_IDX As Long = 15
Const COL_ASSIGNEE_IDX As Long = 17
Const CONFIG_ROW_START As Long = 8
Const CONFIG_ROW_END As Long = 14
Const CONFIG_COL_NAME As Long = 17
' ... バラバラに定義
```

**リファクタリング後**:
```vba
' TaskConfig: タスク関連の定数をグループ化
Type TaskConfig
    ROW_START As Long
    COL_LEVEL As Long
    COL_OFFSET As Long
    COL_LOCK As Long
    COL_DURATION As Long
    COL_ASSIGNEE As Long
    STR_LOCK_MARK As String
End Type

' CalendarConfig: カレンダー関連の定数をグループ化
Type CalendarConfig
    ROW_HEADER As Long
    ROW_HOLIDAY As Long
    COL_CALENDAR_START As Long
    STR_HOLIDAY_MARK As String
End Type

' AssigneeConfig: 担当者関連の定数をグループ化
Type AssigneeConfig
    COL_NAME As Long
    COL_LIMIT As Long
    ROW_START As Long
    ROW_END As Long
End Type
```

**メリット**:
- 定数が意味別に分類される → **概念が明確**
- パラメータ渡しが簡潔（`taskCfg` 1 つで全て）
- 変更が局所化（`TaskConfig` だけを変更すれば OK）

### 改善 2: Main Sub の簡潔化

**リファクタリング前** (SubProcedure 全体が 393 行):
```vba
Sub AutoScheduleTasks()
    ' Const（50 行）
    ' 変数宣言（30 行）
    ' Range 判定（20 行）
    ' データ読み込み（20 行）
    ' Config 読み込み（15 行）
    ' Phase 1: Locked Rows 処理（30 行）
    ' Phase 2: Scheduling（280 行）← ここが巨大！
    '   - Level ロジック
    '   - 依存関係計算
    '   - スケジューリングループ
    '   - Level Max 更新
    ' Write Back（5 行）
    ' Cleanup（10 行）
End Sub
```

**リファクタリング後** (メイン Sub が 203 行に短縮):
```vba
Sub AutoScheduleTasks()
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' ===== 初期化フェーズ（読みやすい）=====
    Dim calCfg As CalendarConfig
    calCfg = InitCalendarConfig()         ' ← Config は別の関数で初期化
    Dim taskCfg As TaskConfig
    taskCfg = InitTaskConfig()
    Dim assigneeCfg As AssigneeConfig
    assigneeCfg = InitAssigneeConfig()
    
    ' Performance Settings
    Dim screenUpdateState As Boolean
    screenUpdateState = Application.ScreenUpdating
    Application.ScreenUpdating = False
    ' ...
    
    ' ===== Range 判定フェーズ（独立した関数） =====
    Dim lastRow As Long
    lastRow = GetLastTaskRow(ws, taskCfg)
    Dim lastCol As Long
    lastCol = GetLastCalendarCol(ws, calCfg)
    Dim numRows As Long
    numRows = CalcNumRows(lastRow, taskCfg)
    Dim numDays As Long
    numDays = CalcNumDays(lastCol, calCfg)
    
    ' ===== データ読み込みフェーズ（独立した関数） =====
    Dim rangeTask As Range
    Set rangeTask = GetTaskRange(ws, taskCfg, lastRow)
    Dim taskDataFrame As Variant
    taskDataFrame = rangeTask.Value
    
    Dim rangeSchedule As Range
    Set rangeSchedule = GetScheduleRange(ws, taskCfg, calCfg, lastRow, lastCol)
    Dim scheduleGrid As Variant
    scheduleGrid = rangeSchedule.Value
    
    ' ===== Phase 1: Locked Rows スキャン =====
    Call ScanLockedRows(taskCfg, numRows, numDays, taskDataFrame, scheduleGrid, assigneeUsage)
    
    ' ===== Phase 2: Scheduling ループ（明確な構造） =====
    Dim maxLevel As Long
    maxLevel = GetMaxLevel(taskDataFrame, numRows, taskCfg)
    Dim levelMaxFinish() As Long
    Dim levelMaxFinishAlloc() As Double
    ReDim levelMaxFinish(0 To maxLevel)
    ReDim levelMaxFinishAlloc(0 To maxLevel)
    
    For taskRow = 1 To numRows
        isLocked = IsRowLocked(taskDataFrame, taskRow, taskCfg)
        assigneeName = GetAssigneeName(taskDataFrame, taskRow, taskCfg)
        currentLevel = GetTaskLevel(taskDataFrame, taskRow, taskCfg)
        
        If currentLevel = 1 Then
            Erase levelMaxFinish
            Erase levelMaxFinishAlloc
        End If
        
        parentFinishIdx = levelMaxFinish(currentLevel - 1)
        parentFinishAlloc = levelMaxFinishAlloc(currentLevel - 1)
        baseStartIdx = CalcBaseStartIdx(currentLevel, parentFinishIdx, parentFinishAlloc)
        
        If isLocked Then
            Call FindLockedTaskFinish(...)
            Call UpdateLevelFinish(...)
        Else
            Call ScheduleUnlockedTask(...)
            Call UpdateLevelFinish(...)
        End If
    Next taskRow
    
    ' ===== Write Back =====
    rangeSchedule.Value = scheduleGrid
    
Cleanup:
    Application.ScreenUpdating = screenUpdateState
End Sub
```

**視認性の改善**:

```
| メイン Sub（203 行）
├─ InitXxxConfig() ........ 設定初期化（専門化）
├─ GetLastTaskRow() ....... Range 判定（専門化）
├─ GetTaskRange() ......... Range 構築（専門化）
├─ ScanLockedRows() ....... Phase 1（専門化）
├─ ScheduleUnlockedTask() . Phase 2（専門化）
└─ UpdateLevelFinish() .... 状態更新（専門化）
```

**読み方が自由**:
1. 全体構造を知りたい → Main Sub を読む（203 行で OK）
2. Phase 1 の詳細を知りたい → `ScanLockedRows()` を読む
3. スケジューリングロジックを知りたい → `ScheduleUnlockedTask()` を読む
4. Level 計算を知りたい → `CalcBaseStartIdx()` を読む

---

## 複雑度の分散

### リファクタリング前

```
複雑度分布
|
| ████████████████████████████████████ (393行が1つに集中)
|
+────────────────────────────────────
  Main Sub (393行)
```

**問題**: 
- 1 つの関数で全複雑度を処理
- 認知負荷が高い
- バグの可能性が高い

### リファクタリング後

```
複雑度分布
|
| ███ (Main Sub: 203行)
| ██ (ScanLockedRows: 20行)
| ██ (ScheduleUnlockedTask: 50行)
| █  (Helper Functions: 各 5-10行)
|
+────────────────────────────────────
```

**メリット**:
- 複雑度が分散
- 各関数の認知負荷が低い
- テストしやすい（1 つずつ検証可能）

---

## 実際の改善例：Locked Row 処理

### レガシーコード（30 行が Phase 1 に埋まっている）

```vba
' Phase 1: Scan Locked Rows （埋もれた）
For taskRow = 1 To numRows
    assigneeName = Trim(metaData(taskRow, COL_ASSIGNEE_IDX))
    
    If assigneeName <> "" Then
        If Not personUsage.Exists(assigneeName) Then
            ReDim newAllocArray(1 To numDays) As Double
            personUsage.Add assigneeName, newAllocArray
        End If
        
        If UCase(Trim(metaData(taskRow, COL_LOCK_IDX))) = STR_LOCK_MARK Then
            newAllocArray = personUsage(assigneeName)
            For dayIdx = 1 To numDays
                cellVal = gridData(taskRow, dayIdx)
                existingAlloc = 0
                If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
                    existingAlloc = CDbl(cellVal)
                End If
                
                If existingAlloc > 0 Then
                    newAllocArray(dayIdx) = newAllocArray(dayIdx) + existingAlloc
                End If
            Next dayIdx
            personUsage(assigneeName) = newAllocArray
        End If
    End If
Next taskRow
```

**問題**:
- Phase 1 の処理が分からない（概要を読む必要）
- 「何をしているのか」が不明確
- デバッグ時に多くのコンテキストを保持する必要

### リファクタリング後（専門化された関数）

**メイン Sub**:
```vba
Call ScanLockedRows(taskCfg, numRows, numDays, taskDataFrame, scheduleGrid, assigneeUsage)
```

**TaskScheduler_Core.vba**:
```vba
Sub ScanLockedRows(taskCfg As TaskConfig, numRows As Long, numDays As Long, _
                   taskDataFrame As Variant, scheduleGrid As Variant, assigneeUsage As Object)
    Dim taskRow As Long, dayIdx As Long
    Dim assigneeName As String
    Dim newAllocArray() As Double
    
    For taskRow = 1 To numRows
        assigneeName = GetAssigneeName(taskDataFrame, taskRow, taskCfg)
        
        If assigneeName <> "" And IsRowLocked(taskDataFrame, taskRow, taskCfg) Then
            If Not assigneeUsage.Exists(assigneeName) Then
                ReDim newAllocArray(1 To numDays) As Double
                assigneeUsage.Add assigneeName, newAllocArray
            End If
            
            ' 負荷を集計
            Call AccumulateTaskAllocation(taskRow, numDays, scheduleGrid, assigneeUsage(assigneeName))
        End If
    Next taskRow
End Sub

' ヘルパー関数（責任が明確）
Sub AccumulateTaskAllocation(taskRow As Long, numDays As Long, scheduleGrid As Variant, allocArray() As Double)
    Dim dayIdx As Long
    Dim cellVal As Variant
    Dim existingAlloc As Double
    
    For dayIdx = 1 To numDays
        cellVal = scheduleGrid(taskRow, dayIdx)
        existingAlloc = 0
        If IsNumeric(cellVal) And Not IsEmpty(cellVal) Then
            existingAlloc = CDbl(cellVal)
        End If
        
        If existingAlloc > 0 Then
            allocArray(dayIdx) = allocArray(dayIdx) + existingAlloc
        End If
    Next dayIdx
End Sub
```

**メリット**:
1. **責任が明確** — ScanLockedRows は「Locked Row をスキャン」のみ
2. **テスト可能** — AccumulateTaskAllocation を単独でテスト
3. **変更が容易** — Locked Row の定義を変更？→ IsRowLocked() だけ変更
4. **再利用可能** — AccumulateTaskAllocation は他の場所からも呼び出し可能

---

## 複雑度が「分散」されることの利点

### ✅ メリット：複数の読み方が可能

**パターン 1: 全体構造を知りたい人**
```
1. Main Sub（203 行）を読む
   ↓ 3 分で全体像理解
```

**パターン 2: Phase 1 の詳細を知りたい人**
```
1. Main Sub の「Phase 1」セクション（1 行）
   ↓
2. ScanLockedRows() の実装（15 行）
   ↓ 1 分で完全理解
```

**パターン 3: スケジューリングロジックを知りたい人**
```
1. Main Sub の「Phase 2」セクション（20 行）
   ↓
2. ScheduleUnlockedTask() の実装（50 行）
   ↓
3. Micro-Task ロジック（20 行）
   ↓ 5 分で完全理解
```

**パターン 4: バグを修正したい人（変更箇所が明確）**
```
「Level 計算に bug がある」
  → CalcBaseStartIdx() だけを見る（5 行）
```

### ❌ レガシーコードの問題：読み方が 1 つ

```
全体を理解するには
  → すべての 393 行を順に読む必要がある
  ↓
「Locked Row の処理ってどこ？」
  → 250 行周辺を探す
  ↓
「でも Phase 1 と Phase 2 が混在している…」
  → 結局全体を追跡する羽目に
```

---

## 状態管理の分散

### レガシーコード（状態が共有される）

```
┌─────────────────────────────────┐
│ Sub AutoScheduleTasks()         │
├─────────────────────────────────┤
│ Dim metaData As Variant         │
│ Dim gridData As Variant         │◄─── Phase 1,2 で共有
│ Dim personUsage As Object       │
│ Dim levelMaxFinish() As Long    │
│ Dim capacityLimits As Object    │
├─────────────────────────────────┤
│ Phase 1: ロック行スキャン        │
│   metaData と gridData を読む     │
│   personUsage を更新             │
├─────────────────────────────────┤
│ Phase 2: スケジューリング        │
│   metaData, personUsage, 等を   │
│   同時に読み書き                 │
├─────────────────────────────────┤
│ Phase 3: Write Back             │
│   gridData をシートに書く        │
└─────────────────────────────────┘
```

**問題**:
- すべての状態がグローバル（Sub スコープ）
- Phase 1 と Phase 2 の依存関係が不明確
- デバッグ時に「どこで状態が狂ったのか」を追跡困難

### リファクタリング後（状態が局所化）

```
Main Sub
├─ Init Phase
│  ├─ calCfg ........ 初期化（使い回し）
│  ├─ taskCfg ....... 初期化（使い回し）
│  └─ assigneeCfg ... 初期化（使い回し）
│
├─ ScanLockedRows()
│  ├─ 入力: taskDataFrame, scheduleGrid
│  ├─ 出力: assigneeUsage（更新）
│  └─ 責任: Locked Row の負荷を集計
│
├─ Main Loop (Phase 2)
│  ├─ 入力: levelMaxFinish, levelMaxFinishAlloc（状態）
│  ├─ 各 Task で
│  │  ├─ GetTaskLevel()... Level を抽出
│  │  ├─ CalcBaseStartIdx()... 開始日を計算
│  │  └─ ScheduleUnlockedTask()... スケジュール
│  └─ 出力: scheduleGrid（更新）
│
└─ Write Back
   └─ scheduleGrid をシートに書く
```

**メリット**:
- 各関数の責任が明確
- 状態の流れが追跡可能
- テストがしやすい（状態を準備 → 関数呼び出し → 結果検証）

---

## チェックリスト：リファクタリングの効果測定

改善されたかどうかを判定：

- [ ] **読み方が複数ある** — Main Sub を読むだけで全体構造が分かるか？
- [ ] **複雑度が分散** — 各関数が 20 行程度か？
- [ ] **責任が明確** — 各関数の目的が関数名から分かるか？
- [ ] **パラメータが明確** — 「何が入力で、何が出力か」が関数シグネチャから分かるか？
- [ ] **テスト可能** — 1 つずつ単独でテスト（Excel シート依存なし）できるか？
- [ ] **変更が容易** — 要件変更時に修正箇所が明確か？

---

## 行数の比較

| 項目 | レガシー | リファクタリング後 |
|------|----------|------------------|
| Main Sub | 393 行 | 203 行（48%に短縮） |
| 定数定義 | 分散 | UDT で集約 |
| 関数数 | 1 | 20+ |
| 平均関数長 | 393 行 | 10-20 行 |
| 読むべき行数（全体理解） | 393 行 | 203 行（50%削減） |

---

## まとめ

### リファクタリングで何が改善されたのか

**表面的には「複雑になった」ように見えるが、実は：**

1. **複雑度が分散** — 1 つの巨大な関数が複数の小さな関数に分割
2. **読み方が自由** — 全体を読まなくても、興味のある部分だけ読める
3. **デバッグが容易** — 問題を特定しやすくなった
4. **テストが可能** — Excel シート依存なしで単体テスト可能
5. **変更が容易** — 要件変更時に修正範囲が明確

**結論**: 
```
「一見複雑に見えるが、実は理解しやすい」
= 複雑度が分散され、認知負荷が低い
```

このリファクタリングは、**REFACTORING_GUIDE.md で説明した「分割統治」と「関心の分離」の実践例**です。
