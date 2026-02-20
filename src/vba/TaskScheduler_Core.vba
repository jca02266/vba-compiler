Option Explicit

' Refactor #1: Extract Base Start Index Calculation
' 依存タスクの開始日を計算する (Calculate the start date for dependent tasks)
' Logic: If parentFinishAlloc < 0.5, start on parentFinishIdx. Else, parentFinishIdx + 1.
' ロジック: 親タスクの最終日の割り当てが0.5未満なら親の最終日と同日に開始、それ以外は翌日から開始する
Function CalcBaseStartIdx(currentLevel As Long, parentFinishIdx As Long, parentFinishAlloc As Double) As Long
    Dim baseStartIdx As Long
    baseStartIdx = 1

    If currentLevel > 1 Then
        If parentFinishIdx > 0 Then
            ' ロジック: 前レベルのタスクの最終日の割り当てが0.5未満の場合、
            ' 依存タスクは同日に開始可能です。
            ' それ以外（0.5以上）の場合、翌日に開始する必要があります。
            If parentFinishAlloc < 0.5 Then
                baseStartIdx = parentFinishIdx
            Else
                baseStartIdx = parentFinishIdx + 1
            End If
        End If
    End If
    
    CalcBaseStartIdx = baseStartIdx
End Function

' Refactor #2: Extract Max Daily Load Lookup
' 担当者の1日あたりの最大稼働工数を取得する (Get the maximum daily workload for an assignee)
' Logic: Look up assignee in limits dict, default to 1.0.
' ロジック: キャパシティ設定辞書から担当者を検索し、存在しなければデフォルト値 1.0 を返す
Function GetMaxDailyLoad(assigneeName As String, capacityLimits As Object) As Double
    Dim maxDailyLoad As Double
    maxDailyLoad = 1# ' Default 1.0

    If capacityLimits.Exists(assigneeName) Then
        maxDailyLoad = capacityLimits(assigneeName)
    End If

    GetMaxDailyLoad = maxDailyLoad
End Function

' Refactor #3: Extract Daily Allocation Calculation
' 1日あたりのタスク割り当て工数を計算する (Calculate daily task allocation)
' Logic: Handle 0.25 unit allocations and the 0.1 micro-task minimum based on capacity.
' ロジック: 残り容量と工数から0.25単位の標準割り当てを計算する。マイクロタスクの場合は最低0.1を保証する
Function CalcDailyAllocation(capacity As Double, remaining As Double, isMicroTask As Boolean) As Double
    Dim dailyAlloc As Double
    dailyAlloc = 0
    
    Dim maxUnits As Long
    maxUnits = Int(capacity / 0.25)
    
    Dim neededUnits As Long
    Dim allocateUnits As Long
    
    If isMicroTask Then
         ' Special Case: Micro-Task (e.g. 0.1, 0.05)
         ' Only allocate if capacity >= 0.1
         ' If allocated, it consumes 0.1 and we are done.
         If capacity >= 0.1 Then
            dailyAlloc = 0.1
         End If
    Else
         ' Standard Logic (0.25 units)
         neededUnits = Int((remaining / 0.25) + 0.5)
         
         If maxUnits > 0 And neededUnits > 0 Then
            allocateUnits = neededUnits
            If allocateUnits > maxUnits Then allocateUnits = maxUnits
            dailyAlloc = allocateUnits * 0.25
         End If
    End If
    
    CalcDailyAllocation = dailyAlloc
End Function

' Refactor #4: Extract Level Finish Update Logic
' 指定された階層レベルの最終完了日（最も遅い日）とその日の工数を更新する (Update the maximum finish date and its allocation for a given level)
' Logic: Update the max finish index and allocation arrays based on the task's completion.
' ロジック: タスクの完了日が現在の最大値より後ろなら更新、同日なら工数が多い方で上書きする
Sub UpdateLevelFinish(currentLevel As Long, taskFinishIdx As Long, taskFinishAlloc As Double, ByRef levelMaxFinish, ByRef levelMaxFinishAlloc)
    If currentLevel > 0 Then
        If taskFinishIdx > levelMaxFinish(currentLevel) Then
            levelMaxFinish(currentLevel) = taskFinishIdx
            levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
        ElseIf taskFinishIdx = levelMaxFinish(currentLevel) Then
            If taskFinishAlloc > levelMaxFinishAlloc(currentLevel) Then
                levelMaxFinishAlloc(currentLevel) = taskFinishAlloc
            End If
        End If
    End If
End Sub



