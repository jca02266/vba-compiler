import { runVBATest, assert } from '../../../tests/ts/test-runner';

async function main() {
    console.log("--- Starting VBA Unit Tests ---");
    const vbaFile = 'sample/src/vba/TaskScheduler_Core.vba';

    // Test #1: CalcBaseStartIdx
    console.log("\n[Test Suite] CalcBaseStartIdx");

    // Level 1: Always starts at 1, regardless of parent finish
    assert.strictEqual(runVBATest(vbaFile, 'CalcBaseStartIdx', [1, 10, 1.0]), 1, "Level 1 ignores parent finish");

    // Level 2+ with parent alloc < 0.5: Starts on SAME day
    assert.strictEqual(runVBATest(vbaFile, 'CalcBaseStartIdx', [2, 10, 0.49]), 10, "Alloc < 0.5 starts same day");
    assert.strictEqual(runVBATest(vbaFile, 'CalcBaseStartIdx', [3, 5, 0.0]), 5, "Alloc 0 starts same day");

    // Level 2+ with parent alloc >= 0.5: Starts NEXT day
    assert.strictEqual(runVBATest(vbaFile, 'CalcBaseStartIdx', [2, 10, 0.5]), 11, "Alloc >= 0.5 starts next day");
    assert.strictEqual(runVBATest(vbaFile, 'CalcBaseStartIdx', [4, 20, 1.0]), 21, "Alloc 1.0 starts next day");

    // Test #2: GetMaxDailyLoad
    console.log("\n[Test Suite] GetMaxDailyLoad");

    // Setup simulated VBA Dictionary using JS Map proxy
    const mockDict = new Map<string, number>();
    mockDict.set("Alice", 0.5);
    mockDict.set("Bob", 0.8);

    // Add VBA Dictionary methods proxy for JS object
    // In JS, we cannot easily proxy `obj("Alice")` for a Map because it's interpreted as a function call.
    // Instead of complex proxying, we just attach a bare function.
    // VBA evaluator treats `capacityLimits("Alice")` as a CallExpression, so if capacityLimits is a function, it just calls it.
    const capacityLimits: any = function (key: string) {
        return mockDict.get(key);
    };
    capacityLimits.Exists = function (key: string) {
        return mockDict.has(key);
    };

    // Default when assignee is empty/not explicitly handled properly by dictionary vs missing
    assert.strictEqual(runVBATest(vbaFile, 'GetMaxDailyLoad', ["Charlie", capacityLimits]), 1.0, "Missing user defaults to 1.0");
    assert.strictEqual(runVBATest(vbaFile, 'GetMaxDailyLoad', ["Alice", capacityLimits]), 0.5, "Known user returns configured capacity");
    assert.strictEqual(runVBATest(vbaFile, 'GetMaxDailyLoad', ["Bob", capacityLimits]), 0.8, "Known user returns configured capacity");

    // Test #3: CalcDailyAllocation
    console.log("\n[Test Suite] CalcDailyAllocation");

    // Standard Task logic
    assert.strictEqual(runVBATest(vbaFile, 'CalcDailyAllocation', [1.0, 0.5, false]), 0.5, "Standard: exact fit");
    assert.strictEqual(runVBATest(vbaFile, 'CalcDailyAllocation', [0.25, 0.5, false]), 0.25, "Standard: limited by capacity");
    assert.strictEqual(runVBATest(vbaFile, 'CalcDailyAllocation', [1.0, 0.1, false]), 0, "Standard: rounds to 0 if < 0.125 needed");

    // Micro-Task logic
    assert.strictEqual(runVBATest(vbaFile, 'CalcDailyAllocation', [1.0, 0.1, true]), 0.1, "Micro: capacity 1.0 > 0.1, gets 0.1");
    assert.strictEqual(runVBATest(vbaFile, 'CalcDailyAllocation', [0.05, 0.1, true]), 0, "Micro: capacity < 0.1, skips");

    // Test #4: UpdateLevelFinish (ByRef array mutation)
    console.log("\n[Test Suite] UpdateLevelFinish");

    // Arrays representing levelMaxFinish and levelMaxFinishAlloc (1-indexed for VBA, so we use index 1,2)
    const levelFinish = [0, 10, 5];
    const levelAlloc = [0.0, 0.5, 1.0];

    // Subroutine returns EmptyVBA, so we strictly check array mutations
    runVBATest(vbaFile, 'UpdateLevelFinish', [1, 15, 0.8, levelFinish, levelAlloc]);
    assert.strictEqual(levelFinish[1], 15, "Level 1 finish index updated due to > 10");
    assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc updated due to index change");

    // Same day index (15), but smaller alloc (0.1) -> No change
    runVBATest(vbaFile, 'UpdateLevelFinish', [1, 15, 0.1, levelFinish, levelAlloc]);
    assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc unchanged, 0.1 not > 0.8");

    // Same day index (15), but larger alloc (1.0) -> Alloc updates
    runVBATest(vbaFile, 'UpdateLevelFinish', [1, 15, 1.0, levelFinish, levelAlloc]);
    assert.strictEqual(levelAlloc[1], 1.0, "Level 1 alloc updated, 1.0 > 0.8");

    console.log("\n[Test Suite] BuildCapacityDict");

    // Simulate 1-based 2D array of config data for arr(row, col):
    // Evaluator mock reads `current[row][col]`.
    // row 1: [null, "Alice", 0.5] -> arr[1][1] = "Alice", arr[1][2] = 0.5
    const mockConfigData = [
        null, // row 0
        [null, "Alice", 0.5],   // row 1
        [null, "Bob", null],    // row 2
        [null, "   ", 2.0],     // row 3
        [null, "Dan", 1.25]     // row 4
    ];

    // BuildCapacityDict returns a VBA Dictionary. We can call .Exists and .Item() via evaluator if we test it properly.
    // Instead of complex dict inspection, let's wrap it in another evaluation or test the keys specifically, or we can just 
    // run it and see if the stub dictionary mechanism throws.
    // However, our `runVBATest` returns the evaluated result. For objects, it returns the JS representation.
    const resultDict = runVBATest(vbaFile, 'BuildCapacityDict', [mockConfigData]);

    // Check if the Map/Dictionary populated correctly
    assert.strictEqual(resultDict.__map__.get("Alice"), 0.5, "Alice config matches");
    assert.strictEqual(resultDict.__map__.get("Bob"), 1.0, "Bob config defaults to 1.0");
    assert.strictEqual(resultDict.__map__.get("Dan"), 1.25, "Dan config matches");
    assert.strictEqual(resultDict.exists("   ") || resultDict.exists(""), false, "Empty name skipped");
    assert.strictEqual(resultDict.__map__.size, 3, "Only Alice, Bob, and Dan were added");

    console.log("\n[Test Suite] ScanLockedRows");

    // Mock dictionaries/objects are initialized by evaluator natively if we let it
    // but the test runner runs strictly pure function. So we'll run a mini script inside runVBATest
    // However, runVBATest just calls the target function.

    // Simulate metaData: [null, [null, col1... col13=Lock, col14, col15, col16, col17=Name] ]
    const mockMetaData = [
        null,
        // Row 1: Unlocked, Alice
        Array(20).fill("").map((_, i) => i === 13 ? "" : (i === 17 ? "Alice" : "")),
        // Row 2: Locked, Bob
        Array(20).fill("").map((_, i) => i === 13 ? "L" : (i === 17 ? "Bob" : "")),
        // Row 3: Locked, Alice
        Array(20).fill("").map((_, i) => i === 13 ? "L" : (i === 17 ? "Alice" : ""))
    ];

    // Simulate gridData: 1-based [null, [null, Day1, Day2, Day3], ... ]
    const mockGridData = [
        null,
        [null, 1.0, 0.0, 0.0], // Row 1 (Unlocked) - should be ignored by ScanLockedRows
        [null, 0.0, 0.5, 0.5], // Row 2 (Bob, Locked)
        [null, 0.5, 0.0, 0.25] // Row 3 (Alice, Locked)
    ];

    // The signature: ScanLockedRows(numRows, numDays, metaData, gridData, ByRef personUsage)
    // Create an empty personUsage dictionary map
    const personUsage = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
        add: function (k: string, v: any) { this.__map__.set(k, v); },
        exists: function (k: string) { return this.__map__.has(k); }
    };

    // runVBATest wrapper doesn't extract by-ref cleanly, but the argument `personUsage` is mutated 
    runVBATest(vbaFile, 'ScanLockedRows', [3, 3, mockMetaData, mockGridData, personUsage]);

    // Bob has 0.5 on Day 2, 0.5 on Day 3
    const bobUsage = personUsage.__map__.get("Bob");
    assert.strictEqual(bobUsage[2], 0.5, "Bob day 2 usage pre-allocated");
    assert.strictEqual(bobUsage[3], 0.5, "Bob day 3 usage pre-allocated");

    // Alice is Row 1 (Unlocked) padding to 0, Row 3 (Locked) adding 0.5 and 0.25
    const aliceUsage = personUsage.__map__.get("Alice");
    assert.strictEqual(aliceUsage[1], 0.5, "Alice day 1 usage pre-allocated");
    assert.strictEqual(aliceUsage[2], 0.0, "Alice day 2 remains 0");
    assert.strictEqual(aliceUsage[3], 0.25, "Alice day 3 usage pre-allocated");

    console.log("\n[Test Suite] ClearTaskGridRow");

    // Simulate gridData with some mock usage values we want to clear
    const gridToClear = [
        null,
        [null, 0.5, 0.5, 0.5], // Row 1
        [null, 1.0, 1.0, 1.0], // Row 2
        [null, 0.25, 0.0, 0.5] // Row 3
    ];

    // Signature: ClearTaskGridRow(taskRow, numDays, ByRef gridData)
    // Clear row 2
    runVBATest(vbaFile, 'ClearTaskGridRow', [2, 3, gridToClear]);

    // Row 2 should be empty (null is the Evaluator's representation of VBA Empty)
    assert.strictEqual(gridToClear[2][1], null, "Row 2 Day 1 cleared");
    assert.strictEqual(gridToClear[2][2], null, "Row 2 Day 2 cleared");
    assert.strictEqual(gridToClear[2][3], null, "Row 2 Day 3 cleared");

    // Row 1 should remain unaffected
    assert.strictEqual(gridToClear[1][1], 0.5, "Row 1 unaffected");

    console.log("\n[Test Suite] FindLockedTaskFinish");
    const mockLockedGrid = [
        null, // 0-based offset
        [null, 0.5, 1.0, 0.0, 0.0], // Row 1 (finishes on day 2 with 1.0)
        [null, 0.0, 0.0, 0.0, 0.0]  // Row 2 (empty)
    ];

    // Signature: FindLockedTaskFinish(taskRow, numDays, gridData, ByRef finishIdx, ByRef finishAlloc)
    // We use a VBA wrapper Function `TestFindLockedTaskFinish` to return the ByRef primitives as a string.

    // Test Row 1
    const resRow1 = runVBATest(vbaFile, 'TestFindLockedTaskFinish', [1, 4, mockLockedGrid]);
    assert.strictEqual(resRow1, "2|1", "Row 1 Finish Idx 2, Alloc 1.0");

    // Test Row 2
    const resRow2 = runVBATest(vbaFile, 'TestFindLockedTaskFinish', [2, 4, mockLockedGrid]);
    assert.strictEqual(resRow2, "0|0", "Row 2 empty Finish Idx 0, Alloc 0");

    console.log("\n[Test Suite] ScheduleUnlockedTask");

    // mockMetaData: 1-based [null, [null, col1... col15=Duration, col16, col17=Name] ]
    // COL_OFFSET_IDX = 9, COL_DURATION_IDX = 15, COL_ASSIGNEE_IDX = 17
    const mockMetaSchedule = [
        null,
        // Row 1: Alice, duration = 1.0, lag = 1 (col 9)
        Array(20).fill("").map((_, i) => i === 9 ? 1 : (i === 15 ? 1.0 : (i === 17 ? "Alice" : ""))),
        // Row 2: Bob, duration = 0.5, lag = "" 
        Array(20).fill("").map((_, i) => i === 9 ? "" : (i === 15 ? 0.5 : (i === 17 ? "Bob" : ""))),
        // Row 3: Charlie, duration = 0.1 (microtask true), lag = ""
        Array(20).fill("").map((_, i) => i === 9 ? "" : (i === 15 ? 0.1 : (i === 17 ? "Charlie" : "")))
    ];

    const mockHolidayData = [
        null,
        [null, "", "休", "", "", "休"] // Day 2 and Day 5 are holidays
    ];

    const mockGridSchedule: any[] = [
        null,
        [null, 0, 0, 0, 0, 0], // Row 1
        [null, 0, 0, 0, 0, 0], // Row 2
        [null, 0, 0, 0, 0, 0]  // Row 3
    ];

    const mockPersonUsage = {
        __isVbaDict__: true,
        __map__: new Map<string, any>(),
        add: function (k: string, v: any) { this.__map__.set(k, v); },
        exists: function (k: string) { return this.__map__.has(k); }
    };

    // Row 1: Alice, capacity 0.5 (from mocked limits), duration 1.0. BaseStart 1. Lag 1 => start 2.
    // Day 2 is holiday => skip. 
    // Day 3: allocate 0.5 limit. remaining 0.5
    // Day 4: allocate 0.5 limit. remaining 0. FinishIdx = 4, Alloc = 0.5.
    const resRow1Schedule = runVBATest(vbaFile, 'TestScheduleUnlockedTask', [1, 5, 1, mockMetaSchedule, mockHolidayData, capacityLimits, mockGridSchedule, mockPersonUsage]);
    assert.strictEqual(resRow1Schedule, "4|0.5", "Row 1 Finish Idx 4, Alloc 0.5");
    assert.strictEqual(mockGridSchedule[1][2], null, "Alice Day 2 is holiday (cleared to null)");
    assert.strictEqual(mockGridSchedule[1][3], 0.5, "Alice Day 3 gets 0.5 (capacity limit)");
    assert.strictEqual(mockGridSchedule[1][4], 0.5, "Alice Day 4 gets 0.5");

    // Row 2: Bob, capacity 0.8. duration 0.5. BaseStart 3. Lag "". => start 3.
    // Day 3: allocate 0.5. remaining 0. FinishIdx = 3, Alloc = 0.5.
    const resRow2Schedule = runVBATest(vbaFile, 'TestScheduleUnlockedTask', [2, 5, 3, mockMetaSchedule, mockHolidayData, capacityLimits, mockGridSchedule, mockPersonUsage]);
    assert.strictEqual(resRow2Schedule, "3|0.5", "Row 2 Finish Idx 3, Alloc 0.5");
    assert.strictEqual(mockGridSchedule[2][3], 0.5, "Bob Day 3 gets 0.5");

    console.log("\n--- All tests passed! ---");
}

main().catch(console.error);
