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

    console.log("\n--- All tests passed! ---");
}

main().catch(console.error);
