const fs = require('fs');
const { Lexer } = require('../../../src/compiler/lexer.js');
const { Parser } = require('../../../src/compiler/parser.js');
const { Evaluator } = require('../../../src/compiler/evaluator.js');

const code = `
Function BuildCapacityDict(configData)
    Dim capacityLimits As Object
    Set capacityLimits = CreateObject("Scripting.Dictionary")
    
    Dim cfgRow As Long
    Dim cfgCapacity As Double
    Dim cfgName As String
    
    For cfgRow = 1 To UBound(configData, 1)
        cfgName = Trim(configData(cfgRow, 1))
        If cfgName <> "" Then
            cfgCapacity = 1
            If IsNumeric(configData(cfgRow, 2)) And Not IsEmpty(configData(cfgRow, 2)) Then
                cfgCapacity = CDbl(configData(cfgRow, 2))
            End If
            capacityLimits(cfgName) = cfgCapacity
            debug.print "Added " & cfgName & " = " & cfgCapacity
        End If
    Next cfgRow
    
    Set BuildCapacityDict = capacityLimits
End Function

Dim dict As Object
Dim arr(4)
arr(1) = "dummy" ' Not 2D, but we mock it
Set dict = BuildCapacityDict(arr)
`;

const lexer = new Lexer(code);
const tokens = lexer.tokenize();
const parser = new Parser(tokens);
const ast = parser.parse();

const evaluator = new Evaluator((o) => console.log(o));
const mockConfigData = [
    null,
    [null, "Alice", 0.5],
    [null, "Bob", null],
    [null, "   ", 2.0],
    [null, "Dan", 1.25]
];

evaluator.env.setLocally("arr", mockConfigData);
evaluator.evaluate(ast);
const result = evaluator.get("dict");
console.log(result);
