import {
    Program,
    Statement,
    ForStatement,
    IfStatement,
    DoWhileStatement,
    DebugPrintStatement,
    Expression,
    BinaryExpression,
    Identifier,
    NumberLiteral,
    StringLiteral
} from './parser';

export class Environment {
    private variables: Map<string, any> = new Map();

    // Future use for Option Explicit
    // private explicitlyDeclared: Set<string> = new Set();

    set(name: string, value: any) {
        this.variables.set(name.toLowerCase(), value);
    }

    get(name: string): any {
        const key = name.toLowerCase();
        if (!this.variables.has(key)) {
            // If Option Explicit was strict, we would throw here.
            // But user requested to just warn later, not error for now.
            // So we implicitly initialize undefined variables to 0.
            this.variables.set(key, 0);
            return 0;
        }
        return this.variables.get(key);
    }
}

export type PrintCallback = (output: string) => void;

export class Evaluator {
    private env: Environment;
    private onPrint: PrintCallback;

    constructor(onPrint: PrintCallback) {
        this.env = new Environment();
        this.onPrint = onPrint;
    }

    public evaluate(program: Program) {
        for (const stmt of program.body) {
            this.evaluateStatement(stmt);
        }
    }

    private evaluateStatement(stmt: Statement) {
        switch (stmt.type) {
            case 'ForStatement':
                this.evaluateForStatement(stmt as ForStatement);
                break;
            case 'IfStatement':
                this.evaluateIfStatement(stmt as IfStatement);
                break;
            case 'DoWhileStatement':
                this.evaluateDoWhileStatement(stmt as DoWhileStatement);
                break;
            case 'AssignmentStatement':
                this.evaluateAssignmentStatement(stmt as any);
                break;
            case 'DebugPrintStatement':
                this.evaluateDebugPrintStatement(stmt as DebugPrintStatement);
                break;
            default:
                throw new Error(`Execution error: Unknown statement type ${stmt.type}`);
        }
    }

    private evaluateForStatement(stmt: ForStatement) {
        const startValue = this.evaluateExpression(stmt.start);
        const endValue = this.evaluateExpression(stmt.end);
        const varName = stmt.identifier.name;

        for (let i = startValue; i <= endValue; i++) {
            this.env.set(varName, i);
            for (const bodyStmt of stmt.body) {
                this.evaluateStatement(bodyStmt);
            }
        }
    }

    private evaluateIfStatement(stmt: IfStatement) {
        const conditionVal = this.evaluateExpression(stmt.condition);
        if (conditionVal) {
            for (const bodyStmt of stmt.consequent) {
                this.evaluateStatement(bodyStmt);
            }
        } else if (stmt.alternate) {
            if (Array.isArray(stmt.alternate)) {
                for (const bodyStmt of stmt.alternate) {
                    this.evaluateStatement(bodyStmt);
                }
            } else {
                this.evaluateIfStatement(stmt.alternate as IfStatement);
            }
        }
    }

    private evaluateDoWhileStatement(stmt: DoWhileStatement) {
        while (this.evaluateExpression(stmt.condition)) {
            for (const bodyStmt of stmt.body) {
                this.evaluateStatement(bodyStmt);
            }
        }
    }

    private evaluateAssignmentStatement(stmt: any) {
        const val = this.evaluateExpression(stmt.value);
        this.env.set(stmt.identifier.name, val);
    }

    private evaluateDebugPrintStatement(stmt: DebugPrintStatement) {
        const val = this.evaluateExpression(stmt.expression);
        this.onPrint(String(val));
    }

    private evaluateExpression(expr: Expression): any {
        switch (expr.type) {
            case 'NumberLiteral':
                return (expr as NumberLiteral).value;
            case 'StringLiteral':
                return (expr as StringLiteral).value;
            case 'Identifier':
                return this.env.get((expr as Identifier).name);
            case 'BinaryExpression':
                return this.evaluateBinaryExpression(expr as BinaryExpression);
            default:
                throw new Error(`Execution error: Unknown expression type ${expr.type}`);
        }
    }

    private evaluateBinaryExpression(expr: BinaryExpression): any {
        const left = this.evaluateExpression(expr.left);
        const right = this.evaluateExpression(expr.right);

        switch (expr.operator.toLowerCase()) {
            case '+': return left + right;
            case '-': return left - right;
            case '=': return left === right;
            case '<>': return left !== right;
            case '<': return left < right;
            case '>': return left > right;
            case '<=': return left <= right;
            case '>=': return left >= right;
            case 'and': return left && right;
            case 'or': return left || right;
            default:
                throw new Error(`Execution error: Unknown operator ${expr.operator}`);
        }
    }
}
