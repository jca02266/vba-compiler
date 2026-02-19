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
    StringLiteral,
    AssignmentStatement,
    ProcedureDeclaration,
    VariableDeclaration,
    CallStatement,
    CallExpression,
    MemberExpression
} from './parser';

export class Environment {
    private variables: Map<string, any> = new Map();
    private procedures: Map<string, ProcedureDeclaration> = new Map();
    public enclosing?: Environment;

    constructor(enclosing?: Environment) {
        this.enclosing = enclosing;
    }

    set(name: string, value: any) {
        // Find scope where variable is defined to update it, or set locally
        const key = name.toLowerCase();
        let env: Environment | undefined = this;
        while (env) {
            if (env.variables.has(key)) {
                env.variables.set(key, value);
                return;
            }
            env = env.enclosing;
        }
        // If not found, implicitly declare locally
        this.variables.set(key, value);
    }

    get(name: string): any {
        const key = name.toLowerCase();
        let env: Environment | undefined = this;
        while (env) {
            if (env.variables.has(key)) {
                return env.variables.get(key);
            }
            env = env.enclosing;
        }

        // Implicit initialization
        this.variables.set(key, 0);
        return 0;
    }

    setProcedure(name: string, proc: ProcedureDeclaration) {
        this.procedures.set(name.toLowerCase(), proc);
    }

    getProcedure(name: string): ProcedureDeclaration | undefined {
        const key = name.toLowerCase();
        let env: Environment | undefined = this;
        while (env) {
            if (env.procedures.has(key)) {
                return env.procedures.get(key);
            }
            env = env.enclosing;
        }
        return undefined;
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
                this.evaluateAssignmentStatement(stmt as AssignmentStatement);
                break;
            case 'ProcedureDeclaration':
                this.env.setProcedure((stmt as ProcedureDeclaration).name.name, stmt as ProcedureDeclaration);
                break;
            case 'VariableDeclaration':
                this.evaluateVariableDeclaration(stmt as VariableDeclaration);
                break;
            case 'CallStatement':
                this.evaluateCallStatement(stmt as CallStatement);
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

    private evaluateAssignmentStatement(stmt: AssignmentStatement) {
        const val = this.evaluateExpression(stmt.right);

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            this.env.set(name, val);
        } else if (stmt.left.type === 'CallExpression') {
            // Array assignment: arr(0) = val
            const call = stmt.left as CallExpression;
            if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const arr = this.env.get(name);
                if (Array.isArray(arr)) {
                    const idx = this.evaluateExpression(call.args[0]);
                    arr[idx] = val;
                } else {
                    throw new Error(`Execution error: ${name} is not an array`);
                }
            } else {
                throw new Error("Execution error: Complex left hand assignments not supported yet");
            }
        }
    }

    private evaluateVariableDeclaration(stmt: VariableDeclaration) {
        if (stmt.isArray && stmt.arraySize) {
            const size = this.evaluateExpression(stmt.arraySize);
            const arr = new Array(size + 1).fill(0); // VBA arrays typically 0 to N
            this.env.set(stmt.name.name, arr);
        } else if (stmt.objectType?.toLowerCase() === 'collection' || stmt.isNew) {
            // Currently represent Collection as an array with some wrapper behavior
            const coll: any[] = [];
            // We attach methods to the object to simulate VBA Collection methods
            (coll as any).add = (item: any) => coll.push(item);
            (coll as any).item = (idx: number) => coll[idx - 1]; // VBA collections are 1-indexed
            (coll as any).count = () => coll.length;
            this.env.set(stmt.name.name, coll);
        } else {
            this.env.set(stmt.name.name, 0);
        }
    }

    private evaluateCallStatement(stmt: CallStatement) {
        this.evaluateExpression(stmt.expression);
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
            case 'CallExpression':
                return this.evaluateCallExpression(expr as CallExpression);
            case 'MemberExpression':
                return this.evaluateMemberExpression(expr as MemberExpression);
            case 'BinaryExpression':
                return this.evaluateBinaryExpression(expr as BinaryExpression);
            default:
                throw new Error(`Execution error: Unknown expression type ${expr.type}`);
        }
    }

    private evaluateCallExpression(expr: CallExpression): any {
        if (expr.callee.type === 'Identifier') {
            const name = (expr.callee as Identifier).name;
            const proc = this.env.getProcedure(name);

            if (proc) {
                // Procedure call (Function/Sub)
                const localEnv = new Environment(this.env);

                // Map arguments to parameters
                for (let i = 0; i < proc.parameters.length; i++) {
                    const argVal = i < expr.args.length ? this.evaluateExpression(expr.args[i]) : 0;
                    localEnv.set(proc.parameters[i].name, argVal);
                }

                if (proc.isFunction) {
                    localEnv.set(proc.name.name, 0); // initial return value
                }

                const prevEnv = this.env;
                this.env = localEnv;

                try {
                    for (const s of proc.body) {
                        this.evaluateStatement(s);
                    }
                } finally {
                    this.env = prevEnv;
                }

                if (proc.isFunction) {
                    return localEnv.get(proc.name.name);
                }
                return undefined;
            } else {
                // Might be an array access
                const variable = this.env.get(name);
                if (Array.isArray(variable)) {
                    if (expr.args.length === 0) throw new Error(`Execution error: Missing index for array ${name}`);
                    const idx = this.evaluateExpression(expr.args[0]);
                    return variable[idx]; // 0-indexed in JS/VBA unless Option Base 1
                }
                throw new Error(`Execution error: Cannot call unknown procedure or index unknown array '${name}'`);
            }
        } else if (expr.callee.type === 'MemberExpression') {
            // Example: col.Add("Cherry")
            const member = expr.callee as MemberExpression;
            const obj = this.evaluateExpression(member.object);
            const methodName = member.property.name.toLowerCase();

            if (obj && typeof obj[methodName] === 'function') {
                const argsVals = expr.args.map(a => this.evaluateExpression(a));
                return obj[methodName](...argsVals);
            } else {
                throw new Error(`Execution error: Object does not support property or method '${methodName}'`);
            }
        }

        throw new Error(`Execution error: Unsupported call expression`);
    }

    private evaluateMemberExpression(expr: MemberExpression): any {
        const obj = this.evaluateExpression(expr.object);
        const propName = expr.property.name.toLowerCase();

        // Handling function calls without parens like `col.Count` -> `col.count()` 
        // This is a simplification for VBA collections.
        if (obj && typeof obj[propName] === 'function') {
            return obj[propName]();
        } else if (obj && propName in obj) {
            return obj[propName];
        }
        throw new Error(`Execution error: Method or property not found '${propName}'`);
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
