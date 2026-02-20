import {
    Program,
    Statement,
    ForStatement,
    IfStatement,
    DoWhileStatement,
    Expression,
    UnaryExpression,
    BinaryExpression,
    Identifier,
    NumberLiteral,
    StringLiteral,
    AssignmentStatement,
    ProcedureDeclaration,
    VariableDeclaration,
    CallStatement,
    CallExpression,
    MemberExpression,
    ConstDeclaration,
    SetStatement,
    EraseStatement,
    ReDimStatement,
    ExitStatement,
    LabelStatement
} from './parser';

export const EmptyVBA = null;

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

    setLocally(name: string, value: any) {
        this.variables.set(name.toLowerCase(), value);
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
        // Add built-in debug object
        this.env.set('debug', {
            print: (...args: any[]) => this.onPrint(args.join(' '))
        });

        // Add typical VBA built-ins
        this.env.set('isempty', (val: any) => val === undefined || val === null || val === '');
        this.env.set('isnumeric', (val: any) => !isNaN(parseFloat(val)) && isFinite(val));
        this.env.set('cdbl', (val: any) => parseFloat(val) || 0);
        this.env.set('clng', (val: any) => Math.round(parseFloat(val)) || 0);
        this.env.set('int', (val: any) => Math.floor(parseFloat(val)) || 0);
        this.env.set('ucase', (val: any) => String(val || '').toUpperCase());
        this.env.set('trim', (val: any) => String(val || '').trim());
        this.env.set('ubound', (arr: any[], dimension?: number) => {
            if (Array.isArray(arr)) {
                if (dimension === 2 && arr.length > 0 && Array.isArray(arr[0])) {
                    return arr[0].length - 1;
                } else if (dimension === 2 && arr.length > 1 && Array.isArray(arr[1])) {
                    return arr[1].length - 1; // Fallback if arr[0] is null/empty in mocks
                }
                return arr.length - 1; // Assuming 0-indexed in JS (or 1-indexed filled with nulls)
            }
            return 0;
        });
        this.env.set('createobject', (progId: string) => {
            if (progId.toLowerCase() === 'scripting.dictionary') {
                const dict = new Map<string, any>();
                // We return an object that can act as a dictionary for method calls
                // but ALSO retains the inner Map for our `CallExpression` hacks
                const vbaDict: any = {
                    __isVbaDict__: true,
                    __map__: dict,
                    add: (k: string, v: any) => dict.set(k, v),
                    exists: (k: string) => dict.has(k),
                    items: () => Array.from(dict.values()),
                    keys: () => Array.from(dict.keys())
                };
                return vbaDict;
            }
            throw new Error(`Execution error: Unsupported CreateObject '${progId}'`);
        });

        // Add VBA intrinsic constants
        this.env.set('true', true);
        this.env.set('false', false);
        this.env.set('empty', EmptyVBA);
    }

    public get(name: string): any {
        return this.env.get(name);
    }

    public callProcedure(name: string, args: any[]): any {
        const procName = name.toLowerCase();
        const proc = this.env.getProcedure(procName);

        if (!proc) {
            throw new Error(`Execution error: Procedure '${name}' not found`);
        }

        // Create a new local environment for the procedure call
        const localEnv = new Environment(this.env);

        // Map arguments to parameter names
        for (let i = 0; i < proc.parameters.length; i++) {
            const paramName = proc.parameters[i].name;
            const argValue = i < args.length ? args[i] : EmptyVBA;
            localEnv.setLocally(paramName, argValue);
        }

        // Save current env and swap to local
        const previousEnv = this.env;
        this.env = localEnv;

        try {
            // Execute procedure body
            for (const stmt of proc.body) {
                this.evaluateStatement(stmt);
            }
        } catch (e: any) {
            if (e && e.type === 'Exit') {
                if ((e.target === 'Function' && proc.isFunction) || (e.target === 'Sub' && !proc.isFunction)) {
                    // Valid exit, caught and swallowed
                } else {
                    throw e; // Unhandled exit type
                }
            } else {
                throw e; // Real error
            }
        } finally {
            // Restore previous environment
            this.env = previousEnv;
        }

        // Return the function value if it was a function
        if (proc.isFunction) {
            return localEnv.get(procName);
        }
        return EmptyVBA;
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
            case 'ConstDeclaration':
                this.evaluateConstDeclaration(stmt as ConstDeclaration);
                break;
            case 'SetStatement':
                this.evaluateSetStatement(stmt as SetStatement);
                break;
            case 'OnErrorStatement':
                // Ignore for now (No-op)
                break;
            case 'EraseStatement':
                this.evaluateEraseStatement(stmt as EraseStatement);
                break;
            case 'ReDimStatement':
                this.evaluateReDimStatement(stmt as ReDimStatement);
                break;
            case 'ExitStatement':
                this.evaluateExitStatement(stmt as ExitStatement);
                break;
            case 'LabelStatement':
                // No-op for now. Label execution just passes through.
                break;
            default:
                throw new Error(`Execution error: Unknown statement type ${stmt.type}`);
        }
    }

    private evaluateForStatement(stmt: ForStatement) {
        let startValue = this.evaluateExpression(stmt.start);
        const endValue = this.evaluateExpression(stmt.end);
        let stepValue = stmt.step ? this.evaluateExpression(stmt.step) : 1;
        const varName = stmt.identifier.name;

        // Initialize block scope variable if it doesn't exist
        if (this.env.get(varName) === EmptyVBA) { // Check against EmptyVBA
            this.env.set(varName, startValue);
        } else {
            this.env.setLocally(varName, startValue);
        }

        const condition = () => stepValue > 0 ? this.env.get(varName) <= endValue : this.env.get(varName) >= endValue;

        while (condition()) {
            try {
                for (const bodyStmt of stmt.body) {
                    this.evaluateStatement(bodyStmt);
                }
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'For') {
                    break;
                }
                throw e; // re-throw if it wasn't an Exit For
            }
            // Increment/decrement loop variable
            this.env.setLocally(varName, this.env.get(varName) + stepValue);
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
            try {
                for (const bodyStmt of stmt.body) {
                    this.evaluateStatement(bodyStmt);
                }
            } catch (e: any) {
                if (e && e.type === 'Exit' && e.target === 'Do') {
                    break;
                }
                throw e;
            }
        }
    }

    private evaluateAssignmentStatement(stmt: AssignmentStatement) {
        const val = this.evaluateExpression(stmt.right);

        if (stmt.left.type === 'Identifier') {
            const name = (stmt.left as Identifier).name;
            this.env.set(name, val);
        } else if (stmt.left.type === 'CallExpression') {
            // Array/Dictionary assignment: arr(0) = val OR dict("key") = val
            const call = stmt.left as CallExpression;
            if (call.callee.type === 'Identifier') {
                const name = (call.callee as Identifier).name;
                const target = this.env.get(name);
                const idx = this.evaluateExpression(call.args[0]);

                if (Array.isArray(target)) {
                    // Support 1D or multi-dimensional array assignment arr(0, 1) = val -> arr[0][1] = val
                    let current = target;
                    for (let i = 0; i < call.args.length - 1; i++) {
                        const d = this.evaluateExpression(call.args[i]);
                        if (!current[d]) current[d] = []; // Auto-instantiate inner arrays
                        current = current[d];
                    }
                    const lastIdx = this.evaluateExpression(call.args[call.args.length - 1]);
                    current[lastIdx] = val;
                } else if (target && target.__isVbaDict__) {
                    // Treat as Dictionary assignment dict("key") = val
                    target.__map__.set(idx, val);
                } else {
                    throw new Error(`Execution error: ${name} is not an array or dictionary`);
                }
            } else {
                throw new Error("Execution error: Complex left hand assignments not supported yet");
            }
        } else if (stmt.left.type === 'MemberExpression') {
            const member = stmt.left as MemberExpression;
            // E.g. dict(key) = val -> wait, Dictionary default property is dict.Item(key) = val
            // But let's handle object property assignment Application.ScreenUpdating = false
            const obj = this.evaluateExpression(member.object);
            const propName = member.property.name.toLowerCase();
            if (obj && typeof obj === 'object') {
                obj[propName] = val;
            } else {
                throw new Error(`Execution error: Cannot assign property '${propName}' of undefined or primitive`);
            }
        } else {
            throw new Error(`Execution error: Invalid assignment target`);
        }
    }

    private evaluateVariableDeclaration(stmt: VariableDeclaration) {
        for (const decl of stmt.declarations) {
            let initialValue: any = EmptyVBA;
            if (decl.isArray) {
                if (decl.arraySize) {
                    const size = this.evaluateExpression(decl.arraySize);
                    initialValue = new Array(size + 1).fill(EmptyVBA);
                } else {
                    initialValue = [];
                }
            } else if (decl.isNew && decl.objectType === 'Collection') {
                initialValue = {
                    items: [],
                    add: function (item: any) { this.items.push(item); },
                    count: function () { return this.items.length; },
                    item: function (index: number) { return this.items[index - 1]; }
                };
            }
            this.env.set(decl.name.name, initialValue);
        }
    }

    private evaluateCallStatement(stmt: CallStatement) {
        this.evaluateExpression(stmt.expression);
    }

    private evaluateConstDeclaration(stmt: ConstDeclaration) {
        const value = this.evaluateExpression(stmt.value);
        this.env.set(stmt.name.name, value);
    }

    private evaluateSetStatement(stmt: SetStatement) {
        let value = this.evaluateExpression(stmt.right);
        // If the right side evaluates to a variable name (string), resolve it
        if (typeof value === 'string' && stmt.right.type === 'Identifier') {
            value = this.env.get(value);
        }

        if (stmt.left.type === 'Identifier') {
            this.env.set((stmt.left as Identifier).name, value);
        } else {
            throw new Error(`Execution error: Unsupported Set target ${stmt.left.type}`);
        }
    }

    private evaluateEraseStatement(stmt: EraseStatement) {
        this.env.set(stmt.name.name, []);
    }

    private evaluateReDimStatement(stmt: ReDimStatement) {
        // Evaluate bounds (just size for 1D for now)
        if (stmt.bounds.length > 0) {
            const size = this.evaluateExpression(stmt.bounds[stmt.bounds.length - 1]); // naive: takes last bound as size
            const arr = new Array(size + 1).fill(0); // VBA numeric arrays default to 0
            this.env.set(stmt.name.name, arr);
        }
    }

    private evaluateExitStatement(stmt: ExitStatement) {
        throw { type: 'Exit', target: stmt.exitType };
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
            case 'UnaryExpression':
                return this.evaluateUnaryExpression(expr as UnaryExpression);
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
                const byRefArgs: { paramName: string, identifierName: string }[] = [];
                for (let i = 0; i < proc.parameters.length; i++) {
                    const argVal = i < expr.args.length ? this.evaluateExpression(expr.args[i]) : 0;
                    localEnv.set(proc.parameters[i].name, argVal);

                    if (i < expr.args.length && expr.args[i].type === 'Identifier') {
                        // VBA Default is ByRef. If it is NOT explicitly ByVal, it is ByRef
                        if (!proc.parameters[i].isByVal) {
                            byRefArgs.push({
                                paramName: proc.parameters[i].name,
                                identifierName: (expr.args[i] as Identifier).name
                            });
                        }
                    }
                }

                if (proc.isFunction) {
                    // Implicit variable for function return value
                    localEnv.setLocally(proc.name.name, EmptyVBA);
                }

                const previousEnv = this.env;
                this.env = localEnv;

                try {
                    for (const s of proc.body) {
                        this.evaluateStatement(s);
                    }
                } catch (e: any) {
                    if (e && e.type === 'Exit' && (e.target === 'Sub' || e.target === 'Function')) {
                        // Exit the procedure cleanly
                    } else {
                        throw e;
                    }
                }

                this.env = previousEnv; // Restore scope

                // Synchronize ByRef arguments back to caller scope
                for (const ref of byRefArgs) {
                    const updatedVal = localEnv.get(ref.paramName);
                    this.env.setLocally(ref.identifierName, updatedVal);
                }

                if (proc.isFunction) {
                    return localEnv.get(proc.name.name);
                }
                return undefined;
            } else {
                // Might be an array access or built-in function
                const variable = this.env.get(name);
                if (typeof variable === 'function') {
                    const argsVals = expr.args.map(a => this.evaluateExpression(a));
                    return variable(...argsVals);
                } else if (Array.isArray(variable)) {
                    if (expr.args.length === 0) throw new Error(`Execution error: Missing index for array ${name}`);
                    // Support multi-dimensional array lookup arr(0, 1) -> arr[0][1]
                    let current = variable;
                    for (let i = 0; i < expr.args.length; i++) {
                        if (!current) return EmptyVBA; // Out of bounds or jagged array
                        const idx = this.evaluateExpression(expr.args[i]);
                        current = current[idx];
                    }
                    if (current === undefined) return EmptyVBA;
                    return current;
                } else if (variable && variable.__isVbaDict__) {
                    // Dictionary read: dict("key")
                    if (expr.args.length === 0) throw new Error(`Execution error: Missing key for dictionary ${name}`);
                    const key = this.evaluateExpression(expr.args[0]);
                    return variable.__map__.get(key);
                }
                throw new Error(`Execution error: Cannot call unknown procedure or index unknown array '${name}'`);
            }
        } else if (expr.callee.type === 'MemberExpression') {
            // Example: col.Add("Cherry")
            const member = expr.callee as MemberExpression;
            const obj = this.evaluateExpression(member.object);
            const methodNameLower = member.property.name.toLowerCase();
            const methodNameOriginal = member.property.name;

            if (obj) {
                // Try case-insensitive lookup first, then fallback to original casing
                let targetMethod = obj[methodNameLower];

                if (typeof targetMethod !== 'function') {
                    // Search object keys for case-insensitive match (for JS proxies/objects)
                    const keys = Object.keys(obj);
                    // If proxy, Object.keys might not work perfectly, so also check original
                    if (typeof obj[methodNameOriginal] === 'function') {
                        targetMethod = obj[methodNameOriginal];
                    } else {
                        const match = keys.find(k => k.toLowerCase() === methodNameLower);
                        if (match) targetMethod = obj[match];
                    }
                }

                if (typeof targetMethod === 'function') {
                    const argsVals = expr.args.map(a => this.evaluateExpression(a));
                    return targetMethod.apply(obj, argsVals);
                }
            }
            throw new Error(`Execution error: Object does not support property or method '${methodNameOriginal}'`);
        }

        throw new Error(`Execution error: Unsupported call expression`);
    }

    private evaluateUnaryExpression(expr: UnaryExpression): any {
        const argument = this.evaluateExpression(expr.argument);
        switch (expr.operator.toLowerCase()) {
            case 'not':
                return !argument;
            case '-':
                return -argument;
            case '+':
                return +argument;
            default:
                throw new Error(`Execution error: Unknown unary operator ${expr.operator}`);
        }
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
        const leftVal = this.evaluateExpression(expr.left);
        const rightVal = this.evaluateExpression(expr.right);

        switch (expr.operator.toLowerCase()) {
            case '+': return leftVal + rightVal;
            case '&': return String(leftVal) + String(rightVal);
            case '-': return leftVal - rightVal;
            case '*': return leftVal * rightVal;
            case '/': return leftVal / rightVal;
            case '\\': return Math.floor(leftVal / rightVal);
            case 'mod': return leftVal % rightVal;
            case '^': return Math.pow(leftVal, rightVal);
            case '=': return leftVal === rightVal;
            case '<>': return leftVal !== rightVal;
            case '<': return leftVal < rightVal;
            case '>': return leftVal > rightVal;
            case '<=': return leftVal <= rightVal;
            case '>=': return leftVal >= rightVal;
            case 'and': return leftVal && rightVal;
            case 'or': return leftVal || rightVal;
            default:
                throw new Error(`Execution error: Unknown operator ${expr.operator}`);
        }
    }
}
