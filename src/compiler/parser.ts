import { Token, TokenType } from './lexer';

export interface ASTNode {
    type: string;
}

export interface Program extends ASTNode {
    type: 'Program';
    body: Statement[];
}

export interface Statement extends ASTNode { }

export interface ForStatement extends Statement {
    type: 'ForStatement';
    identifier: Identifier;
    start: Expression;
    end: Expression;
    step?: Expression;
    body: Statement[];
    nextIdentifier?: Identifier;
}

export interface IfStatement extends Statement {
    type: 'IfStatement';
    condition: Expression;
    consequent: Statement[];
    alternate: Statement[] | IfStatement | null; // For Else and ElseIf
}

export interface DoWhileStatement extends Statement {
    type: 'DoWhileStatement';
    condition: Expression;
    body: Statement[];
}

export interface Parameter extends ASTNode {
    type: 'Parameter';
    name: string;
}

export interface ProcedureDeclaration extends Statement {
    type: 'ProcedureDeclaration';
    isFunction: boolean;
    name: Identifier;
    parameters: Parameter[];
    body: Statement[];
}

export interface VariableDeclarator {
    name: Identifier;
    isArray: boolean;
    arraySize?: Expression; // TODO: Array size should ideally support multiple dimensions
    isNew: boolean;
    objectType?: string;
}

export interface VariableDeclaration extends Statement {
    type: 'VariableDeclaration';
    declarations: VariableDeclarator[];
}

export interface ConstDeclaration extends Statement {
    type: 'ConstDeclaration';
    name: Identifier;
    value: Expression;
}

export interface SetStatement extends Statement {
    type: 'SetStatement';
    left: Expression;
    right: Expression;
}

export interface OnErrorStatement extends Statement {
    type: 'OnErrorStatement';
    label: string; // "Cleanup", "Resume Next", "0"
}

export interface EraseStatement extends Statement {
    type: 'EraseStatement';
    name: Identifier;
}

export interface ReDimStatement extends Statement {
    type: 'ReDimStatement';
    name: Identifier;
    bounds: Expression[]; // Multi-dimensional bounds (e.g. 1 To numDays)
}

export interface ExitStatement extends Statement {
    type: 'ExitStatement';
    exitType: 'For' | 'Do' | 'Sub' | 'Function';
}

export interface LabelStatement extends Statement {
    type: 'LabelStatement';
    label: string;
}

export interface AssignmentStatement extends Statement {
    type: 'AssignmentStatement';
    left: Expression; // Identifier, CallExpression (for arrays), MemberExpression
    right: Expression;
}

export interface CallStatement extends Statement {
    type: 'CallStatement';
    expression: CallExpression;
}

export interface Expression extends ASTNode { }

export interface CallExpression extends Expression {
    type: 'CallExpression';
    callee: Expression;
    args: Expression[];
}

export interface MemberExpression extends Expression {
    type: 'MemberExpression';
    object: Expression;
    property: Identifier;
}

export interface Identifier extends Expression {
    type: 'Identifier';
    name: string;
}

export interface NumberLiteral extends Expression {
    type: 'NumberLiteral';
    value: number;
}

export interface StringLiteral extends Expression {
    type: 'StringLiteral';
    value: string;
}

export interface BinaryExpression extends Expression {
    type: 'BinaryExpression';
    operator: string;
    left: Expression;
    right: Expression;
}

export interface UnaryExpression extends Expression {
    type: 'UnaryExpression';
    operator: string;
    argument: Expression;
}

export class Parser {
    private tokens: Token[];
    private pos: number = 0;

    constructor(tokens: Token[]) {
        this.tokens = tokens;
    }

    private peek(): Token {
        if (this.pos >= this.tokens.length) {
            return this.tokens[this.tokens.length - 1]; // EOF
        }
        return this.tokens[this.pos];
    }

    private advance(): Token {
        const token = this.peek();
        this.pos++;
        return token;
    }

    private match(expectedType: TokenType): boolean {
        if (this.peek().type === expectedType) {
            this.advance();
            return true;
        }
        return false;
    }

    private skipNewlines() {
        while (this.match(TokenType.Newline)) { }
    }

    public parse(): Program {
        const program: Program = {
            type: 'Program',
            body: []
        };

        this.skipNewlines();
        while (this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) {
                program.body.push(stmt);
            }
            this.skipNewlines();
        }

        return program;
    }

    private parseStatement(): Statement | null {
        const token = this.peek();

        if (token.type === TokenType.KeywordFor) {
            return this.parseForStatement();
        } else if (token.type === TokenType.KeywordIf) {
            return this.parseIfStatement();
        } else if (token.type === TokenType.KeywordDo) {
            return this.parseDoWhileStatement();
        } else if (token.type === TokenType.KeywordSub || token.type === TokenType.KeywordFunction) {
            return this.parseProcedureDeclaration();
        } else if (token.type === TokenType.KeywordDim) {
            return this.parseDimStatement();
        } else if (token.type === TokenType.KeywordConst) {
            return this.parseConstDeclaration();
        } else if (token.type === TokenType.KeywordSet) {
            return this.parseSetStatement();
        } else if (token.type === TokenType.KeywordOn) {
            return this.parseOnErrorStatement();
        } else if (token.type === TokenType.KeywordExit) {
            return this.parseExitStatement();
        } else if (token.type === TokenType.KeywordErase) {
            return this.parseEraseStatement();
        } else if (token.type === TokenType.KeywordReDim) {
            return this.parseReDimStatement();
        } else if (token.type === TokenType.KeywordOption) {
            this.advance(); // 'Option'
            if (this.match(TokenType.KeywordExplicit)) {
                // Ignore for now
            }
            return null;
        } else if (token.type === TokenType.Identifier) {
            // Check if it's a label "Identifier:"
            if (this.pos + 1 < this.tokens.length && this.tokens[this.pos + 1].type === TokenType.OperatorColon) {
                const labelName = token.value;
                this.advance(); // consume Identifier
                this.advance(); // consume ':'
                return { type: 'LabelStatement', label: labelName } as any;
            }

            // Unify assignment, array access, method call
            const expr = this.parsePrimary(); // will parse `foo`, `foo()`, `foo.bar`, `arr(0)` etc

            if (this.match(TokenType.OperatorEquals)) {
                return {
                    type: 'AssignmentStatement',
                    left: expr,
                    right: this.parseExpression()
                } as AssignmentStatement;
            } else {
                // If it's not an assignment, maybe it's a CallStatement with arguments separated by comma
                const args: Expression[] = [];
                // Check if there are args on the same line
                if (
                    this.peek().type !== TokenType.Newline &&
                    this.peek().type !== TokenType.EOF &&
                    this.peek().type !== TokenType.KeywordElse &&
                    this.peek().type !== TokenType.KeywordElseIf &&
                    this.peek().type !== TokenType.KeywordEnd &&
                    this.peek().type !== TokenType.KeywordNext &&
                    this.peek().type !== TokenType.KeywordLoop
                ) {
                    args.push(this.parseExpression());
                    while (this.match(TokenType.OperatorComma)) {
                        args.push(this.parseExpression());
                    }
                }

                if (args.length > 0) {
                    return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args } } as CallStatement;
                } else if (expr.type === 'CallExpression') {
                    // Call matched via parens e.g. `MainLoop()`
                    return { type: 'CallStatement', expression: expr } as CallStatement;
                } else {
                    // Call matched without parens e.g. `MainLoop`
                    return { type: 'CallStatement', expression: { type: 'CallExpression', callee: expr, args: [] } } as CallStatement;
                }
            }
        } else if (token.type === TokenType.Unknown) {
            throw new Error(`Parse error: Unknown token '${token.value}' at line ${token.line}`);
        } else {
            // Unknown or unexpected top-level token
            this.advance();
        }
        return null;
    }

    private parseProcedureDeclaration(): ProcedureDeclaration {
        const isFunction = this.peek().type === TokenType.KeywordFunction;
        this.advance(); // consume Sub or Function

        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after Sub/Function at line ${idToken.line}`);
        }
        const name: Identifier = { type: 'Identifier', name: idToken.value };
        const parameters: Parameter[] = [];

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                let paramNameToken = this.peek();
                if (paramNameToken.type === TokenType.KeywordByVal || paramNameToken.type === TokenType.KeywordByRef) {
                    this.advance(); // consume ByVal/ByRef
                    paramNameToken = this.peek();
                }

                let paramName = this.advance(); // consume name

                // Optional parameter type (e.g. 'As Long')
                if (this.match(TokenType.KeywordAs)) {
                    this.advance(); // consume Type name
                }

                parameters.push({ type: 'Parameter', name: paramName.value });

                while (this.match(TokenType.OperatorComma)) {
                    // Optional ByVal/ByRef
                    let nextParamToken = this.peek();
                    if (nextParamToken.type === TokenType.KeywordByVal || nextParamToken.type === TokenType.KeywordByRef) {
                        this.advance(); // consume ByVal/ByRef
                        nextParamToken = this.peek();
                    }
                    paramName = this.advance(); // consume name

                    if (this.match(TokenType.KeywordAs)) {
                        this.advance(); // consume Type name
                    }
                    parameters.push({ type: 'Parameter', name: paramName.value });
                }
            }
            if (!this.match(TokenType.OperatorRParen)) {
                throw new Error(`Parse error: Expected ')' at line ${this.peek().line}`);
            }
        }

        // Optional Function return type (e.g. 'As Long')
        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // consume Type name
        }

        this.skipNewlines();
        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordEnd && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // consume 'End'
            const expectedEndStr = isFunction ? 'Function' : 'Sub';
            const endToken = this.advance();
            if (endToken.value.toLowerCase() !== expectedEndStr.toLowerCase()) {
                throw new Error(`Parse error: Expected '${expectedEndStr}' after 'End' at line ${endToken.line}`);
            }
        }

        return { type: 'ProcedureDeclaration', isFunction, name, parameters, body };
    }

    private parseDimStatement(): VariableDeclaration {
        this.advance(); // 'Dim'
        const declarations: VariableDeclarator[] = [];

        while (true) {
            const idToken = this.advance();
            const name: Identifier = { type: 'Identifier', name: idToken.value };

            let isArray = false;
            let arraySize: Expression | undefined;
            let isNew = false;
            let objectType: string | undefined;

            if (this.match(TokenType.OperatorLParen)) {
                isArray = true;
                if (this.peek().type !== TokenType.OperatorRParen) {
                    arraySize = this.parseExpression();
                }
                this.match(TokenType.OperatorRParen);
            }

            if (this.match(TokenType.KeywordAs)) {
                if (this.match(TokenType.KeywordNew)) {
                    isNew = true;
                }
                const typeToken = this.peek();
                if (typeToken.type === TokenType.KeywordCollection || typeToken.type === TokenType.Identifier) {
                    objectType = this.advance().value;
                }
            }

            declarations.push({ name, isArray, arraySize, isNew, objectType });

            if (this.match(TokenType.OperatorComma)) {
                continue;
            } else {
                break;
            }
        }

        return { type: 'VariableDeclaration', declarations };
    }

    private parseConstDeclaration(): ConstDeclaration {
        this.advance(); // 'Const'
        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier) throw new Error(`Parse error: Expected identifier after Const at line ${idToken.line}`);
        const name = { type: 'Identifier', name: idToken.value } as Identifier;

        // Optional 'As Type'
        if (this.match(TokenType.KeywordAs)) {
            this.advance(); // Ignore Type for now
        }

        if (!this.match(TokenType.OperatorEquals)) throw new Error(`Parse error: Expected '=' in Const at line ${this.peek().line}`);
        const value = this.parseExpression();

        return { type: 'ConstDeclaration', name, value };
    }

    private parseSetStatement(): SetStatement {
        this.advance(); // 'Set'
        const left = this.parsePrimary(); // parse identifier or member access
        if (!this.match(TokenType.OperatorEquals)) throw new Error(`Parse error: Expected '=' in Set statement at line ${this.peek().line}`);
        const right = this.parseExpression();
        return { type: 'SetStatement', left, right };
    }

    private parseOnErrorStatement(): OnErrorStatement {
        this.advance(); // 'On'
        if (!this.match(TokenType.KeywordError)) throw new Error(`Parse error: Expected 'Error' after 'On' at line ${this.peek().line}`);

        let label = '';
        if (this.match(TokenType.KeywordGoTo)) {
            const labelToken = this.advance(); // Identifier or 0
            label = labelToken.value;
        } else {
            // "Resume Next" fallback
            while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
                label += this.advance().value + ' ';
            }
            label = label.trim();
        }
        return { type: 'OnErrorStatement', label };
    }

    private parseExitStatement(): ExitStatement {
        this.advance(); // 'Exit'
        const typeToken = this.advance();
        let exitType: 'For' | 'Do' | 'Sub' | 'Function';

        if (typeToken.type === TokenType.KeywordFor) {
            exitType = 'For';
        } else if (typeToken.type === TokenType.KeywordDo) {
            exitType = 'Do';
        } else if (typeToken.type === TokenType.KeywordSub) {
            exitType = 'Sub';
        } else if (typeToken.type === TokenType.KeywordFunction) {
            exitType = 'Function';
        } else {
            throw new Error(`Parse error: Unexpected token after Exit '${typeToken.value}' at line ${typeToken.line} `);
        }

        return { type: 'ExitStatement', exitType };
    }

    private parseEraseStatement(): EraseStatement {
        this.advance(); // 'Erase'
        const idToken = this.advance();
        const name = { type: 'Identifier', name: idToken.value } as Identifier;
        return { type: 'EraseStatement', name };
    }

    private parseReDimStatement(): ReDimStatement {
        this.advance(); // 'ReDim'

        // Optional 'Preserve' keyword
        if (this.peek().type === TokenType.Identifier && this.peek().value.toLowerCase() === 'preserve') {
            this.advance();
        }

        const idToken = this.advance();
        const name = { type: 'Identifier', name: idToken.value } as Identifier;
        const bounds: Expression[] = [];

        if (this.match(TokenType.OperatorLParen)) {
            if (this.peek().type !== TokenType.OperatorRParen) {
                bounds.push(this.parseExpression());
                while (this.match(TokenType.KeywordTo) || this.match(TokenType.OperatorComma)) {
                    bounds.push(this.parseExpression());
                }
            }
            this.match(TokenType.OperatorRParen);
        }

        if (this.match(TokenType.KeywordAs)) {
            this.advance();
        }

        return { type: 'ReDimStatement', name, bounds };
    }

    private parseForStatement(): ForStatement {
        this.advance(); // consume 'For'

        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'For' at line ${idToken.line} `);
        }
        const identifier: Identifier = { type: 'Identifier', name: idToken.value };

        if (!this.match(TokenType.OperatorEquals)) {
            throw new Error(`Parse error: Expected '=' in For statement at line ${this.peek().line} `);
        }

        const startExpr = this.parseExpression();

        if (!this.match(TokenType.KeywordTo)) {
            throw new Error(`Parse error: Expected 'To' in For statement at line ${this.peek().line} `);
        }

        const endExpr = this.parseExpression();

        let stepExpr: Expression | undefined;
        if (this.match(TokenType.KeywordStep)) {
            stepExpr = this.parseExpression();
        }

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            throw new Error(`Parse error: Expected 'Next' at line ${this.peek().line} `);
        }

        let nextIdentifier: Identifier | undefined;
        if (this.peek().type === TokenType.Identifier) {
            const nextIdToken = this.advance();
            nextIdentifier = { type: 'Identifier', name: nextIdToken.value } as Identifier;
        }

        return {
            type: 'ForStatement',
            identifier,
            start: startExpr,
            end: endExpr,
            step: stepExpr,
            body,
            nextIdentifier
        };
    }



    private parseIfStatement(): IfStatement {
        this.advance(); // Consume 'If' or 'ElseIf'
        const condition = this.parseExpression();

        if (!this.match(TokenType.KeywordThen)) {
            throw new Error(`Parse error: Expected 'Then' after condition at line ${this.peek().line}`);
        }

        const isMultiLine = this.peek().type === TokenType.Newline;

        const consequent: Statement[] = [];
        let alternate: Statement[] | IfStatement | null = null;

        if (!isMultiLine) {
            // Single-line If
            while (this.peek().type !== TokenType.Newline && this.peek().type !== TokenType.EOF) {
                const stmt = this.parseStatement();
                if (stmt) consequent.push(stmt);
            }
            // No End If expected
            return {
                type: 'IfStatement',
                condition,
                consequent,
                alternate
            };
        }

        this.skipNewlines();

        while (
            this.peek().type !== TokenType.KeywordEnd &&
            this.peek().type !== TokenType.KeywordElse &&
            this.peek().type !== TokenType.KeywordElseIf &&
            this.peek().type !== TokenType.EOF
        ) {
            const stmt = this.parseStatement();
            if (stmt) consequent.push(stmt);
            this.skipNewlines();
        }

        if (this.peek().type === TokenType.KeywordElseIf) {
            alternate = this.parseIfStatement(); // Recursive ElseIf
        } else if (this.match(TokenType.KeywordElse)) {
            this.skipNewlines();
            alternate = [];
            while (
                this.peek().type !== TokenType.KeywordEnd &&
                this.peek().type !== TokenType.EOF
            ) {
                const stmt = this.parseStatement();
                if (stmt) alternate.push(stmt);
                this.skipNewlines();
            }
        }

        // Only end the top-level IF chain with End If
        if (this.peek().type === TokenType.KeywordEnd) {
            this.advance(); // Consume 'End'
            if (!this.match(TokenType.KeywordIf)) {
                throw new Error(`Parse error: Expected 'If' after 'End' at line ${this.peek().line}`);
            }
        }

        return {
            type: 'IfStatement',
            condition,
            consequent,
            alternate
        };
    }

    private parseDoWhileStatement(): DoWhileStatement {
        this.advance(); // consume 'Do'
        if (!this.match(TokenType.KeywordWhile)) {
            throw new Error(`Parse error: Expected 'While' after 'Do' at line ${this.peek().line} `);
        }

        const condition = this.parseExpression();
        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordLoop && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordLoop)) {
            throw new Error(`Parse error: Expected 'Loop' at line ${this.peek().line} `);
        }

        return {
            type: 'DoWhileStatement',
            condition,
            body
        };
    }

    private parseExpression(): Expression {
        return this.parseLogicalOr();
    }

    private parseLogicalOr(): Expression {
        let left = this.parseLogicalAnd();
        while (this.peek().type === TokenType.KeywordOr) {
            const operator = this.advance().value;
            const right = this.parseLogicalAnd();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalAnd(): Expression {
        let left = this.parseLogicalNot();
        while (this.peek().type === TokenType.KeywordAnd) {
            const operator = this.advance().value;
            const right = this.parseLogicalNot();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseLogicalNot(): Expression {
        if (this.match(TokenType.KeywordNot)) {
            const argument = this.parseEquality();
            return { type: 'UnaryExpression', operator: 'Not', argument } as UnaryExpression;
        }
        return this.parseEquality();
    }

    private parseEquality(): Expression {
        let left = this.parseRelational();
        while (
            this.peek().type === TokenType.OperatorEquals ||
            this.peek().type === TokenType.OperatorNotEquals
        ) {
            const operator = this.advance().value;
            const right = this.parseRelational();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseRelational(): Expression {
        let left = this.parseAdditive();
        while (
            this.peek().type === TokenType.OperatorLessThan ||
            this.peek().type === TokenType.OperatorGreaterThan ||
            this.peek().type === TokenType.OperatorLessThanOrEqual ||
            this.peek().type === TokenType.OperatorGreaterThanOrEqual
        ) {
            const operator = this.advance().value;
            const right = this.parseAdditive();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseAdditive(): Expression {
        let left = this.parseModulo();
        while (
            this.peek().type === TokenType.OperatorPlus ||
            this.peek().type === TokenType.OperatorMinus
        ) {
            const operator = this.advance().value;
            const right = this.parseModulo();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseModulo(): Expression {
        let left = this.parseIntDivision();
        while (this.peek().type === TokenType.KeywordMod) {
            const operator = this.advance().value;
            const right = this.parseIntDivision();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseIntDivision(): Expression {
        let left = this.parseMultiplicative();
        while (this.peek().type === TokenType.OperatorIntDivide) {
            const operator = this.advance().value;
            const right = this.parseMultiplicative();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseMultiplicative(): Expression {
        let left = this.parseUnary();
        while (
            this.peek().type === TokenType.OperatorMultiply ||
            this.peek().type === TokenType.OperatorDivide
        ) {
            const operator = this.advance().value;
            const right = this.parseUnary();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parseUnary(): Expression {
        if (this.peek().type === TokenType.OperatorMinus || this.peek().type === TokenType.OperatorPlus) {
            const operator = this.advance().value;
            const argument = this.parseUnary();
            return { type: 'UnaryExpression', operator, argument } as UnaryExpression;
        }
        return this.parseExponentiation();
    }

    private parseExponentiation(): Expression {
        let left = this.parsePrimary();
        while (this.peek().type === TokenType.OperatorPower) {
            const operator = this.advance().value;
            const right = this.parsePrimary();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parsePrimary(): Expression {
        const token = this.advance();
        let expr: Expression;
        if (token.type === TokenType.Number) {
            // Remove VBA type suffix (%, &, @, !, #) before parsing into a float
            const cleanVal = token.value.replace(/[%&@!#]$/, '');
            expr = { type: 'NumberLiteral', value: parseFloat(cleanVal) } as NumberLiteral;
        } else if (token.type === TokenType.String) {
            expr = { type: 'StringLiteral', value: token.value } as StringLiteral;
        } else if (token.type === TokenType.Identifier) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.KeywordEmpty) {
            expr = { type: 'Identifier', name: token.value } as Identifier;
        } else if (token.type === TokenType.OperatorLParen) {
            expr = this.parseExpression();
            if (!this.match(TokenType.OperatorRParen)) {
                throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
            }
        } else {
            throw new Error(`Parse error: Unexpected token in expression '${token.value}' at line ${token.line} `);
        }

        while (true) {
            if (this.match(TokenType.OperatorDot)) {
                const propToken = this.advance();
                const property = { type: 'Identifier', name: propToken.value } as Identifier;
                expr = { type: 'MemberExpression', object: expr, property } as MemberExpression;
            } else if (this.match(TokenType.OperatorLParen)) {
                const args: Expression[] = [];
                if (this.peek().type !== TokenType.OperatorRParen) {
                    args.push(this.parseExpression());
                    while (this.match(TokenType.OperatorComma)) {
                        args.push(this.parseExpression());
                    }
                }
                if (!this.match(TokenType.OperatorRParen)) {
                    throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
                }
                expr = { type: 'CallExpression', callee: expr, args } as CallExpression;
            } else {
                break;
            }
        }
        return expr;
    }
}
