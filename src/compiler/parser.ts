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

export interface DebugPrintStatement extends Statement {
    type: 'DebugPrintStatement';
    expression: Expression;
}

export interface Expression extends ASTNode { }

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
        } else if (token.type === TokenType.Identifier) {
            // Can be assignment (count = count + 1)
            // or just a call (though we only have debug.print so far)
            if (this.pos + 1 < this.tokens.length && this.tokens[this.pos + 1].type === TokenType.OperatorEquals) {
                return this.parseAssignmentStatement();
            }
        } else if (token.type === TokenType.KeywordDebugPrint) {
            return this.parseDebugPrintStatement();
        } else if (token.type === TokenType.Unknown) {
            throw new Error(`Parse error: Unknown token '${token.value}' at line ${token.line}`);
        } else {
            // Unknown or unexpected top-level token
            this.advance();
        }
        return null;
    }

    private parseForStatement(): ForStatement {
        this.advance(); // consume 'For'

        const idToken = this.advance();
        if (idToken.type !== TokenType.Identifier) {
            throw new Error(`Parse error: Expected identifier after 'For' at line ${idToken.line}`);
        }
        const identifier: Identifier = { type: 'Identifier', name: idToken.value };

        if (!this.match(TokenType.OperatorEquals)) {
            throw new Error(`Parse error: Expected '=' in For statement at line ${this.peek().line}`);
        }

        const startExpr = this.parseExpression();

        if (!this.match(TokenType.KeywordTo)) {
            throw new Error(`Parse error: Expected 'To' in For statement at line ${this.peek().line}`);
        }

        const endExpr = this.parseExpression();

        this.skipNewlines();

        const body: Statement[] = [];
        while (this.peek().type !== TokenType.KeywordNext && this.peek().type !== TokenType.EOF) {
            const stmt = this.parseStatement();
            if (stmt) body.push(stmt);
            this.skipNewlines();
        }

        if (!this.match(TokenType.KeywordNext)) {
            throw new Error(`Parse error: Expected 'Next' at line ${this.peek().line}`);
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
            body,
            nextIdentifier
        };
    }

    private parseAssignmentStatement(): Statement {
        const idToken = this.advance();
        const identifier: Identifier = { type: 'Identifier', name: idToken.value } as Identifier;

        this.advance(); // consume '='

        const expression = this.parseExpression();
        // Since we don't have an Assignment AST node, we'll fake it as a binary operation
        // that evaluator handles specifically (by convention, not great but works for Step 2)
        // A better approach is explicit AssignmentStatement
        return {
            type: 'AssignmentStatement',
            identifier,
            value: expression
        } as any;
    }

    private parseIfStatement(): IfStatement {
        this.advance(); // Consume 'If' or 'ElseIf'
        const condition = this.parseExpression();

        if (!this.match(TokenType.KeywordThen)) {
            throw new Error(`Parse error: Expected 'Then' after condition at line ${this.peek().line}`);
        }

        this.skipNewlines();

        const consequent: Statement[] = [];
        let alternate: Statement[] | IfStatement | null = null;

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
            throw new Error(`Parse error: Expected 'While' after 'Do' at line ${this.peek().line}`);
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
            throw new Error(`Parse error: Expected 'Loop' at line ${this.peek().line}`);
        }

        return {
            type: 'DoWhileStatement',
            condition,
            body
        };
    }

    private parseDebugPrintStatement(): DebugPrintStatement {
        this.advance(); // consume 'Debug.Print'
        const expression = this.parseExpression();
        return {
            type: 'DebugPrintStatement',
            expression
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
        let left = this.parseEquality();
        while (this.peek().type === TokenType.KeywordAnd) {
            const operator = this.advance().value;
            const right = this.parseEquality();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
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
        let left = this.parsePrimary();
        while (
            this.peek().type === TokenType.OperatorPlus ||
            this.peek().type === TokenType.OperatorMinus
        ) {
            const operator = this.advance().value;
            const right = this.parsePrimary();
            left = { type: 'BinaryExpression', operator, left, right } as BinaryExpression;
        }
        return left;
    }

    private parsePrimary(): Expression {
        const token = this.advance();
        if (token.type === TokenType.Number) {
            return { type: 'NumberLiteral', value: parseFloat(token.value) } as NumberLiteral;
        } else if (token.type === TokenType.String) {
            return { type: 'StringLiteral', value: token.value } as StringLiteral;
        } else if (token.type === TokenType.Identifier) {
            return { type: 'Identifier', name: token.value } as Identifier;
        }
        throw new Error(`Parse error: Unexpected token in expression '${token.value}' at line ${token.line}`);
    }
}
