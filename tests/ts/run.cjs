var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));

// tests/ts/test-runner.ts
var fs = __toESM(require("fs"), 1);

// src/compiler/lexer.ts
var Lexer = class {
  input = "";
  pos = 0;
  line = 1;
  constructor(input) {
    this.input = input;
  }
  peek() {
    if (this.pos >= this.input.length) return "\0";
    return this.input[this.pos];
  }
  advance() {
    if (this.pos >= this.input.length) return "\0";
    const char = this.input[this.pos++];
    if (char === "\n") {
      this.line++;
    }
    return char;
  }
  isWhitespace(char) {
    return char === " " || char === "	" || char === "\r";
  }
  isAlpha(char) {
    return char >= "a" && char <= "z" || char >= "A" && char <= "Z" || char === "_";
  }
  isAlphaNumeric(char) {
    return this.isAlpha(char) || char >= "0" && char <= "9";
  }
  isDigit(char) {
    return char >= "0" && char <= "9";
  }
  skipWhitespace() {
    while (this.isWhitespace(this.peek())) {
      this.advance();
    }
  }
  getNextToken() {
    while (true) {
      this.skipWhitespace();
      if (this.pos >= this.input.length) {
        return { type: 56 /* EOF */, value: "", line: this.line };
      }
      const char = this.peek();
      if (char === "'") {
        while (this.peek() !== "\n" && this.peek() !== "\0") {
          this.advance();
        }
        continue;
      }
      if (char === "\n") {
        this.advance();
        return { type: 55 /* Newline */, value: "\n", line: this.line - 1 };
      }
      if (char === '"') {
        this.advance();
        let strValue = "";
        while (this.peek() !== '"' && this.peek() !== "\0") {
          strValue += this.advance();
        }
        if (this.peek() === '"') {
          this.advance();
        }
        return { type: 2 /* String */, value: strValue, line: this.line };
      }
      if (char === "=") {
        this.advance();
        return { type: 44 /* OperatorEquals */, value: "=", line: this.line };
      }
      if (char === "<") {
        this.advance();
        if (this.peek() === ">") {
          this.advance();
          return { type: 45 /* OperatorNotEquals */, value: "<>", line: this.line };
        }
        if (this.peek() === "=") {
          this.advance();
          return { type: 48 /* OperatorLessThanOrEqual */, value: "<=", line: this.line };
        }
        return { type: 46 /* OperatorLessThan */, value: "<", line: this.line };
      }
      if (char === ">") {
        this.advance();
        if (this.peek() === "=") {
          this.advance();
          return { type: 49 /* OperatorGreaterThanOrEqual */, value: ">=", line: this.line };
        }
        return { type: 47 /* OperatorGreaterThan */, value: ">", line: this.line };
      }
      if (char === "+") {
        this.advance();
        return { type: 37 /* OperatorPlus */, value: "+", line: this.line };
      }
      if (char === "-") {
        this.advance();
        return { type: 38 /* OperatorMinus */, value: "-", line: this.line };
      }
      if (char === ",") {
        this.advance();
        return { type: 50 /* OperatorComma */, value: ",", line: this.line };
      }
      if (char === "(") {
        this.advance();
        return { type: 51 /* OperatorLParen */, value: "(", line: this.line };
      }
      if (char === ")") {
        this.advance();
        return { type: 52 /* OperatorRParen */, value: ")", line: this.line };
      }
      if (char === ".") {
        this.advance();
        return { type: 53 /* OperatorDot */, value: ".", line: this.line };
      }
      if (char === ":") {
        this.advance();
        return { type: 54 /* OperatorColon */, value: ":", line: this.line };
      }
      if (this.isDigit(char)) {
        let numStr = "";
        while (this.isDigit(this.peek())) {
          numStr += this.advance();
        }
        if (this.peek() === ".") {
          numStr += this.advance();
          while (this.isDigit(this.peek())) {
            numStr += this.advance();
          }
        }
        const peekChar = this.peek();
        if (["%", "&", "@", "!", "#"].indexOf(peekChar) !== -1) {
          numStr += this.advance();
        }
        return { type: 1 /* Number */, value: numStr, line: this.line };
      }
      if (this.isAlpha(char)) {
        let idStr = "";
        while (this.isAlphaNumeric(this.peek())) {
          idStr += this.advance();
        }
        const lowerId = idStr.toLowerCase();
        if (lowerId === "rem") {
          while (this.peek() !== "\n" && this.peek() !== "\0") {
            this.advance();
          }
          continue;
        }
        if (lowerId === "for") return { type: 3 /* KeywordFor */, value: idStr, line: this.line };
        if (lowerId === "to") return { type: 4 /* KeywordTo */, value: idStr, line: this.line };
        if (lowerId === "next") return { type: 5 /* KeywordNext */, value: idStr, line: this.line };
        if (lowerId === "if") return { type: 6 /* KeywordIf */, value: idStr, line: this.line };
        if (lowerId === "then") return { type: 7 /* KeywordThen */, value: idStr, line: this.line };
        if (lowerId === "elseif") return { type: 8 /* KeywordElseIf */, value: idStr, line: this.line };
        if (lowerId === "else") return { type: 9 /* KeywordElse */, value: idStr, line: this.line };
        if (lowerId === "end") return { type: 10 /* KeywordEnd */, value: idStr, line: this.line };
        if (lowerId === "do") return { type: 11 /* KeywordDo */, value: idStr, line: this.line };
        if (lowerId === "while") return { type: 12 /* KeywordWhile */, value: idStr, line: this.line };
        if (lowerId === "loop") return { type: 13 /* KeywordLoop */, value: idStr, line: this.line };
        if (lowerId === "sub") return { type: 14 /* KeywordSub */, value: idStr, line: this.line };
        if (lowerId === "function") return { type: 15 /* KeywordFunction */, value: idStr, line: this.line };
        if (lowerId === "dim") return { type: 16 /* KeywordDim */, value: idStr, line: this.line };
        if (lowerId === "as") return { type: 17 /* KeywordAs */, value: idStr, line: this.line };
        if (lowerId === "new") return { type: 18 /* KeywordNew */, value: idStr, line: this.line };
        if (lowerId === "collection") return { type: 19 /* KeywordCollection */, value: idStr, line: this.line };
        if (lowerId === "and") return { type: 20 /* KeywordAnd */, value: idStr, line: this.line };
        if (lowerId === "or") return { type: 21 /* KeywordOr */, value: idStr, line: this.line };
        if (lowerId === "not") return { type: 22 /* KeywordNot */, value: idStr, line: this.line };
        if (lowerId === "option") return { type: 23 /* KeywordOption */, value: idStr, line: this.line };
        if (lowerId === "explicit") return { type: 24 /* KeywordExplicit */, value: idStr, line: this.line };
        if (lowerId === "const") return { type: 25 /* KeywordConst */, value: idStr, line: this.line };
        if (lowerId === "set") return { type: 26 /* KeywordSet */, value: idStr, line: this.line };
        if (lowerId === "on") return { type: 27 /* KeywordOn */, value: idStr, line: this.line };
        if (lowerId === "error") return { type: 28 /* KeywordError */, value: idStr, line: this.line };
        if (lowerId === "goto") return { type: 29 /* KeywordGoTo */, value: idStr, line: this.line };
        if (lowerId === "erase") return { type: 30 /* KeywordErase */, value: idStr, line: this.line };
        if (lowerId === "redim") return { type: 31 /* KeywordReDim */, value: idStr, line: this.line };
        if (lowerId === "step") return { type: 32 /* KeywordStep */, value: idStr, line: this.line };
        if (lowerId === "empty") return { type: 33 /* KeywordEmpty */, value: idStr, line: this.line };
        if (lowerId === "exit") return { type: 34 /* KeywordExit */, value: idStr, line: this.line };
        if (lowerId === "byref") return { type: 35 /* KeywordByRef */, value: idStr, line: this.line };
        if (lowerId === "byval") return { type: 36 /* KeywordByVal */, value: idStr, line: this.line };
        if (lowerId === "mod") return { type: 42 /* KeywordMod */, value: idStr, line: this.line };
        return { type: 0 /* Identifier */, value: idStr, line: this.line };
      }
      if (char === "*") {
        this.advance();
        return { type: 39 /* OperatorMultiply */, value: "*", line: this.line };
      }
      if (char === "/") {
        this.advance();
        return { type: 40 /* OperatorDivide */, value: "/", line: this.line };
      }
      if (char === "\\") {
        this.advance();
        return { type: 41 /* OperatorIntDivide */, value: "\\", line: this.line };
      }
      if (char === "^") {
        this.advance();
        return { type: 43 /* OperatorPower */, value: "^", line: this.line };
      }
      const unknownChar = this.advance();
      return { type: 57 /* Unknown */, value: unknownChar, line: this.line };
    }
  }
  tokenize() {
    const tokens = [];
    let token;
    do {
      token = this.getNextToken();
      tokens.push(token);
    } while (token.type !== 56 /* EOF */);
    return tokens;
  }
};

// src/compiler/parser.ts
var Parser = class {
  tokens;
  pos = 0;
  constructor(tokens) {
    this.tokens = tokens;
  }
  peek() {
    if (this.pos >= this.tokens.length) {
      return this.tokens[this.tokens.length - 1];
    }
    return this.tokens[this.pos];
  }
  advance() {
    const token = this.peek();
    this.pos++;
    return token;
  }
  match(expectedType) {
    if (this.peek().type === expectedType) {
      this.advance();
      return true;
    }
    return false;
  }
  skipNewlines() {
    while (this.match(55 /* Newline */)) {
    }
  }
  parse() {
    const program = {
      type: "Program",
      body: []
    };
    this.skipNewlines();
    while (this.peek().type !== 56 /* EOF */) {
      const stmt = this.parseStatement();
      if (stmt) {
        program.body.push(stmt);
      }
      this.skipNewlines();
    }
    return program;
  }
  parseStatement() {
    const token = this.peek();
    if (token.type === 3 /* KeywordFor */) {
      return this.parseForStatement();
    } else if (token.type === 6 /* KeywordIf */) {
      return this.parseIfStatement();
    } else if (token.type === 11 /* KeywordDo */) {
      return this.parseDoWhileStatement();
    } else if (token.type === 14 /* KeywordSub */ || token.type === 15 /* KeywordFunction */) {
      return this.parseProcedureDeclaration();
    } else if (token.type === 16 /* KeywordDim */) {
      return this.parseDimStatement();
    } else if (token.type === 25 /* KeywordConst */) {
      return this.parseConstDeclaration();
    } else if (token.type === 26 /* KeywordSet */) {
      return this.parseSetStatement();
    } else if (token.type === 27 /* KeywordOn */) {
      return this.parseOnErrorStatement();
    } else if (token.type === 34 /* KeywordExit */) {
      return this.parseExitStatement();
    } else if (token.type === 30 /* KeywordErase */) {
      return this.parseEraseStatement();
    } else if (token.type === 31 /* KeywordReDim */) {
      return this.parseReDimStatement();
    } else if (token.type === 23 /* KeywordOption */) {
      this.advance();
      if (this.match(24 /* KeywordExplicit */)) {
      }
      return null;
    } else if (token.type === 0 /* Identifier */) {
      if (this.pos + 1 < this.tokens.length && this.tokens[this.pos + 1].type === 54 /* OperatorColon */) {
        const labelName = token.value;
        this.advance();
        this.advance();
        return { type: "LabelStatement", label: labelName };
      }
      const expr = this.parsePrimary();
      if (this.match(44 /* OperatorEquals */)) {
        return {
          type: "AssignmentStatement",
          left: expr,
          right: this.parseExpression()
        };
      } else {
        const args = [];
        if (this.peek().type !== 55 /* Newline */ && this.peek().type !== 56 /* EOF */ && this.peek().type !== 9 /* KeywordElse */ && this.peek().type !== 8 /* KeywordElseIf */ && this.peek().type !== 10 /* KeywordEnd */ && this.peek().type !== 5 /* KeywordNext */ && this.peek().type !== 13 /* KeywordLoop */) {
          args.push(this.parseExpression());
          while (this.match(50 /* OperatorComma */)) {
            args.push(this.parseExpression());
          }
        }
        if (args.length > 0) {
          return { type: "CallStatement", expression: { type: "CallExpression", callee: expr, args } };
        } else if (expr.type === "CallExpression") {
          return { type: "CallStatement", expression: expr };
        } else {
          return { type: "CallStatement", expression: { type: "CallExpression", callee: expr, args: [] } };
        }
      }
    } else if (token.type === 57 /* Unknown */) {
      throw new Error(`Parse error: Unknown token '${token.value}' at line ${token.line}`);
    } else {
      this.advance();
    }
    return null;
  }
  parseProcedureDeclaration() {
    const isFunction = this.peek().type === 15 /* KeywordFunction */;
    this.advance();
    const idToken = this.advance();
    if (idToken.type !== 0 /* Identifier */) {
      throw new Error(`Parse error: Expected identifier after Sub/Function at line ${idToken.line}`);
    }
    const name = { type: "Identifier", name: idToken.value };
    const parameters = [];
    if (this.match(51 /* OperatorLParen */)) {
      if (this.peek().type !== 52 /* OperatorRParen */) {
        let paramNameToken = this.peek();
        if (paramNameToken.type === 36 /* KeywordByVal */ || paramNameToken.type === 35 /* KeywordByRef */) {
          this.advance();
          paramNameToken = this.peek();
        }
        let paramName = this.advance();
        if (this.match(17 /* KeywordAs */)) {
          this.advance();
        }
        parameters.push({ type: "Parameter", name: paramName.value });
        while (this.match(50 /* OperatorComma */)) {
          let nextParamToken = this.peek();
          if (nextParamToken.type === 36 /* KeywordByVal */ || nextParamToken.type === 35 /* KeywordByRef */) {
            this.advance();
            nextParamToken = this.peek();
          }
          paramName = this.advance();
          if (this.match(17 /* KeywordAs */)) {
            this.advance();
          }
          parameters.push({ type: "Parameter", name: paramName.value });
        }
      }
      if (!this.match(52 /* OperatorRParen */)) {
        throw new Error(`Parse error: Expected ')' at line ${this.peek().line}`);
      }
    }
    if (this.match(17 /* KeywordAs */)) {
      this.advance();
    }
    this.skipNewlines();
    const body = [];
    while (this.peek().type !== 10 /* KeywordEnd */ && this.peek().type !== 56 /* EOF */) {
      const stmt = this.parseStatement();
      if (stmt) body.push(stmt);
      this.skipNewlines();
    }
    if (this.peek().type === 10 /* KeywordEnd */) {
      this.advance();
      const expectedEndStr = isFunction ? "Function" : "Sub";
      const endToken = this.advance();
      if (endToken.value.toLowerCase() !== expectedEndStr.toLowerCase()) {
        throw new Error(`Parse error: Expected '${expectedEndStr}' after 'End' at line ${endToken.line}`);
      }
    }
    return { type: "ProcedureDeclaration", isFunction, name, parameters, body };
  }
  parseDimStatement() {
    this.advance();
    const declarations = [];
    while (true) {
      const idToken = this.advance();
      const name = { type: "Identifier", name: idToken.value };
      let isArray = false;
      let arraySize;
      let isNew = false;
      let objectType;
      if (this.match(51 /* OperatorLParen */)) {
        isArray = true;
        if (this.peek().type !== 52 /* OperatorRParen */) {
          arraySize = this.parseExpression();
        }
        this.match(52 /* OperatorRParen */);
      }
      if (this.match(17 /* KeywordAs */)) {
        if (this.match(18 /* KeywordNew */)) {
          isNew = true;
        }
        const typeToken = this.peek();
        if (typeToken.type === 19 /* KeywordCollection */ || typeToken.type === 0 /* Identifier */) {
          objectType = this.advance().value;
        }
      }
      declarations.push({ name, isArray, arraySize, isNew, objectType });
      if (this.match(50 /* OperatorComma */)) {
        continue;
      } else {
        break;
      }
    }
    return { type: "VariableDeclaration", declarations };
  }
  parseConstDeclaration() {
    this.advance();
    const idToken = this.advance();
    if (idToken.type !== 0 /* Identifier */) throw new Error(`Parse error: Expected identifier after Const at line ${idToken.line}`);
    const name = { type: "Identifier", name: idToken.value };
    if (this.match(17 /* KeywordAs */)) {
      this.advance();
    }
    if (!this.match(44 /* OperatorEquals */)) throw new Error(`Parse error: Expected '=' in Const at line ${this.peek().line}`);
    const value = this.parseExpression();
    return { type: "ConstDeclaration", name, value };
  }
  parseSetStatement() {
    this.advance();
    const left = this.parsePrimary();
    if (!this.match(44 /* OperatorEquals */)) throw new Error(`Parse error: Expected '=' in Set statement at line ${this.peek().line}`);
    const right = this.parseExpression();
    return { type: "SetStatement", left, right };
  }
  parseOnErrorStatement() {
    this.advance();
    if (!this.match(28 /* KeywordError */)) throw new Error(`Parse error: Expected 'Error' after 'On' at line ${this.peek().line}`);
    let label = "";
    if (this.match(29 /* KeywordGoTo */)) {
      const labelToken = this.advance();
      label = labelToken.value;
    } else {
      while (this.peek().type !== 55 /* Newline */ && this.peek().type !== 56 /* EOF */) {
        label += this.advance().value + " ";
      }
      label = label.trim();
    }
    return { type: "OnErrorStatement", label };
  }
  parseExitStatement() {
    this.advance();
    const typeToken = this.advance();
    let exitType;
    if (typeToken.type === 3 /* KeywordFor */) {
      exitType = "For";
    } else if (typeToken.type === 11 /* KeywordDo */) {
      exitType = "Do";
    } else if (typeToken.type === 14 /* KeywordSub */) {
      exitType = "Sub";
    } else if (typeToken.type === 15 /* KeywordFunction */) {
      exitType = "Function";
    } else {
      throw new Error(`Parse error: Unexpected token after Exit '${typeToken.value}' at line ${typeToken.line} `);
    }
    return { type: "ExitStatement", exitType };
  }
  parseEraseStatement() {
    this.advance();
    const idToken = this.advance();
    const name = { type: "Identifier", name: idToken.value };
    return { type: "EraseStatement", name };
  }
  parseReDimStatement() {
    this.advance();
    if (this.peek().type === 0 /* Identifier */ && this.peek().value.toLowerCase() === "preserve") {
      this.advance();
    }
    const idToken = this.advance();
    const name = { type: "Identifier", name: idToken.value };
    const bounds = [];
    if (this.match(51 /* OperatorLParen */)) {
      if (this.peek().type !== 52 /* OperatorRParen */) {
        bounds.push(this.parseExpression());
        while (this.match(4 /* KeywordTo */) || this.match(50 /* OperatorComma */)) {
          bounds.push(this.parseExpression());
        }
      }
      this.match(52 /* OperatorRParen */);
    }
    if (this.match(17 /* KeywordAs */)) {
      this.advance();
    }
    return { type: "ReDimStatement", name, bounds };
  }
  parseForStatement() {
    this.advance();
    const idToken = this.advance();
    if (idToken.type !== 0 /* Identifier */) {
      throw new Error(`Parse error: Expected identifier after 'For' at line ${idToken.line} `);
    }
    const identifier = { type: "Identifier", name: idToken.value };
    if (!this.match(44 /* OperatorEquals */)) {
      throw new Error(`Parse error: Expected '=' in For statement at line ${this.peek().line} `);
    }
    const startExpr = this.parseExpression();
    if (!this.match(4 /* KeywordTo */)) {
      throw new Error(`Parse error: Expected 'To' in For statement at line ${this.peek().line} `);
    }
    const endExpr = this.parseExpression();
    let stepExpr;
    if (this.match(32 /* KeywordStep */)) {
      stepExpr = this.parseExpression();
    }
    this.skipNewlines();
    const body = [];
    while (this.peek().type !== 5 /* KeywordNext */ && this.peek().type !== 56 /* EOF */) {
      const stmt = this.parseStatement();
      if (stmt) body.push(stmt);
      this.skipNewlines();
    }
    if (!this.match(5 /* KeywordNext */)) {
      throw new Error(`Parse error: Expected 'Next' at line ${this.peek().line} `);
    }
    let nextIdentifier;
    if (this.peek().type === 0 /* Identifier */) {
      const nextIdToken = this.advance();
      nextIdentifier = { type: "Identifier", name: nextIdToken.value };
    }
    return {
      type: "ForStatement",
      identifier,
      start: startExpr,
      end: endExpr,
      step: stepExpr,
      body,
      nextIdentifier
    };
  }
  parseIfStatement() {
    this.advance();
    const condition = this.parseExpression();
    if (!this.match(7 /* KeywordThen */)) {
      throw new Error(`Parse error: Expected 'Then' after condition at line ${this.peek().line}`);
    }
    const isMultiLine = this.peek().type === 55 /* Newline */;
    const consequent = [];
    let alternate = null;
    if (!isMultiLine) {
      while (this.peek().type !== 55 /* Newline */ && this.peek().type !== 56 /* EOF */) {
        const stmt = this.parseStatement();
        if (stmt) consequent.push(stmt);
      }
      return {
        type: "IfStatement",
        condition,
        consequent,
        alternate
      };
    }
    this.skipNewlines();
    while (this.peek().type !== 10 /* KeywordEnd */ && this.peek().type !== 9 /* KeywordElse */ && this.peek().type !== 8 /* KeywordElseIf */ && this.peek().type !== 56 /* EOF */) {
      const stmt = this.parseStatement();
      if (stmt) consequent.push(stmt);
      this.skipNewlines();
    }
    if (this.peek().type === 8 /* KeywordElseIf */) {
      alternate = this.parseIfStatement();
    } else if (this.match(9 /* KeywordElse */)) {
      this.skipNewlines();
      alternate = [];
      while (this.peek().type !== 10 /* KeywordEnd */ && this.peek().type !== 56 /* EOF */) {
        const stmt = this.parseStatement();
        if (stmt) alternate.push(stmt);
        this.skipNewlines();
      }
    }
    if (this.peek().type === 10 /* KeywordEnd */) {
      this.advance();
      if (!this.match(6 /* KeywordIf */)) {
        throw new Error(`Parse error: Expected 'If' after 'End' at line ${this.peek().line}`);
      }
    }
    return {
      type: "IfStatement",
      condition,
      consequent,
      alternate
    };
  }
  parseDoWhileStatement() {
    this.advance();
    if (!this.match(12 /* KeywordWhile */)) {
      throw new Error(`Parse error: Expected 'While' after 'Do' at line ${this.peek().line} `);
    }
    const condition = this.parseExpression();
    this.skipNewlines();
    const body = [];
    while (this.peek().type !== 13 /* KeywordLoop */ && this.peek().type !== 56 /* EOF */) {
      const stmt = this.parseStatement();
      if (stmt) body.push(stmt);
      this.skipNewlines();
    }
    if (!this.match(13 /* KeywordLoop */)) {
      throw new Error(`Parse error: Expected 'Loop' at line ${this.peek().line} `);
    }
    return {
      type: "DoWhileStatement",
      condition,
      body
    };
  }
  parseExpression() {
    return this.parseLogicalOr();
  }
  parseLogicalOr() {
    let left = this.parseLogicalAnd();
    while (this.peek().type === 21 /* KeywordOr */) {
      const operator = this.advance().value;
      const right = this.parseLogicalAnd();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseLogicalAnd() {
    let left = this.parseLogicalNot();
    while (this.peek().type === 20 /* KeywordAnd */) {
      const operator = this.advance().value;
      const right = this.parseLogicalNot();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseLogicalNot() {
    if (this.match(22 /* KeywordNot */)) {
      const argument = this.parseEquality();
      return { type: "UnaryExpression", operator: "Not", argument };
    }
    return this.parseEquality();
  }
  parseEquality() {
    let left = this.parseRelational();
    while (this.peek().type === 44 /* OperatorEquals */ || this.peek().type === 45 /* OperatorNotEquals */) {
      const operator = this.advance().value;
      const right = this.parseRelational();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseRelational() {
    let left = this.parseAdditive();
    while (this.peek().type === 46 /* OperatorLessThan */ || this.peek().type === 47 /* OperatorGreaterThan */ || this.peek().type === 48 /* OperatorLessThanOrEqual */ || this.peek().type === 49 /* OperatorGreaterThanOrEqual */) {
      const operator = this.advance().value;
      const right = this.parseAdditive();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseAdditive() {
    let left = this.parseModulo();
    while (this.peek().type === 37 /* OperatorPlus */ || this.peek().type === 38 /* OperatorMinus */) {
      const operator = this.advance().value;
      const right = this.parseModulo();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseModulo() {
    let left = this.parseIntDivision();
    while (this.peek().type === 42 /* KeywordMod */) {
      const operator = this.advance().value;
      const right = this.parseIntDivision();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseIntDivision() {
    let left = this.parseMultiplicative();
    while (this.peek().type === 41 /* OperatorIntDivide */) {
      const operator = this.advance().value;
      const right = this.parseMultiplicative();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseMultiplicative() {
    let left = this.parseUnary();
    while (this.peek().type === 39 /* OperatorMultiply */ || this.peek().type === 40 /* OperatorDivide */) {
      const operator = this.advance().value;
      const right = this.parseUnary();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parseUnary() {
    if (this.peek().type === 38 /* OperatorMinus */ || this.peek().type === 37 /* OperatorPlus */) {
      const operator = this.advance().value;
      const argument = this.parseUnary();
      return { type: "UnaryExpression", operator, argument };
    }
    return this.parseExponentiation();
  }
  parseExponentiation() {
    let left = this.parsePrimary();
    while (this.peek().type === 43 /* OperatorPower */) {
      const operator = this.advance().value;
      const right = this.parsePrimary();
      left = { type: "BinaryExpression", operator, left, right };
    }
    return left;
  }
  parsePrimary() {
    const token = this.advance();
    let expr;
    if (token.type === 1 /* Number */) {
      const cleanVal = token.value.replace(/[%&@!#]$/, "");
      expr = { type: "NumberLiteral", value: parseFloat(cleanVal) };
    } else if (token.type === 2 /* String */) {
      expr = { type: "StringLiteral", value: token.value };
    } else if (token.type === 0 /* Identifier */) {
      expr = { type: "Identifier", name: token.value };
    } else if (token.type === 33 /* KeywordEmpty */) {
      expr = { type: "Identifier", name: token.value };
    } else if (token.type === 51 /* OperatorLParen */) {
      expr = this.parseExpression();
      if (!this.match(52 /* OperatorRParen */)) {
        throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
      }
    } else {
      throw new Error(`Parse error: Unexpected token in expression '${token.value}' at line ${token.line} `);
    }
    while (true) {
      if (this.match(53 /* OperatorDot */)) {
        const propToken = this.advance();
        const property = { type: "Identifier", name: propToken.value };
        expr = { type: "MemberExpression", object: expr, property };
      } else if (this.match(51 /* OperatorLParen */)) {
        const args = [];
        if (this.peek().type !== 52 /* OperatorRParen */) {
          args.push(this.parseExpression());
          while (this.match(50 /* OperatorComma */)) {
            args.push(this.parseExpression());
          }
        }
        if (!this.match(52 /* OperatorRParen */)) {
          throw new Error(`Parse error: Expected ')' at line ${this.peek().line} `);
        }
        expr = { type: "CallExpression", callee: expr, args };
      } else {
        break;
      }
    }
    return expr;
  }
};

// src/compiler/evaluator.ts
var EmptyVBA = null;
var Environment = class {
  variables = /* @__PURE__ */ new Map();
  procedures = /* @__PURE__ */ new Map();
  enclosing;
  constructor(enclosing) {
    this.enclosing = enclosing;
  }
  set(name, value) {
    const key = name.toLowerCase();
    let env = this;
    while (env) {
      if (env.variables.has(key)) {
        env.variables.set(key, value);
        return;
      }
      env = env.enclosing;
    }
    this.variables.set(key, value);
  }
  setLocally(name, value) {
    this.variables.set(name.toLowerCase(), value);
  }
  get(name) {
    const key = name.toLowerCase();
    let env = this;
    while (env) {
      if (env.variables.has(key)) {
        return env.variables.get(key);
      }
      env = env.enclosing;
    }
    this.variables.set(key, 0);
    return 0;
  }
  setProcedure(name, proc) {
    this.procedures.set(name.toLowerCase(), proc);
  }
  getProcedure(name) {
    const key = name.toLowerCase();
    let env = this;
    while (env) {
      if (env.procedures.has(key)) {
        return env.procedures.get(key);
      }
      env = env.enclosing;
    }
    return void 0;
  }
};
var Evaluator = class {
  env;
  onPrint;
  constructor(onPrint) {
    this.env = new Environment();
    this.onPrint = onPrint;
    this.env.set("debug", {
      print: (...args) => this.onPrint(args.join(" "))
    });
    this.env.set("isempty", (val) => val === void 0 || val === null || val === "" || val === 0);
    this.env.set("isnumeric", (val) => !isNaN(parseFloat(val)) && isFinite(val));
    this.env.set("cdbl", (val) => parseFloat(val) || 0);
    this.env.set("clng", (val) => Math.round(parseFloat(val)) || 0);
    this.env.set("int", (val) => Math.floor(parseFloat(val)) || 0);
    this.env.set("ucase", (val) => String(val || "").toUpperCase());
    this.env.set("trim", (val) => String(val || "").trim());
    this.env.set("ubound", (arr) => {
      if (Array.isArray(arr)) return arr.length - 1;
      return 0;
    });
    this.env.set("createobject", (progId) => {
      if (progId.toLowerCase() === "scripting.dictionary") {
        const dict = /* @__PURE__ */ new Map();
        return {
          add: (k, v) => dict.set(k, v),
          exists: (k) => dict.has(k),
          items: () => Array.from(dict.values()),
          keys: () => Array.from(dict.keys())
        };
      }
      throw new Error(`Execution error: Unsupported CreateObject '${progId}'`);
    });
    this.env.set("true", true);
    this.env.set("false", false);
    this.env.set("empty", EmptyVBA);
  }
  get(name) {
    return this.env.get(name);
  }
  callProcedure(name, args) {
    const procName = name.toLowerCase();
    const proc = this.env.getProcedure(procName);
    if (!proc) {
      throw new Error(`Execution error: Procedure '${name}' not found`);
    }
    const localEnv = new Environment(this.env);
    for (let i = 0; i < proc.parameters.length; i++) {
      const paramName = proc.parameters[i].name;
      const argValue = i < args.length ? args[i] : EmptyVBA;
      localEnv.setLocally(paramName, argValue);
    }
    const previousEnv = this.env;
    this.env = localEnv;
    try {
      for (const stmt of proc.body) {
        this.evaluateStatement(stmt);
      }
    } catch (e) {
      if (e && e.type === "Exit") {
        if (e.target === "Function" && proc.isFunction || e.target === "Sub" && !proc.isFunction) {
        } else {
          throw e;
        }
      } else {
        throw e;
      }
    } finally {
      this.env = previousEnv;
    }
    if (proc.isFunction) {
      return localEnv.get(procName);
    }
    return EmptyVBA;
  }
  evaluate(program) {
    for (const stmt of program.body) {
      this.evaluateStatement(stmt);
    }
  }
  evaluateStatement(stmt) {
    switch (stmt.type) {
      case "ForStatement":
        this.evaluateForStatement(stmt);
        break;
      case "IfStatement":
        this.evaluateIfStatement(stmt);
        break;
      case "DoWhileStatement":
        this.evaluateDoWhileStatement(stmt);
        break;
      case "AssignmentStatement":
        this.evaluateAssignmentStatement(stmt);
        break;
      case "ProcedureDeclaration":
        this.env.setProcedure(stmt.name.name, stmt);
        break;
      case "VariableDeclaration":
        this.evaluateVariableDeclaration(stmt);
        break;
      case "CallStatement":
        this.evaluateCallStatement(stmt);
        break;
      case "ConstDeclaration":
        this.evaluateConstDeclaration(stmt);
        break;
      case "SetStatement":
        this.evaluateSetStatement(stmt);
        break;
      case "OnErrorStatement":
        break;
      case "EraseStatement":
        this.evaluateEraseStatement(stmt);
        break;
      case "ReDimStatement":
        this.evaluateReDimStatement(stmt);
        break;
      case "ExitStatement":
        this.evaluateExitStatement(stmt);
        break;
      case "LabelStatement":
        break;
      default:
        throw new Error(`Execution error: Unknown statement type ${stmt.type}`);
    }
  }
  evaluateForStatement(stmt) {
    let startValue = this.evaluateExpression(stmt.start);
    const endValue = this.evaluateExpression(stmt.end);
    let stepValue = stmt.step ? this.evaluateExpression(stmt.step) : 1;
    const varName = stmt.identifier.name;
    if (this.env.get(varName) === EmptyVBA) {
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
      } catch (e) {
        if (e && e.type === "Exit" && e.target === "For") {
          break;
        }
        throw e;
      }
      this.env.setLocally(varName, this.env.get(varName) + stepValue);
    }
  }
  evaluateIfStatement(stmt) {
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
        this.evaluateIfStatement(stmt.alternate);
      }
    }
  }
  evaluateDoWhileStatement(stmt) {
    while (this.evaluateExpression(stmt.condition)) {
      try {
        for (const bodyStmt of stmt.body) {
          this.evaluateStatement(bodyStmt);
        }
      } catch (e) {
        if (e && e.type === "Exit" && e.target === "Do") {
          break;
        }
        throw e;
      }
    }
  }
  evaluateAssignmentStatement(stmt) {
    const val = this.evaluateExpression(stmt.right);
    if (stmt.left.type === "Identifier") {
      const name = stmt.left.name;
      this.env.set(name, val);
    } else if (stmt.left.type === "CallExpression") {
      const call = stmt.left;
      if (call.callee.type === "Identifier") {
        const name = call.callee.name;
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
    } else if (stmt.left.type === "MemberExpression") {
      const member = stmt.left;
      const obj = this.evaluateExpression(member.object);
      const propName = member.property.name.toLowerCase();
      if (obj && typeof obj === "object") {
        obj[propName] = val;
      } else {
        throw new Error(`Execution error: Cannot assign property '${propName}' of undefined or primitive`);
      }
    } else {
      throw new Error(`Execution error: Invalid assignment target`);
    }
  }
  evaluateVariableDeclaration(stmt) {
    for (const decl of stmt.declarations) {
      let initialValue = EmptyVBA;
      if (decl.isArray) {
        if (decl.arraySize) {
          const size = this.evaluateExpression(decl.arraySize);
          initialValue = new Array(size + 1).fill(EmptyVBA);
        } else {
          initialValue = [];
        }
      } else if (decl.isNew && decl.objectType === "Collection") {
        initialValue = {
          items: [],
          add: function(item) {
            this.items.push(item);
          },
          count: function() {
            return this.items.length;
          },
          item: function(index) {
            return this.items[index - 1];
          }
        };
      }
      this.env.set(decl.name.name, initialValue);
    }
  }
  evaluateCallStatement(stmt) {
    this.evaluateExpression(stmt.expression);
  }
  evaluateConstDeclaration(stmt) {
    const value = this.evaluateExpression(stmt.value);
    this.env.set(stmt.name.name, value);
  }
  evaluateSetStatement(stmt) {
    const value = this.evaluateExpression(stmt.right);
    if (stmt.left.type === "Identifier") {
      this.env.set(stmt.left.name, value);
    } else {
      throw new Error(`Execution error: Unsupported Set target ${stmt.left.type}`);
    }
  }
  evaluateEraseStatement(stmt) {
    this.env.set(stmt.name.name, []);
  }
  evaluateReDimStatement(stmt) {
    if (stmt.bounds.length > 0) {
      const size = this.evaluateExpression(stmt.bounds[stmt.bounds.length - 1]);
      const arr = new Array(size + 1).fill(EmptyVBA);
      this.env.set(stmt.name.name, arr);
    }
  }
  evaluateExitStatement(stmt) {
    throw { type: "Exit", target: stmt.exitType };
  }
  evaluateExpression(expr) {
    switch (expr.type) {
      case "NumberLiteral":
        return expr.value;
      case "StringLiteral":
        return expr.value;
      case "Identifier":
        return this.env.get(expr.name);
      case "CallExpression":
        return this.evaluateCallExpression(expr);
      case "MemberExpression":
        return this.evaluateMemberExpression(expr);
      case "UnaryExpression":
        return this.evaluateUnaryExpression(expr);
      case "BinaryExpression":
        return this.evaluateBinaryExpression(expr);
      default:
        throw new Error(`Execution error: Unknown expression type ${expr.type}`);
    }
  }
  evaluateCallExpression(expr) {
    if (expr.callee.type === "Identifier") {
      const name = expr.callee.name;
      const proc = this.env.getProcedure(name);
      if (proc) {
        const localEnv = new Environment(this.env);
        for (let i = 0; i < proc.parameters.length; i++) {
          const argVal = i < expr.args.length ? this.evaluateExpression(expr.args[i]) : 0;
          localEnv.set(proc.parameters[i].name, argVal);
        }
        if (proc.isFunction) {
          localEnv.setLocally(proc.name.name, EmptyVBA);
        }
        const previousEnv = this.env;
        this.env = localEnv;
        try {
          for (const s of proc.body) {
            this.evaluateStatement(s);
          }
        } catch (e) {
          if (e && e.type === "Exit" && (e.target === "Sub" || e.target === "Function")) {
          } else {
            throw e;
          }
        }
        this.env = previousEnv;
        if (proc.isFunction) {
          return localEnv.get(proc.name.name);
        }
        return void 0;
      } else {
        const variable = this.env.get(name);
        if (typeof variable === "function") {
          const argsVals = expr.args.map((a) => this.evaluateExpression(a));
          return variable(...argsVals);
        } else if (Array.isArray(variable)) {
          if (expr.args.length === 0) throw new Error(`Execution error: Missing index for array ${name}`);
          const idx = this.evaluateExpression(expr.args[0]);
          return variable[idx];
        }
        throw new Error(`Execution error: Cannot call unknown procedure or index unknown array '${name}'`);
      }
    } else if (expr.callee.type === "MemberExpression") {
      const member = expr.callee;
      const obj = this.evaluateExpression(member.object);
      const methodNameLower = member.property.name.toLowerCase();
      const methodNameOriginal = member.property.name;
      if (obj) {
        let targetMethod = obj[methodNameLower];
        if (typeof targetMethod !== "function") {
          const keys = Object.keys(obj);
          if (typeof obj[methodNameOriginal] === "function") {
            targetMethod = obj[methodNameOriginal];
          } else {
            const match = keys.find((k) => k.toLowerCase() === methodNameLower);
            if (match) targetMethod = obj[match];
          }
        }
        if (typeof targetMethod === "function") {
          const argsVals = expr.args.map((a) => this.evaluateExpression(a));
          return targetMethod.apply(obj, argsVals);
        }
      }
      throw new Error(`Execution error: Object does not support property or method '${methodNameOriginal}'`);
    }
    throw new Error(`Execution error: Unsupported call expression`);
  }
  evaluateUnaryExpression(expr) {
    const argument = this.evaluateExpression(expr.argument);
    switch (expr.operator.toLowerCase()) {
      case "not":
        return !argument;
      case "-":
        return -argument;
      case "+":
        return +argument;
      default:
        throw new Error(`Execution error: Unknown unary operator ${expr.operator}`);
    }
  }
  evaluateMemberExpression(expr) {
    const obj = this.evaluateExpression(expr.object);
    const propName = expr.property.name.toLowerCase();
    if (obj && typeof obj[propName] === "function") {
      return obj[propName]();
    } else if (obj && propName in obj) {
      return obj[propName];
    }
    throw new Error(`Execution error: Method or property not found '${propName}'`);
  }
  evaluateBinaryExpression(expr) {
    const leftVal = this.evaluateExpression(expr.left);
    const rightVal = this.evaluateExpression(expr.right);
    switch (expr.operator.toLowerCase()) {
      case "+":
        return leftVal + rightVal;
      case "-":
        return leftVal - rightVal;
      case "*":
        return leftVal * rightVal;
      case "/":
        return leftVal / rightVal;
      case "\\":
        return Math.floor(leftVal / rightVal);
      case "mod":
        return leftVal % rightVal;
      case "^":
        return Math.pow(leftVal, rightVal);
      case "=":
        return leftVal === rightVal;
      case "<>":
        return leftVal !== rightVal;
      case "<":
        return leftVal < rightVal;
      case ">":
        return leftVal > rightVal;
      case "<=":
        return leftVal <= rightVal;
      case ">=":
        return leftVal >= rightVal;
      case "and":
        return leftVal && rightVal;
      case "or":
        return leftVal || rightVal;
      default:
        throw new Error(`Execution error: Unknown operator ${expr.operator}`);
    }
  }
};

// tests/ts/test-runner.ts
function runVBATest(filePath, procedureName, args) {
  const start = Date.now();
  const sourceCode = fs.readFileSync(filePath, "utf-8");
  const lexer = new Lexer(sourceCode);
  const tokens = lexer.tokenize();
  const parser = new Parser(tokens);
  const ast = parser.parse();
  const evaluator = new Evaluator();
  evaluator.evaluate(ast);
  const result = evaluator.callProcedure(procedureName, args);
  const duration = Date.now() - start;
  const formatArgs = args.map((a) => typeof a === "object" ? "[Object]" : String(a)).join(", ");
  console.log(`[PASS] ${procedureName}(${formatArgs}) -> ${result} (${duration}ms)`);
  return result;
}
var assert = {
  strictEqual: (actual, expected, message) => {
    if (actual !== expected) {
      console.error(`[FAIL] ${message || "Assertion failed"} - Expected ${expected} but got ${actual}`);
      throw new Error(`Assertion Failed`);
    }
  }
};

// tests/ts/run.ts
async function main() {
  console.log("--- Starting VBA Unit Tests ---");
  const vbaFile = "tests/vba/test_refactored.vba";
  console.log("\n[Test Suite] CalcBaseStartIdx");
  assert.strictEqual(runVBATest(vbaFile, "CalcBaseStartIdx", [1, 10, 1]), 1, "Level 1 ignores parent finish");
  assert.strictEqual(runVBATest(vbaFile, "CalcBaseStartIdx", [2, 10, 0.49]), 10, "Alloc < 0.5 starts same day");
  assert.strictEqual(runVBATest(vbaFile, "CalcBaseStartIdx", [3, 5, 0]), 5, "Alloc 0 starts same day");
  assert.strictEqual(runVBATest(vbaFile, "CalcBaseStartIdx", [2, 10, 0.5]), 11, "Alloc >= 0.5 starts next day");
  assert.strictEqual(runVBATest(vbaFile, "CalcBaseStartIdx", [4, 20, 1]), 21, "Alloc 1.0 starts next day");
  console.log("\n[Test Suite] GetMaxDailyLoad");
  const mockDict = /* @__PURE__ */ new Map();
  mockDict.set("Alice", 0.5);
  mockDict.set("Bob", 0.8);
  const capacityLimits = function(key) {
    return mockDict.get(key);
  };
  capacityLimits.Exists = function(key) {
    return mockDict.has(key);
  };
  assert.strictEqual(runVBATest(vbaFile, "GetMaxDailyLoad", ["Charlie", capacityLimits]), 1, "Missing user defaults to 1.0");
  assert.strictEqual(runVBATest(vbaFile, "GetMaxDailyLoad", ["Alice", capacityLimits]), 0.5, "Known user returns configured capacity");
  assert.strictEqual(runVBATest(vbaFile, "GetMaxDailyLoad", ["Bob", capacityLimits]), 0.8, "Known user returns configured capacity");
  console.log("\n[Test Suite] CalcDailyAllocation");
  assert.strictEqual(runVBATest(vbaFile, "CalcDailyAllocation", [1, 0.5, false]), 0.5, "Standard: exact fit");
  assert.strictEqual(runVBATest(vbaFile, "CalcDailyAllocation", [0.25, 0.5, false]), 0.25, "Standard: limited by capacity");
  assert.strictEqual(runVBATest(vbaFile, "CalcDailyAllocation", [1, 0.1, false]), 0, "Standard: rounds to 0 if < 0.125 needed");
  assert.strictEqual(runVBATest(vbaFile, "CalcDailyAllocation", [1, 0.1, true]), 0.1, "Micro: capacity 1.0 > 0.1, gets 0.1");
  assert.strictEqual(runVBATest(vbaFile, "CalcDailyAllocation", [0.05, 0.1, true]), 0, "Micro: capacity < 0.1, skips");
  console.log("\n[Test Suite] UpdateLevelFinish");
  const levelFinish = [0, 10, 5];
  const levelAlloc = [0, 0.5, 1];
  runVBATest(vbaFile, "UpdateLevelFinish", [1, 15, 0.8, levelFinish, levelAlloc]);
  assert.strictEqual(levelFinish[1], 15, "Level 1 finish index updated due to > 10");
  assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc updated due to index change");
  runVBATest(vbaFile, "UpdateLevelFinish", [1, 15, 0.1, levelFinish, levelAlloc]);
  assert.strictEqual(levelAlloc[1], 0.8, "Level 1 alloc unchanged, 0.1 not > 0.8");
  runVBATest(vbaFile, "UpdateLevelFinish", [1, 15, 1, levelFinish, levelAlloc]);
  assert.strictEqual(levelAlloc[1], 1, "Level 1 alloc updated, 1.0 > 0.8");
  console.log("\n--- All tests passed! ---");
}
main().catch(console.error);
