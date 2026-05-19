import { Statement } from '../engine/parser';
import { buildSymbolTable, getWordAtPosition } from './symbol-table';

export interface LocationInfo {
    uri: string;
    range: {
        start: { line: number; character: number };
        end: { line: number; character: number };
    };
}

export class ReferencesProvider {
    private uri: string = '';

    setDocumentUri(uri: string): void {
        this.uri = uri;
    }

    getReferences(
        statements: Statement[],
        sourceText: string,
        line: number,
        character: number,
        includeDeclaration: boolean,
    ): LocationInfo[] {
        const word = getWordAtPosition(sourceText, line, character);
        if (!word) return [];
        return findAllReferences(sourceText, word, this.uri, statements, includeDeclaration);
    }
}

export function findAllReferences(
    sourceText: string,
    targetWord: string,
    uri: string,
    statements: Statement[],
    includeDeclaration: boolean,
): LocationInfo[] {
    const refs: LocationInfo[] = [];
    const targetLower = targetWord.toLowerCase();

    const symbols = buildSymbolTable(statements);
    const declEntry = symbols.get(targetLower);

    const lines = sourceText.split('\n');
    const pattern = new RegExp(`(?<![a-zA-Z0-9_])${escapeRegex(targetWord)}(?![a-zA-Z0-9_])`, 'gi');

    for (let lineIdx = 0; lineIdx < lines.length; lineIdx++) {
        const lineText = lines[lineIdx];

        // Skip pure comment lines (lines whose first non-whitespace char is ')
        const trimmed = lineText.trimStart();
        if (trimmed.startsWith("'") || trimmed.toLowerCase().startsWith('rem ')) continue;

        let match: RegExpExecArray | null;
        pattern.lastIndex = 0;

        while ((match = pattern.exec(lineText)) !== null) {
            const charIdx = match.index;

            // Skip occurrence if it falls inside a string literal
            if (isInsideString(lineText, charIdx)) continue;

            // Skip the part after ' comment marker on the same line
            if (isAfterComment(lineText, charIdx)) continue;

            const range = {
                start: { line: lineIdx, character: charIdx },
                end: { line: lineIdx, character: charIdx + targetWord.length },
            };

            // Optionally exclude the declaration site
            if (!includeDeclaration && declEntry) {
                const d = declEntry.range.start;
                if (lineIdx === d.line && charIdx === d.character) continue;
            }

            refs.push({ uri, range });
        }
    }

    return refs;
}

function escapeRegex(s: string): string {
    return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function isInsideString(lineText: string, pos: number): boolean {
    let inString = false;
    for (let i = 0; i < pos; i++) {
        if (lineText[i] === '"') {
            // Check for escaped quote ("")
            if (i + 1 < lineText.length && lineText[i + 1] === '"') {
                i++; // skip escaped quote
            } else {
                inString = !inString;
            }
        }
    }
    return inString;
}

function isAfterComment(lineText: string, pos: number): boolean {
    let inString = false;
    for (let i = 0; i < pos; i++) {
        if (lineText[i] === '"') {
            if (i + 1 < lineText.length && lineText[i + 1] === '"') {
                i++;
            } else {
                inString = !inString;
            }
        } else if (lineText[i] === "'" && !inString) {
            return true;
        }
    }
    return false;
}
