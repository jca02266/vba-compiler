import * as path from 'path';

export interface FileSystem {
    existsSync(path: string): boolean;
    readFileSync(path: string, encoding: 'utf-8' | 'utf8'): string;
    writeFileSync(path: string, content: string): void;
    appendFileSync?(path: string, content: string): void;
    mkdirSync(path: string, options?: { recursive?: boolean }): void;
    rmdirSync?(path: string): void;
    unlinkSync(path: string): void;
    readdirSync(path: string): string[];
    statSync(path: string): {
        size: number;
        isFile(): boolean;
        isDirectory(): boolean;
        mtime: Date;
    };
    openSync(path: string, flags: string): number;
    closeSync(fd: number): void;
    readSync(fd: number, buffer: Uint8Array, offset: number, length: number, position: number | null): number;
    writeSync(fd: number, buffer: string | Uint8Array, offset?: number, length?: number, position?: number | null): number;
}

/**
 * Memory-based file system for browser environments.
 */
export class MemoryFileSystem implements FileSystem {
    private files: Map<string, Uint8Array | string> = new Map();
    private dirs: Set<string> = new Set(['/']);
    private fileHandles: Map<number, { path: string, flags: string, pos: number }> = new Map();
    private nextFd = 1;

    existsSync(p: string): boolean {
        const norm = this.normalize(p);
        return this.files.has(norm) || this.dirs.has(norm);
    }

    readFileSync(p: string, encoding: 'utf-8' | 'utf8'): string {
        const norm = this.normalize(p);
        const data = this.files.get(norm);
        if (data === undefined) throw new Error(`File not found: ${p}`);
        if (typeof data === 'string') return data;
        return new TextDecoder().decode(data);
    }

    writeFileSync(p: string, content: string | Uint8Array): void {
        const norm = this.normalize(p);
        this.ensureDir(path.dirname(norm));
        this.files.set(norm, content);
    }

    mkdirSync(p: string, options?: { recursive?: boolean }): void {
        const norm = this.normalize(p);
        if (options?.recursive) {
            const parts = norm.split('/').filter(Boolean);
            let current = '';
            for (const part of parts) {
                current += '/' + part;
                this.dirs.add(current);
            }
        } else {
            this.dirs.add(norm);
        }
    }

    unlinkSync(p: string): void {
        const norm = this.normalize(p);
        if (!this.files.has(norm)) {
            throw new Error(`ENOENT: no such file or directory, unlink '${p}'`);
        }
        this.files.delete(norm);
    }

    readdirSync(p: string): string[] {
        const norm = this.normalize(p);
        const prefix = norm === '/' ? '/' : norm + '/';
        const results = new Set<string>();
        for (const f of this.files.keys()) {
            if (f.startsWith(prefix)) {
                const rest = f.substring(prefix.length);
                const nextSlash = rest.indexOf('/');
                results.add(nextSlash === -1 ? rest : rest.substring(0, nextSlash));
            }
        }
        for (const d of this.dirs) {
            if (d !== norm && d.startsWith(prefix)) {
                const rest = d.substring(prefix.length);
                const nextSlash = rest.indexOf('/');
                results.add(nextSlash === -1 ? rest : rest.substring(0, nextSlash));
            }
        }
        return Array.from(results);
    }

    statSync(p: string) {
        const norm = this.normalize(p);
        const isFile = this.files.has(norm);
        const isDir = this.dirs.has(norm);
        if (!isFile && !isDir) throw new Error(`Not found: ${p}`);
        
        const data = this.files.get(norm);
        const size = typeof data === 'string' ? new TextEncoder().encode(data).length : (data?.length || 0);

        return {
            size,
            isFile: () => isFile,
            isDirectory: () => isDir,
            mtime: new Date()
        };
    }

    openSync(p: string, flags: string): number {
        const norm = this.normalize(p);
        if (flags === 'r' && !this.existsSync(norm)) throw new Error(`File not found: ${p}`);
        if (flags === 'w' || flags === 'a') {
            if (!this.files.has(norm)) this.writeFileSync(norm, "");
        }
        const fd = this.nextFd++;
        this.fileHandles.set(fd, { path: norm, flags, pos: 0 });
        return fd;
    }

    closeSync(fd: number): void {
        this.fileHandles.delete(fd);
    }

    readSync(fd: number, buffer: Uint8Array, offset: number, length: number, position: number | null): number {
        const h = this.fileHandles.get(fd);
        if (!h) throw new Error("Invalid FD");
        const data = this.files.get(h.path);
        if (!data) return 0;
        
        const bin = typeof data === 'string' ? new TextEncoder().encode(data) : data;
        const start = position !== null ? position : h.pos;
        const available = bin.length - start;
        const toRead = Math.min(length, available);
        
        if (toRead <= 0) return 0;
        buffer.set(bin.subarray(start, start + toRead), offset);
        if (position === null) h.pos += toRead;
        return toRead;
    }

    writeSync(fd: number, buffer: string | Uint8Array, offset?: number, length?: number, position?: number | null): number {
        const h = this.fileHandles.get(fd);
        if (!h) throw new Error("Invalid FD");
        
        const newData = typeof buffer === 'string' ? new TextEncoder().encode(buffer) : buffer.subarray(offset || 0, (offset || 0) + (length || buffer.length));
        const oldData = this.files.get(h.path) || new Uint8Array(0);
        const oldBin = typeof oldData === 'string' ? new TextEncoder().encode(oldData) : oldData;
        
        const start = position !== null ? position : h.pos;
        const totalSize = Math.max(oldBin.length, start + newData.length);
        const combined = new Uint8Array(totalSize);
        combined.set(oldBin);
        combined.set(newData, start);
        
        this.files.set(h.path, combined);
        if (position === null) h.pos += newData.length;
        return newData.length;
    }

    private normalize(p: string): string {
        return p.replace(/\\/g, '/').replace(/\/+$/, '') || '/';
    }

    private ensureDir(p: string) {
        this.mkdirSync(p, { recursive: true });
    }
}
