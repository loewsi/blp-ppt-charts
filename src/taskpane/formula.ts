// A tiny, dependency-free spreadsheet formula engine for the datasheet.
// Cells hold a raw string: a literal ("12", "Q1") or a formula ("=B2+B3", "=SUM(B2:D2)").
// Addressing is A1-style over the WHOLE grid including headers: column letters map to
// grid columns (A=col 0), row numbers to grid rows (1 = row 0). So the first data value
// (grid row 1, col 1) is B2. Supports + - * /, parentheses, unary minus, and the
// functions SUM/AVERAGE/AVG/MIN/MAX/COUNT over cells and ranges. MIT-clean, no deps.

export type Grid = string[][];

const ERR = "#ERR";
const CYCLE = "#CYCLE";
const REF = "#REF";

/** "A"→0, "B"→1, … "AA"→26. */
export function colToIndex(letters: string): number {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

/** Parse an A1 reference to grid coordinates (0-based). */
function refToRC(ref: string): { r: number; c: number } | null {
  const m = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!m) return null;
  return { c: colToIndex(m[1]), r: parseInt(m[2], 10) - 1 };
}

type Tok =
  | { t: "num"; v: number }
  | { t: "ref"; v: string }
  | { t: "range"; a: string; b: string }
  | { t: "func"; v: string }
  | { t: "op"; v: string };

function tokenize(src: string): Tok[] {
  const toks: Tok[] = [];
  let i = 0;
  const s = src.toUpperCase();
  while (i < s.length) {
    const ch = s[i];
    if (ch === " ") {
      i++;
      continue;
    }
    if ("+-*/(),".includes(ch)) {
      toks.push({ t: "op", v: ch });
      i++;
      continue;
    }
    if (/[0-9.]/.test(ch)) {
      let j = i + 1;
      while (j < s.length && /[0-9.]/.test(s[j])) j++;
      toks.push({ t: "num", v: parseFloat(s.slice(i, j)) });
      i = j;
      continue;
    }
    if (/[A-Z]/.test(ch)) {
      let j = i + 1;
      while (j < s.length && /[A-Z0-9]/.test(s[j])) j++;
      let word = s.slice(i, j);
      i = j;
      // Range? A1:B3
      if (s[i] === ":" && /^[A-Z]+\d+$/.test(word)) {
        let k = i + 1;
        while (k < s.length && /[A-Z0-9]/.test(s[k])) k++;
        const b = s.slice(i + 1, k);
        toks.push({ t: "range", a: word, b });
        i = k;
        continue;
      }
      // Function name (letters immediately followed by "(")?
      if (s[i] === "(" && /^[A-Z]+$/.test(word)) {
        toks.push({ t: "func", v: word });
        continue;
      }
      if (/^[A-Z]+\d+$/.test(word)) {
        toks.push({ t: "ref", v: word });
        continue;
      }
      // Bare word that isn't a ref/func (e.g. a stray name) → treat as an error token.
      throw new Error(REF);
    }
    throw new Error(ERR);
  }
  return toks;
}

/** Evaluate a formula grid → display strings and numeric values (parallel to `grid`). */
export function evaluateGrid(grid: Grid): { display: string[][]; values: number[][] } {
  const R = grid.length;
  const display: string[][] = grid.map((row) => row.map((v) => v ?? ""));
  const values: number[][] = grid.map((row) => row.map(() => 0));
  const memo = new Map<string, number>();

  const cellRaw = (r: number, c: number) => (r >= 0 && r < R && c >= 0 && c < (grid[r]?.length ?? 0) ? grid[r][c] ?? "" : "");

  // Numeric value of a cell, resolving a formula recursively (with cycle guard).
  function cellValue(r: number, c: number, seen: Set<string>): number {
    const key = `${r},${c}`;
    if (memo.has(key)) return memo.get(key)!;
    if (seen.has(key)) throw new Error(CYCLE);
    const raw = cellRaw(r, c).trim();
    let out: number;
    if (raw.startsWith("=")) {
      seen.add(key);
      out = evalFormula(raw.slice(1), seen);
      seen.delete(key);
    } else {
      const n = Number(raw.replace(/[,\s%]/g, ""));
      out = raw !== "" && isFinite(n) ? n : 0; // text → 0 for arithmetic
    }
    memo.set(key, out);
    return out;
  }

  function rangeValues(a: string, b: string, seen: Set<string>): number[] {
    const ra = refToRC(a);
    const rb = refToRC(b);
    if (!ra || !rb) throw new Error(REF);
    const out: number[] = [];
    for (let r = Math.min(ra.r, rb.r); r <= Math.max(ra.r, rb.r); r++) {
      for (let c = Math.min(ra.c, rb.c); c <= Math.max(ra.c, rb.c); c++) {
        const raw = cellRaw(r, c).trim();
        if (raw === "") continue; // blank cells don't count
        // For aggregation, skip pure text (non-numeric, non-formula).
        if (!raw.startsWith("=") && !isFinite(Number(raw.replace(/[,\s%]/g, "")))) continue;
        out.push(cellValue(r, c, seen));
      }
    }
    return out;
  }

  function evalFormula(expr: string, seen: Set<string>): number {
    const toks = tokenize(expr);
    let p = 0;
    const peek = () => toks[p];
    const next = () => toks[p++];

    const opValue = (tk: Tok | undefined): string | undefined => (tk && tk.t === "op" ? tk.v : undefined);

    function parseExpr(): number {
      let v = parseTerm();
      for (;;) {
        const op = opValue(peek());
        if (op !== "+" && op !== "-") break;
        next();
        const rhs = parseTerm();
        v = op === "+" ? v + rhs : v - rhs;
      }
      return v;
    }
    function parseTerm(): number {
      let v = parseFactor();
      for (;;) {
        const op = opValue(peek());
        if (op !== "*" && op !== "/") break;
        next();
        const rhs = parseFactor();
        v = op === "*" ? v * rhs : v / rhs;
      }
      return v;
    }
    function parseFactor(): number {
      const tk = peek();
      if (!tk) throw new Error(ERR);
      if (tk.t === "op" && tk.v === "-") {
        next();
        return -parseFactor();
      }
      if (tk.t === "op" && tk.v === "(") {
        next();
        const v = parseExpr();
        if (opValue(peek()) !== ")") throw new Error(ERR);
        next();
        return v;
      }
      if (tk.t === "num") {
        next();
        return tk.v;
      }
      if (tk.t === "ref") {
        next();
        const rc = refToRC(tk.v)!;
        return cellValue(rc.r, rc.c, seen);
      }
      if (tk.t === "func") {
        next();
        return parseFunc(tk.v);
      }
      throw new Error(ERR);
    }
    function parseFunc(name: string): number {
      if (opValue(peek()) !== "(") throw new Error(ERR);
      next(); // "("
      const args: number[] = [];
      if (opValue(peek()) !== ")") {
        for (;;) {
          const tk = peek();
          if (tk && tk.t === "range") {
            next();
            args.push(...rangeValues(tk.a, tk.b, seen));
          } else {
            args.push(parseExpr());
          }
          if (opValue(peek()) === ",") {
            next();
            continue;
          }
          break;
        }
      }
      if (opValue(peek()) !== ")") throw new Error(ERR);
      next();
      return applyFunc(name, args);
    }

    const v = parseExpr();
    if (p !== toks.length) throw new Error(ERR); // trailing junk
    return v;
  }

  for (let r = 0; r < R; r++) {
    for (let c = 0; c < (grid[r]?.length ?? 0); c++) {
      const raw = (grid[r][c] ?? "").trim();
      if (!raw.startsWith("=")) {
        const n = Number(raw.replace(/[,\s%]/g, ""));
        values[r][c] = raw !== "" && isFinite(n) ? n : 0;
        continue;
      }
      try {
        const v = cellValue(r, c, new Set());
        values[r][c] = v;
        display[r][c] = isFinite(v) ? String(round(v)) : ERR;
      } catch (e) {
        const msg = e instanceof Error ? e.message : ERR;
        display[r][c] = [CYCLE, REF, ERR].includes(msg) ? msg : ERR;
        values[r][c] = 0;
      }
    }
  }
  return { display, values };
}

function applyFunc(name: string, args: number[]): number {
  switch (name) {
    case "SUM":
      return args.reduce((a, b) => a + b, 0);
    case "AVERAGE":
    case "AVG":
      return args.length ? args.reduce((a, b) => a + b, 0) / args.length : 0;
    case "MIN":
      return args.length ? Math.min(...args) : 0;
    case "MAX":
      return args.length ? Math.max(...args) : 0;
    case "COUNT":
      return args.length;
    case "ABS":
      return Math.abs(args[0] ?? 0);
    case "ROUND":
      return Number((args[0] ?? 0).toFixed(Math.max(0, Math.round(args[1] ?? 0))));
    default:
      throw new Error(ERR);
  }
}

function round(v: number): number {
  return Math.round(v * 1e6) / 1e6; // kill FP noise in the display
}
