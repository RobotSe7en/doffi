/** First-party Word DATE/TIME picture formatting and Word table formula evaluation (=SUM(ABOVE), etc.). */

export function formatWordDatePicture(date: Date, instruction: string, isTime: boolean): string {
  const picMatch = /\\@\s*(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const picture =
    picMatch?.[1] ??
    picMatch?.[2] ??
    picMatch?.[3] ??
    (isTime ? "HH:mm:ss" : "yyyy/M/d");
  const yyyy = String(date.getFullYear());
  const yy = yyyy.slice(-2);
  const M = String(date.getMonth() + 1);
  const MM = M.padStart(2, "0");
  const d = String(date.getDate());
  const dd = d.padStart(2, "0");
  const H = String(date.getHours());
  const HH = H.padStart(2, "0");
  const hNum = date.getHours() % 12 || 12;
  const h = String(hNum);
  const hh = h.padStart(2, "0");
  const m = String(date.getMinutes());
  const mm = m.padStart(2, "0");
  const s = String(date.getSeconds());
  const ss = s.padStart(2, "0");
  const ampm = date.getHours() < 12 ? "AM" : "PM";
  const monthsLong = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
  ];
  const monthsShort = monthsLong.map((name) => name.slice(0, 3));
  return picture.replace(
    /yyyy|yy|MMMM|MMM|MM|M|dd|d|HH|H|hh|h|mm|m|ss|s|AM\/PM|am\/pm/g,
    (token) => {
      switch (token) {
        case "yyyy": return yyyy;
        case "yy": return yy;
        case "MMMM": return monthsLong[date.getMonth()]!;
        case "MMM": return monthsShort[date.getMonth()]!;
        case "MM": return MM;
        case "M": return M;
        case "dd": return dd;
        case "d": return d;
        case "HH": return HH;
        case "H": return H;
        case "hh": return hh;
        case "h": return h;
        case "mm": return mm;
        case "m": return m;
        case "ss": return ss;
        case "s": return s;
        case "AM/PM": return ampm;
        case "am/pm": return ampm.toLowerCase();
        default: return token;
      }
    },
  );
}

function parseCellNumber(cell: HTMLElement | undefined): number | undefined {
  if (!cell) return undefined;
  const raw = (cell.textContent ?? "").replace(/\uFFFC/g, "").trim();
  if (!raw) return undefined;
  const cleaned = raw.replace(/[,，\s¥$€£]/g, "");
  if (cleaned.endsWith("%")) {
    const n = Number(cleaned.slice(0, -1));
    return Number.isFinite(n) ? n / 100 : undefined;
  }
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : undefined;
}

function buildTableMatrix(table: HTMLTableElement): (HTMLElement | undefined)[][] {
  const rows = [...table.rows];
  const grid: (HTMLElement | undefined)[][] = Array.from({ length: rows.length }, () => []);
  for (let r = 0; r < rows.length; r++) {
    let c = 0;
    for (const cell of [...rows[r]!.cells]) {
      while (grid[r]![c]) c++;
      const rs = Math.max(1, Number(cell.getAttribute("rowspan") ?? "1") || 1);
      const cs = Math.max(1, Number(cell.getAttribute("colspan") ?? "1") || 1);
      for (let dy = 0; dy < rs && r + dy < rows.length; dy++)
        for (let dx = 0; dx < cs; dx++) grid[r + dy]![c + dx] = cell;
      c += cs;
    }
  }
  return grid;
}

function parseA1(ref: string): { row: number; col: number } | undefined {
  const m = /^([A-Z]{1,3})(\d+)$/i.exec(ref.trim());
  if (!m) return undefined;
  let col = 0;
  for (const ch of m[1]!.toUpperCase()) col = col * 26 + (ch.charCodeAt(0) - 64);
  const row = Number(m[2]) - 1;
  return row >= 0 && col >= 1 ? { row, col: col - 1 } : undefined;
}

function formatNumberPicture(value: number, picture?: string): string {
  if (!Number.isFinite(value)) return "!错误";
  if (!picture) {
    const rounded = Math.round(value * 1e10) / 1e10;
    return String(rounded);
  }
  const clean = picture.trim().replace(/^["']|["']$/g, "");
  const isPercent = clean.includes("%");
  const scaled = isPercent ? value * 100 : value;
  const dot = clean.indexOf(".");
  let decimals = 0;
  if (dot >= 0) {
    const frac = clean.slice(dot + 1).match(/^[0#]+/)?.[0] ?? "";
    decimals = frac.length;
  }
  const hasComma = clean.includes(",");
  const prefix = clean.match(/^[^\d#.,%-]+/)?.[0] ?? "";
  const suffix = clean.match(/[^\d#.,%]+$/)?.[0] ?? "";
  const abs = Math.abs(scaled);
  const fixed = abs.toFixed(decimals);
  const [intPart, fracPart] = fixed.split(".");
  const grouped = hasComma ? intPart!.replace(/\B(?=(\d{3})+(?!\d))/g, ",") : intPart!;
  const sign = scaled < 0 ? "-" : "";
  return `${sign}${prefix}${grouped}${fracPart !== undefined && decimals > 0 ? "." + fracPart : ""}${isPercent ? "%" : ""}${suffix}`;
}

/** Evaluate a Word table formula field (`=SUM(ABOVE)`, `=AVERAGE(LEFT)`, `=A1+B2`, etc.). */
export function evaluateWordTableFormula(instruction: string, fieldElement: HTMLElement): string {
  const switchMatch = /\\#\s*(?:"([^"]+)"|'([^']+)'|(\S+))/i.exec(instruction);
  const picture = switchMatch?.[1] ?? switchMatch?.[2] ?? switchMatch?.[3];
  const expr = instruction
    .replace(/\\#\s*(?:"[^"]*"|'[^']*'|\S+)/gi, "")
    .replace(/\\\*\s*\S+/gi, "")
    .trim()
    .replace(/^=\s*/, "");
  if (!expr) return "0";

  const cell = fieldElement.closest<HTMLElement>("td,th");
  const table = cell?.closest<HTMLTableElement>("table");
  const matrix = table ? buildTableMatrix(table) : [];
  let selfRow = -1,
    selfCol = -1;
  if (cell && matrix.length) {
    outer: for (let r = 0; r < matrix.length; r++) {
      for (let c = 0; c < matrix[r]!.length; c++) {
        if (matrix[r]![c] === cell) {
          selfRow = r;
          selfCol = c;
          break outer;
        }
      }
    }
  }

  const collectDirectional = (dir: string): number[] => {
    if (selfRow < 0 || selfCol < 0) return [];
    const values: number[] = [];
    const seen = new Set<HTMLElement>();
    const step = (r: number, c: number): boolean => {
      const target = matrix[r]?.[c];
      if (!target || target === cell || seen.has(target)) return true;
      seen.add(target);
      const n = parseCellNumber(target);
      if (n === undefined) return values.length === 0;
      values.push(n);
      return true;
    };
    if (dir === "ABOVE") {
      for (let r = selfRow - 1; r >= 0; r--) if (!step(r, selfCol)) break;
    } else if (dir === "BELOW") {
      for (let r = selfRow + 1; r < matrix.length; r++) if (!step(r, selfCol)) break;
    } else if (dir === "LEFT") {
      for (let c = selfCol - 1; c >= 0; c--) if (!step(selfRow, c)) break;
    } else if (dir === "RIGHT") {
      for (let c = selfCol + 1; c < (matrix[selfRow]?.length ?? 0); c++)
        if (!step(selfRow, c)) break;
    }
    return values;
  };

  const collectRange = (a: string, b: string): number[] => {
    const start = parseA1(a),
      end = parseA1(b);
    if (!start || !end) return [];
    const r0 = Math.min(start.row, end.row),
      r1 = Math.max(start.row, end.row);
    const c0 = Math.min(start.col, end.col),
      c1 = Math.max(start.col, end.col);
    const values: number[] = [];
    const seen = new Set<HTMLElement>();
    for (let r = r0; r <= r1; r++) {
      for (let c = c0; c <= c1; c++) {
        const target = matrix[r]?.[c];
        if (!target || target === cell || seen.has(target)) continue;
        seen.add(target);
        const n = parseCellNumber(target);
        if (n !== undefined) values.push(n);
      }
    }
    return values;
  };

  const evalArgs = (raw: string): number[] => {
    const parts: string[] = [];
    let depth = 0,
      start = 0;
    for (let i = 0; i < raw.length; i++) {
      if (raw[i] === "(") depth++;
      else if (raw[i] === ")") depth--;
      else if (raw[i] === "," && depth === 0) {
        parts.push(raw.slice(start, i).trim());
        start = i + 1;
      }
    }
    parts.push(raw.slice(start).trim());
    const out: number[] = [];
    for (const part of parts) {
      if (!part) continue;
      const upper = part.toUpperCase();
      if (["ABOVE", "BELOW", "LEFT", "RIGHT"].includes(upper)) {
        out.push(...collectDirectional(upper));
      } else if (/^[A-Z]{1,3}\d+\s*:\s*[A-Z]{1,3}\d+$/i.test(part)) {
        const [a, b] = part.split(":");
        out.push(...collectRange(a!, b!));
      } else {
        out.push(evalExpression(part));
      }
    }
    return out;
  };

  const evalFunc = (name: string, rawArgs: string): number => {
    const fn = name.toUpperCase();
    const nums = evalArgs(rawArgs);
    switch (fn) {
      case "SUM":
        return nums.reduce((a, b) => a + b, 0);
      case "AVERAGE":
        return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0;
      case "COUNT":
        return nums.length;
      case "MIN":
        return nums.length ? Math.min(...nums) : 0;
      case "MAX":
        return nums.length ? Math.max(...nums) : 0;
      case "PRODUCT":
        return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
      case "ABS":
        return Math.abs(nums[0] ?? 0);
      case "INT":
        return Math.trunc(nums[0] ?? 0);
      case "ROUND": {
        const factor = 10 ** (nums[1] ?? 0);
        return Math.round((nums[0] ?? 0) * factor) / factor;
      }
      case "MOD":
        return nums[1] ? (nums[0] ?? 0) % nums[1] : Number.NaN;
      case "IF":
        return nums[0] ? (nums[1] ?? 0) : (nums[2] ?? 0);
      default:
        return Number.NaN;
    }
  };

  const evalExpression = (input: string): number => {
    let s = input.trim();
    // Resolve function calls from innermost to outermost.
    for (let guard = 0; guard < 32; guard++) {
      const m = /\b([A-Z]+)\(([^()]*)\)/i.exec(s);
      if (!m) break;
      const val = evalFunc(m[1]!, m[2]!);
      s = s.slice(0, m.index) + String(val) + s.slice(m.index + m[0].length);
    }
    // Replace standalone A1 references.
    s = s.replace(/\b([A-Z]{1,3}\d+)\b/gi, (tok) => {
      const pos = parseA1(tok);
      if (!pos) return "0";
      const target = matrix[pos.row]?.[pos.col];
      return String(target && target !== cell ? (parseCellNumber(target) ?? 0) : 0);
    });
    // Recursive-descent arithmetic parser (+ - * / ^ and parentheses).
    let pos = 0;
    const skip = () => {
      while (pos < s.length && /\s/.test(s[pos]!)) pos++;
    };
    const parseAdd = (): number => {
      let left = parseMul();
      for (;;) {
        skip();
        if (s[pos] === "+") { pos++; left += parseMul(); }
        else if (s[pos] === "-") { pos++; left -= parseMul(); }
        else break;
      }
      return left;
    };
    const parseMul = (): number => {
      let left = parsePow();
      for (;;) {
        skip();
        if (s[pos] === "*") { pos++; left *= parsePow(); }
        else if (s[pos] === "/") { pos++; const r = parsePow(); left = r === 0 ? Number.NaN : left / r; }
        else break;
      }
      return left;
    };
    const parsePow = (): number => {
      const base = parseUnary();
      skip();
      if (s[pos] === "^") { pos++; return base ** parsePow(); }
      return base;
    };
    const parseUnary = (): number => {
      skip();
      if (s[pos] === "+") { pos++; return parseUnary(); }
      if (s[pos] === "-") { pos++; return -parseUnary(); }
      if (s[pos] === "(") {
        pos++;
        const v = parseAdd();
        skip();
        if (s[pos] === ")") pos++;
        return v;
      }
      const m = /^\d+(?:\.\d+)?/.exec(s.slice(pos));
      if (!m) return Number.NaN;
      pos += m[0].length;
      return Number(m[0]);
    };
    const result = parseAdd();
    skip();
    return pos === s.length ? result : Number.NaN;
  };

  return formatNumberPicture(evalExpression(expr), picture);
}
