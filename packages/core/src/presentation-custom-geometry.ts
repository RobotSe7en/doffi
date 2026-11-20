import type { DeepReadonly } from "./runtime.js";
import type { PresentationCustomGeometry } from "./contracts/presentation.js";
import { DocumentError } from "./errors.js";
import type { PresentationPresetGeometry, PresentationPathCommand } from "./presentation-geometry.js";

export function evaluatePresentationCustomGeometry(source: DeepReadonly<PresentationCustomGeometry>, width: number, height: number): PresentationPresetGeometry {
  const invalid = (): never => { throw new DocumentError("InvalidDocument", "A custom geometry guide or path is invalid."); };
  const emu = 12700;
  width *= emu; height *= emu;
  const values = new Map([["l", 0], ["t", 0], ["r", width], ["b", height], ["w", width], ["h", height], ["hc", width / 2], ["vc", height / 2],
    ["cd2", 10800000], ["cd4", 5400000], ["cd8", 2700000], ["ss", Math.min(width, height)], ["ls", Math.max(width, height)]]);
  for (const divisor of [2, 3, 4, 5, 6, 8, 10, 12, 16, 32]) {
    values.set("wd" + divisor, width / divisor); values.set("hd" + divisor, height / divisor);
    values.set("ssd" + divisor, Math.min(width, height) / divisor);
  }
  const read = (token: string): number => /^-?\d+$/.test(token) ? Number(token) : values.get(token) ?? invalid();
  for (const { name, operator, args } of source.guides) {
    if (!/^[a-z_]/i.test(name) || values.has(name) || args.length !== (operator === "val" ? 1 : 3)) invalid();
    const [first = 0, second = 0, third = 0] = args.map(read);
    const result = operator === "val" ? first : operator === "+-" ? first + second - third : third === 0 ? 0 : first * second / third;
    if (!Number.isFinite(result)) invalid();
    values.set(name, result);
  }
  const finite = (value: number): number => Number.isFinite(value) && Math.abs(value) <= 1e12 ? value / emu : invalid();
  const text = source.textRect;
  const textRect = text ? { x: finite(read(text.left)), y: finite(read(text.top)), width: finite(read(text.right) - read(text.left)), height: finite(read(text.bottom) - read(text.top)) } : { x: 0, y: 0, width: finite(width), height: finite(height) };
  if (textRect.width < 0 || textRect.height < 0) invalid();
  const connections = source.connections.map(({ angle, x, y }) => ({ x: finite(read(x)), y: finite(read(y)), angle: ((read(angle) / 60000) % 360 + 360) % 360 }));
  const paths = source.paths.map(path => {
    if (path.commands[0]?.kind !== "move") invalid();
    let open = false;
    const commands = path.commands.map((command): PresentationPathCommand => {
      if (command.kind === "move") open = true;
      if (!open) invalid();
      if (command.kind === "close") { open = false; return { kind: "close" }; }
      const coordinate = (token: string, extent: number, grid: number | undefined) => finite(grid && /^-?\d+$/.test(token) ? Number(token) / grid * extent : read(token));
      return { kind: command.kind, x: coordinate(command.x, width, path.width), y: coordinate(command.y, height, path.height) };
    });
    return { fill: path.fill, stroke: path.stroke, commands };
  });
  return { textRect, connections, paths };
}
