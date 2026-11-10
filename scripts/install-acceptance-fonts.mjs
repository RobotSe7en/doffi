/** Explicit, isolated acceptance setup. Never imported by an SDK package. */
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { appendFileSync, copyFileSync, existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { basename, join, resolve } from "node:path";

if (process.platform !== "linux") throw new Error("Pinned acceptance font configuration requires Linux Fontconfig and dpkg-deb.");
const root = resolve(import.meta.dirname, ".."), directory = join(root, ".cache/acceptance-fonts"), archives = join(root, ".cache/acceptance-font-archives"), fonts = join(directory, "fonts"), licenses = join(directory, "licenses");
if (process.argv.slice(2).some(arg => arg !== "--system-fonts") || process.argv.length > 3) throw new Error("Use only the optional --system-fonts flag.");
const systemFonts = process.argv.includes("--system-fonts");
const manifest = JSON.parse(readFileSync(join(import.meta.dirname, "fixtures/acceptance-fonts.json"), "utf8"));
const baseline = JSON.parse(readFileSync(join(root, "examples/xlsx/reference.json"), "utf8")).environment.fonts;
const hash = (path, algorithm = "sha256") => createHash(algorithm).update(readFileSync(path)).digest("hex");
for (const path of [directory, archives, fonts, licenses]) mkdirSync(path, { recursive: true });
for (const pkg of manifest.packages) {
  const archive = join(archives, pkg.archive), unpacked = join(directory, "unpacked", pkg.name);
  if (!existsSync(archive)) {
    const partial = archive + ".partial";
    execFileSync("curl", ["--fail", "--location", "--retry", "2", "--connect-timeout", "15", "--max-time", "300", pkg.url, "--output", partial], { stdio: "inherit", timeout: 920_000 });
    if (hash(partial) !== pkg.sha256) { rmSync(partial); throw new Error("Acceptance font archive digest mismatch: " + pkg.name); }
    renameSync(partial, archive);
  }
  if (hash(archive) !== pkg.sha256 || pkg.sha1 && hash(archive, "sha1") !== pkg.sha1) throw new Error("Cached acceptance font archive digest mismatch: " + pkg.name);
  rmSync(unpacked, { recursive: true, force: true }); mkdirSync(unpacked, { recursive: true });
  execFileSync("dpkg-deb", ["--extract", archive, unpacked], { stdio: "inherit", timeout: 60_000 });
  for (const font of pkg.fonts) {
    const file = basename(font.path), source = join(unpacked, font.path);
    if (hash(source) !== font.sha256 || !baseline.some(item => item.file === file && item.sha256 === font.sha256)) throw new Error("Acceptance font differs from the fixed reference: " + file);
    copyFileSync(source, join(fonts, file));
  }
  copyFileSync(join(unpacked, pkg.copyright), join(licenses, pkg.name + ".copyright"));
}
// Firefox may exclude application-cache fonts from its private-context font
// list. Explicit acceptance setup can install these exact bytes in the normal
// system font directory; no browser security preference or SDK code changes.
const selectedFonts = systemFonts ? "/usr/local/share/fonts/doffi-acceptance-" + hash(join(import.meta.dirname, "fixtures/acceptance-fonts.json")).slice(0, 16) : fonts;
if (systemFonts) {
  execFileSync("sudo", ["-n", "install", "-d", "-m", "755", selectedFonts], { stdio: "inherit" });
  for (const pkg of manifest.packages) for (const font of pkg.fonts) {
    const source = join(fonts, basename(font.path)), target = join(selectedFonts, basename(font.path));
    if (existsSync(target) && hash(target) !== font.sha256) throw new Error("Existing host font differs; refusing to overwrite: " + target);
    if (!existsSync(target)) execFileSync("sudo", ["-n", "install", "-m", "644", source, target], { stdio: "inherit" });
  }
}
const names = new Set(manifest.packages.flatMap(pkg => pkg.fonts.map(font => basename(font.path))));
const inherited = { ...process.env }; delete inherited.FONTCONFIG_FILE;
const replaced = [...new Set(execFileSync("fc-list", ["-f", "%{file}\n"], { encoding: "utf8", env: inherited }).trim().split("\n"))].filter(path => names.has(basename(path)) && !path.startsWith(selectedFonts + "/"));
const xml = value => value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
const config = join(directory, "fonts.conf");
writeFileSync(config, '<?xml version="1.0"?>\n<!DOCTYPE fontconfig SYSTEM "urn:fontconfig:fonts.dtd">\n<fontconfig><include>/etc/fonts/fonts.conf</include><dir>' + xml(selectedFonts) + '</dir><selectfont><rejectfont>' + replaced.map(path => '<glob>' + xml(path) + '</glob>').join("") + '</rejectfont></selectfont></fontconfig>\n');
const env = { ...process.env, FONTCONFIG_FILE: config }, resolved = [];
execFileSync("fc-cache", ["--force", selectedFonts], { stdio: "inherit", env, timeout: 60_000 });
for (const pkg of manifest.packages) for (const font of pkg.fonts) {
  const file = execFileSync("fc-match", ["-f", "%{file}", font.pattern], { encoding: "utf8", env }).trim();
  if (resolve(file) !== join(selectedFonts, basename(font.path)) || hash(file) !== font.sha256) throw new Error("Fontconfig did not select the pinned acceptance font: " + font.pattern);
  resolved.push({ pattern: font.pattern, file, sha256: font.sha256 });
}
writeFileSync(join(directory, "environment.json"), JSON.stringify({ ...manifest, installation: systemFonts ? "explicit-system-font-directory" : "application-cache", config, resolved }, null, 2) + "\n");
if (process.env.GITHUB_ENV) appendFileSync(process.env.GITHUB_ENV, "FONTCONFIG_FILE=" + config + "\n");
console.log("Pinned acceptance fonts verified. Use FONTCONFIG_FILE=" + config);
