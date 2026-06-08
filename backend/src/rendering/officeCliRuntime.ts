import { access } from "node:fs/promises";
import { constants } from "node:fs";
import { join, resolve } from "node:path";
import { execFile } from "node:child_process";
import { promisify } from "node:util";

const execFileAsync = promisify(execFile);

export interface OfficeCliResult {
  stdout: string;
  stderr: string;
}

export async function runOfficeCli(args: string[], options?: { cwd?: string; timeoutMs?: number }): Promise<OfficeCliResult> {
  const executable = await resolveOfficeCliPath();
  const result = await execFileAsync(executable, args, {
    cwd: options?.cwd,
    timeout: options?.timeoutMs ?? 120_000,
    windowsHide: true,
    maxBuffer: 8 * 1024 * 1024
  });
  return {
    stdout: result.stdout,
    stderr: result.stderr
  };
}

export async function resolveOfficeCliPath() {
  const configuredPath = process.env.OFFICECLI_PATH?.trim();
  if (configuredPath) {
    await assertExecutable(configuredPath);
    return configuredPath;
  }

  const vendorPath = resolve(process.cwd(), "vendor", "officecli", platformDir(), executableName());
  await assertExecutable(vendorPath);
  return vendorPath;
}

function platformDir() {
  if (process.platform === "win32") {
    return process.arch === "arm64" ? "win-arm64" : "win-x64";
  }
  if (process.platform === "darwin") {
    return process.arch === "arm64" ? "mac-arm64" : "mac-x64";
  }
  return process.arch === "arm64" ? "linux-arm64" : "linux-x64";
}

function executableName() {
  return process.platform === "win32" ? "officecli.exe" : "officecli";
}

async function assertExecutable(path: string) {
  try {
    await access(path, constants.R_OK);
  } catch {
    throw new Error(
      `内置 officeCLI 不可用：${path}。请确认 backend/vendor/officecli 中已包含对应平台的 officecli 二进制，或通过 OFFICECLI_PATH 指向可执行文件。`
    );
  }
}
