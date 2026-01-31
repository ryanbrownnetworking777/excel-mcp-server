#!/usr/bin/env node
import * as path from 'path'
import * as childProcess from 'child_process'
import * as fs from 'fs'

const BINARY_DISTRIBUTION_PACKAGES: any = {
  win32_ia32: "excel-mcp-server_windows_386_sse2",
  win32_x64: "excel-mcp-server_windows_amd64_v1", 
  win32_arm64: "excel-mcp-server_windows_arm64_v8.0",
  darwin_x64: "excel-mcp-server_darwin_amd64_v1",
  darwin_arm64: "excel-mcp-server_darwin_arm64_v8.0",
  linux_ia32: "excel-mcp-server_linux_386_sse2",
  linux_x64: "excel-mcp-server_linux_amd64_v1",
  linux_arm64: "excel-mcp-server_linux_arm64_v8.0",
}

function getBinaryPath(): string {
  const suffix = process.platform === 'win32' ? '.exe' : '';
  const pkg = BINARY_DISTRIBUTION_PACKAGES[`${process.platform}_${process.arch}`];
  if (pkg) {
    return path.resolve(__dirname, pkg, `excel-mcp-server${suffix}`);
  } else {
    throw new Error(`Unsupported platform: ${process.platform}_${process.arch}`);
  }
}

function logDebug(message: string) {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] DEBUG: ${message}\n`;
  
  // Write to stderr so it doesn't interfere with MCP protocol
  process.stderr.write(logMessage);
  
  // Also write to debug log file
  const logFile = path.join(__dirname, 'excel-mcp-debug.log');
  fs.appendFileSync(logFile, logMessage);
}

try {
  const binaryPath = getBinaryPath();
  logDebug(`Platform: ${process.platform}_${process.arch}`);
  logDebug(`Binary path: ${binaryPath}`);
  logDebug(`Arguments: ${JSON.stringify(process.argv)}`);
  logDebug(`Environment: ${JSON.stringify(process.env)}`);
  
  // Check if binary exists
  if (!fs.existsSync(binaryPath)) {
    throw new Error(`Binary not found: ${binaryPath}`);
  }
  
  logDebug(`Binary exists, starting MCP server...`);
  
  const child = childProcess.spawn(binaryPath, process.argv.slice(2), {
    stdio: ['inherit', 'inherit', 'inherit'],
  });
  
  child.on('error', (error) => {
    logDebug(`Child process error: ${error.message}`);
    process.exit(1);
  });
  
  child.on('exit', (code, signal) => {
    logDebug(`Child process exited with code ${code}, signal ${signal}`);
    process.exit(code || 0);
  });
  
  process.on('SIGTERM', () => {
    logDebug('Received SIGTERM, killing child process');
    child.kill('SIGTERM');
  });
  
  process.on('SIGINT', () => {
    logDebug('Received SIGINT, killing child process');
    child.kill('SIGINT');
  });
  
} catch (error) {
  logDebug(`Launch error: ${error instanceof Error ? error.message : String(error)}`);
  process.exit(1);
}