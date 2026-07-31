/**
 * afterPack hook - runs after electron-builder packages the app
 * but BEFORE it creates the installer. Patches icon via rcedit directly.
 */
const path = require('path');
const { execFileSync } = require('child_process');
const fs = require('fs');
const os = require('os');

exports.default = async function afterPack(context) {
  if (context.electronPlatformName !== 'win32') return;

  const rceditPath = path.join(
    os.homedir(),
    'AppData', 'Local', 'electron-builder', 'Cache',
    'winCodeSign', 'winCodeSign-2.6.0', 'rcedit-x64.exe'
  );

  const exePath = path.join(
    context.appOutDir,
    `${context.packager.appInfo.productName}.exe`
  );

  const icoPath = path.join(__dirname, 'icon.ico');

  if (!fs.existsSync(rceditPath)) {
    console.log('[afterPack] rcedit not found, skipping icon patch.');
    return;
  }

  if (!fs.existsSync(exePath)) {
    console.log('[afterPack] .exe not found:', exePath);
    return;
  }

  try {
    execFileSync(rceditPath, [exePath, '--set-icon', icoPath]);
    console.log('[afterPack] ✅ Icon patched successfully into:', exePath);
  } catch (err) {
    console.warn('[afterPack] ⚠️ Icon patch failed:', err.message);
  }
};
