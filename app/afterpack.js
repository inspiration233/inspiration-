// electron-builder 打包后钩子: mac 应用做临时(ad-hoc)签名。
// M 芯片 mac 对完全未签名的应用一律报"已损坏"(xattr 也无效),
// ad-hoc 签名后只需"右键→打开"即可运行, 无需任何终端命令。
const { execSync } = require('child_process');

exports.default = async (context) => {
  if (context.electronPlatformName !== 'darwin') return;
  const path = require('path');
  const appPath = path.join(context.appOutDir, `${context.packager.appInfo.productName}.app`);
  console.log('[afterpack] ad-hoc signing:', appPath);
  execSync(`codesign --force --deep --sign - "${appPath}"`, { stdio: 'inherit' });
};
