const winstaller = require('electron-winstaller');
const path = require('path');

async function createInstaller() {
    console.log('正在建立 Windows 安裝程式...');
    try {
        await winstaller.createWindowsInstaller({
            appDirectory: path.join(__dirname, 'dist/app_folder'),
            outputDirectory: path.join(__dirname, 'installer'),
            authors: 'English Learning App Team',
            exe: 'electron.exe',
            description: '劍橋字典 Excel 表格轉 Anki APKG 檔案工具',
            setupExe: 'EnglishLearningSetup_v0.0.2.exe',
            noMsi: true,
        });
        console.log('安裝程式建立成功！請查看 installer 資料夾。');
    } catch (e) {
        console.error(`安裝程式建立失敗: ${e.message}`);
    }
}

createInstaller();
