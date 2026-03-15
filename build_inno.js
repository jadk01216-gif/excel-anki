const innosetup = require('innosetup-compiler');
const path = require('path');

const setupOptions = {
    gui: false,
    verbose: true,
    iss: path.join(__dirname, 'installer.iss')
};

console.log('正在透過 Inno Setup 建立安裝程式...');

innosetup(setupOptions.iss, {
    gui: setupOptions.gui,
    verbose: setupOptions.verbose
}, function(err) {
    if (err) {
        console.error('安裝程式建立失敗:', err);
    } else {
        console.log('安裝程式建立成功！請查看 output 資料夾。');
    }
});
