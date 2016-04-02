# windows-shortcut-vbs
This is a really simple package to create Windows shortcuts. It uses VBScript snippets from [this comment from user jorangreef](https://github.com/atom/electron/issues/4414#issuecomment-181281228) to generate a .lnk shortcut. There is no native code, and the only hard dependency is [tmp](https://www.npmjs.com/package/tmp). I am using [prettyjson](https://www.npmjs.com/package/prettyjson) for nice trace output, but that could be removed pretty easily.

Scripts are created as needed as temporary files and removed after execution.

### Seriously VBScripts, why?
I tried it and it worked. It may not work in all situations for a variety of reasons, like your anti-virus not being too happy about running VBScripts. The appeal for me is that there is nothing to compile, and the only external program it calls (cscript) exists by default on most Windows systems.

### How do I use it?
Here are the exposed functions:

    createDesktopShortcut(exePath, shortcutName, cb)
    createShortcutInSpecialFolder(specialFolderName, exePath, shortcutName, cb)
    getSpecialFolder(name, cb)

All exposed functions return a Promise AND take a callback, use whichever method you want to continue code execution. The callback gets (error, fullShortcutPath).

## Usage
```js

var ws_vbs = require('windows-shortcut-vbs');
// uncomment line below to see lots of trace information
// ws_vbs.enableTrace(true);

// Creating shortcut to calc.exe using Promises
ws_vbs.createDesktopShortcut('c:\\Windows\\System32\\calc.exe', 'Super Duper Mathematical Adding Machine').then( (shortcutPath) => {
  console.log(`Shortcut path: ${shortcutPath}`);
}).catch( (err) => {
  console.log(err);
});

// Same as above but using a callback
ws_vbs.createDesktopShortcut('c:\\Windows\\System32\\calc.exe', 'Super Duper Mathematical Adding Machine 2', (err, sp) => {
  if (err) return console.log(err);
  console.log('Shortcut path: ' + sp);
});

```
