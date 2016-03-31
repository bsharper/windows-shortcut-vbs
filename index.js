"use strict";

var fs = require('fs');
var tmp = require('tmp');
var path = require('path');
var spawn = require('child_process').spawn;
var readline = require('readline');
var prettyjson = require('prettyjson');

var traceOutput = function (msg) {
  if (typeof msg === "object") msg = prettyjson.render(msg);
  console.log(msg);
}
var trace = () => {};

var specialFolderCommand = `WScript.Echo WScript.CreateObject("Wscript.Shell").SpecialFolders(WScript.Arguments.Item(0))`;
var createShortcutIcon = `set WshShell = WScript.CreateObject("WScript.Shell")
set oShellLink = WshShell.CreateShortcut(WScript.Arguments.Item(0))
oShellLink.TargetPath = WScript.Arguments.Item(1)
oShellLink.IconLocation = WScript.Arguments.Item(2)
oShellLink.WorkingDirectory = WScript.Arguments.Item(3)
oShellLink.Save`;

function enableTrace(tf) {
  if (tf) trace = traceOutput;
  else trace = () => {};
  if (tf) trace({TraceEnabled: tf});
}


function writeScriptToTemp(script) {
  return new Promise( (resolve, reject) => {
    tmp.tmpName({keep: true},  function _tempNameGenerated(err, path) {
      if (err) {
        reject(err);
        return;
      }
      trace({Script: script});
      fs.writeFile(path, script, (err) => {
        if (err) {
          reject(err);
          return;
        }
        resolve(path);
      });
    });
  });
}

function deleteTempFile(path) {
  return new Promise((resolve, reject) => {
    fs.unlink(path, (err) => {
      if (err) {
        reject(err);
        return;
      }
      resolve();
    });
  });
}


function runScript(path, params) {
  return new Promise( (resolve, reject) => {
    if (! Array.isArray(params)) {
      if (typeof params === "string") {
        params = [params];
      } else {
        params = [];
      }
    }
    var args = ["//NoLogo", "//E:VBScript", path].concat(params);
    trace({ScriptArgs:args});
    var proc = spawn("cscript", args);
    var lines = [];
    var stdout = readline.createInterface({
      input: proc.stdout
    });
    var stderr = readline.createInterface({
      input: proc.stderr
    });
    stdout.on('line', (line) => {
      lines.push(line);
      trace({STDOUT: line});
    });
    stderr.on('line', (line) => {
      lines.push(line);
      trace({STDERR: line});
    });
    proc.on('error', (e) => {
      reject(e);
    });

    proc.on('close', (code) => {
      trace("CLOSE: " + lines);
      resolve({code: code + "", lines: lines});
    });
  });
}

function getSpecialFolder(name, cb) {
  if (process.platform !== "win32") throw new Error("Error: only compatible with Windows")
  return new Promise( (resolve, reject) => {
    var tmpPath;
    var resultPath;
    writeScriptToTemp(specialFolderCommand).then( (path) => {
      tmpPath = path;
      return runScript(path, name);
    }).then((obj) => {
      var code = obj.code;
      var lines = obj.lines;
      trace({Lines: lines});
      trace({ReturnValue:code});
      if (lines.length === 0) {
        throw new Error('No value returned');
      } else {
        resultPath = lines[0];
      }
      return deleteTempFile(tmpPath);
    }).then(() => {
      trace('Temp file deleted');
      if (cb) cb(false, resultPath);
      resolve(resultPath);
    }).catch( (err) => {
      if (err) trace({Error: err});
      reject(err);
      if (cb) cb(err, "");
    });
  })
}

function createShortcutInSpecialFolder(specialFolderName, exePath, shortcutName, cb) {
  if (process.platform !== "win32") throw new Error("Error: only compatible with Windows")
  return new Promise ( (resolve, reject) => {
    if (typeof shortcutName === "undefined") shortcutName = path.basename(exePath, path.extname(exePath));
    shortcutName = (shortcutName.length === 0 ? path.basename(exePath, path.extname(exePath)) : shortcutName);
    if (shortcutName.toLowerCase().indexOf(".lnk") === -1) shortcutName += ".lnk";
    trace({ShortcutName: shortcutName});
    var shortcutPath;
    var tmpPath;
    var targetPath = exePath;
    var workingPath = path.dirname(exePath);
    getSpecialFolder(specialFolderName).then( (dtp) => {
      trace({SpecialFolderPath: dtp});
      shortcutPath = path.join(dtp, shortcutName);
      return writeScriptToTemp(createShortcutIcon);
    }).then( (path) => {
      tmpPath = path;
      return runScript(path, [shortcutPath, targetPath, targetPath, workingPath]);
    }).then( (obj) => {
      var code = obj.code;
      var lines = obj.lines;
      trace({Lines: lines});
      trace({ReturnValue:code});
      return deleteTempFile(tmpPath);
    }).then( () => {
      if (cb) cb(false, shortcutPath);
      resolve(shortcutPath);
    }).catch( (err) => {
      if (cb) cb(err, "");
      if (err) trace({Error: err});
      reject(err);

    });
  });
}


function createDesktopShortcut(exePath, shortcutName, cb) {
  if (process.platform !== "win32") throw new Error("Error: only compatible with Windows")
  return new Promise ( (resolve, reject) => {
    createShortcutInSpecialFolder("Desktop", exePath, shortcutName).then( (shortcutPath) => {
      if (cb) cb(false, shortcutPath);
      resolve(shortcutPath);
    }).catch( (err) => {
      if (cb) cb(err, "");
      reject(err);
    });
  });
}

exports.createDesktopShortcut = createDesktopShortcut;
exports.createShortcutInSpecialFolder = createShortcutInSpecialFolder;
exports.getSpecialFolder = getSpecialFolder;
exports.enableTrace = enableTrace;
