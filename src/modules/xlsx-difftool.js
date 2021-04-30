const fs = require('fs');
const path = require('path');
const os = require('os');
const wn = require('when');
const nodefn = require("when/node/function");
const callbacks = require("when/callbacks");
const mkdirp = require('mkdirp');

const difftool = function(cmd, local, remote) {
  return which(cmd).then(function(result) {
    cmd = result;
    return wn.join(difftool.textconv(local)["catch"](function() {
      return local;
    }), difftool.textconv(remote)["catch"](function() {
      return remote;
    }));
  }).then(function(filenames) {
    let cmdString;
    cmdString = "\"" + cmd + "\" \"" + filenames[0] + "\" \"" + filenames[1] + "\"";
    return exec(cmdString);
  })["catch"](function(error) {
    return exit(error.code);
  }).done(function() {
    return exit(0);
  });
};

const textconv = function(src) {
  let filename, textconvFilename, tmpdir;
  tmpdir = os.tmpdir();
  filename = path.basename(src);
  textconvFilename = path.resolve(tmpdir, src + '.textconv');
  return exec("git check-attr diff " + filename).then(function(result) {
    let driver;
    driver = /diff: (.+)/.exec(result)[1];
    if (driver === "unspecified") {
      return wn.reject(new Error("'git check-attr diff " + filename + "' returns: unspecified"));
    } else {
      return exec("git config diff." + driver + ".textconv");
    }
  }).then(function(cmd) {
    return exec((cmd.trim()) + " " + src);
  }).then(function(data) {
    return nodefn.call(mkdirp, path.dirname(textconvFilename)).then(function() {
      return nodefn.call(fs.writeFile, textconvFilename, data);
    });
  }).then(function() {
     return textconvFilename;
  });
};

const exit = function(exitCode) {
  if (process.stdout._pendingWriteReqs || process.stderr._pendingWriteReqs) {
    return process.nextTick(function() {
      return exit(exitCode);
    });
  } else {
    return process.exit(exitCode);
  }
};

const which = function(name) {
  let cmd, i, len, maybe;
  if (which.preset[name] && which.preset[name][os.platform()]) {
    maybe = Array.prototype.slice.call(which.preset[name][os.platform()]);
  } else {
    maybe = [];
  }
  maybe.push(name);
  for (i = 0, len = maybe.length; i < len; i++) {
    cmd = maybe[i];
    if (fs.existsSync(cmd)) {
      return wn.resolve(cmd);
    }
  }
  return wn.reject(name + " is not found");
};

which.preset = {
  diffmerge: {
    darwin: ["/Applications/DiffMerge.app/Contents/MacOS/diffmerge"],
    linux: ["/usr/bin/diffmerge"],
    win32: ["C:/Program Files/SourceGear/Common/DiffMerge/sgdm.exe", "C:/Program Files (x86)/SourceGear/Common/DiffMerge/sgdm.exe"]
  }
};

const exec = function(cmd) {
  var deffered;
  deffered = wn.defer();
  require('child_process').exec(cmd, {maxBuffer: 200000*1024}, function(error, stdin, stderr) {
    if (error) {
        console.log(error);
      if (stderr && stderr.length > 0) {
        error.stderr = stderr;
      }
      return deffered.reject(error);
    } else {
      return deffered.resolve(stdin);
    }
  });
  return deffered.promise;
};

difftool.exec = exec;
difftool.textconv = textconv;
difftool.which = which;
module.exports = difftool;
