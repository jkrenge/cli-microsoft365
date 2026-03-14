import fs from 'fs';
import path from 'path';
const copyCommands = {
    bash: {
        copyCommand: 'cp',
        copyDestinationParam: ' '
    },
    powershell: {
        copyCommand: 'Copy-Item',
        copyDestinationParam: ' -Destination '
    },
    cmd: {
        copyCommand: 'copy',
        copyDestinationParam: ' '
    }
};
const createDirectoryCommands = {
    bash: {
        createDirectoryCommand: 'mkdir',
        createDirectoryPathParam: ' "',
        createDirectoryNameParam: '/',
        createDirectoryItemTypeParam: '"'
    },
    powershell: {
        createDirectoryCommand: 'New-Item',
        createDirectoryPathParam: ' -Path "',
        createDirectoryNameParam: '" -Name "',
        createDirectoryItemTypeParam: '" -ItemType "directory"'
    },
    cmd: {
        createDirectoryCommand: 'mkdir',
        createDirectoryPathParam: ' "',
        createDirectoryNameParam: '\\',
        createDirectoryItemTypeParam: '"'
    }
};
const addFileCommands = {
    bash: {
        addFileCommand: 'cat > [FILEPATH] << EOF [FILECONTENT]EOF'
    },
    powershell: {
        addFileCommand: `@'[FILECONTENT]'@ | Out-File -FilePath [FILEPATH]`
    },
    cmd: {
        addFileCommand: `echo [FILECONTENT] > [FILEPATH]`
    }
};
const removeFileCommands = {
    bash: {
        removeFileCommand: 'rm'
    },
    powershell: {
        removeFileCommand: 'Remove-Item'
    },
    cmd: {
        removeFileCommand: 'del'
    }
};
export const fsUtil = {
    readdirR(dir) {
        return fs.statSync(dir).isDirectory()
            ? Array.prototype.concat(...fs.readdirSync(dir).map(f => fsUtil.readdirR(path.join(dir, f))))
            : dir;
    },
    // from: https://stackoverflow.com/a/22185855
    copyRecursiveSync(src, dest, replaceTokens) {
        const exists = fs.existsSync(src);
        const stats = exists && fs.statSync(src);
        const isDirectory = exists && stats.isDirectory();
        if (replaceTokens) {
            dest = replaceTokens(dest);
        }
        if (isDirectory) {
            if (!fs.existsSync(dest)) {
                fs.mkdirSync(dest);
            }
            fs.readdirSync(src).forEach(function (childItemName) {
                fsUtil.copyRecursiveSync(path.join(src, childItemName), path.join(dest, childItemName), replaceTokens);
            });
        }
        else {
            fs.copyFileSync(src, dest);
        }
    },
    getSafeFileName(input) {
        return input.replace(/'/g, "''");
    },
    getCopyCommand(command, shell) {
        return copyCommands[shell][command];
    },
    getDirectoryCommand(command, shell) {
        return createDirectoryCommands[shell][command];
    },
    getAddCommand(command, shell) {
        return addFileCommands[shell][command];
    },
    getRemoveCommand(command, shell) {
        return removeFileCommands[shell][command];
    },
    ensureDirectory(path) {
        if (!fs.existsSync(path)) {
            fs.mkdirSync(path, { recursive: true });
        }
    }
};
//# sourceMappingURL=fsUtil.js.map