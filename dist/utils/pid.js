import child_process from 'child_process';
import fs from 'fs';
import os from 'os';
import { cache } from './cache.js';
function getProcessNameOnMacOs(pid) {
    const res = child_process.spawnSync('ps', ['-o', 'comm=', pid.toString()], { encoding: 'utf8' });
    if (res.error || res.stderr) {
        return undefined;
    }
    else {
        return res.stdout.trim();
    }
}
function getProcessNameOnLinux(pid) {
    if (!fs.existsSync(`/proc/${pid}/stat`)) {
        return undefined;
    }
    const stat = fs.readFileSync(`/proc/${pid}/stat`, 'utf8');
    const start = stat.indexOf('(');
    const procName = stat.substring(start + 1, stat.indexOf(')') - start);
    return procName;
}
function getProcessNameOnWindows(pid) {
    const findProcess = child_process.spawnSync('wmic', ['PROCESS', 'where', `ProcessId=${pid}`, 'get', 'Caption'], { encoding: 'utf8' });
    if (findProcess.error || findProcess.stderr) {
        return undefined;
    }
    else {
        const getPid = child_process.spawnSync('find', ['/V', '"Caption"'], {
            encoding: 'utf8',
            input: findProcess.stdout,
            // must include or passing quoted "Caption" will fail
            windowsVerbatimArguments: true
        });
        if (getPid.error || getPid.stderr) {
            return undefined;
        }
        else {
            return getPid.stdout.trim();
        }
    }
}
export const pid = {
    getProcessName(pid) {
        let processName = cache.getValue(pid.toString());
        if (processName) {
            return processName;
        }
        let getPidName = undefined;
        const platform = os.platform();
        if (platform.indexOf('win') === 0) {
            getPidName = getProcessNameOnWindows;
        }
        if (platform === 'darwin') {
            getPidName = getProcessNameOnMacOs;
        }
        if (platform === 'linux') {
            getPidName = getProcessNameOnLinux;
        }
        if (getPidName) {
            processName = getPidName(pid);
            if (processName) {
                cache.setValue(pid.toString(), processName);
            }
            return processName;
        }
        return undefined;
    },
    isPowerShell() {
        const processName = pid.getProcessName(process.ppid) || '';
        return processName.indexOf('powershell') > -1 ||
            processName.indexOf('pwsh') > -1;
    }
};
//# sourceMappingURL=pid.js.map