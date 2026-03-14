import fs from 'fs';
import os from 'os';
import path from 'path';
export class FileTokenStorage {
    static msalCacheFilePath() {
        return path.join(os.homedir(), '.cli-m365-msal.json');
    }
    static connectionInfoFilePath() {
        return path.join(os.homedir(), '.cli-m365-connection.json');
    }
    static allConnectionsFilePath() {
        return path.join(os.homedir(), '.cli-m365-all-connections.json');
    }
    constructor(filePath) {
        this.filePath = filePath;
    }
    get() {
        return new Promise((resolve, reject) => {
            if (!fs.existsSync(this.filePath)) {
                reject('File not found');
                return;
            }
            const contents = fs.readFileSync(this.filePath, 'utf8');
            resolve(contents);
        });
    }
    set(connectionInfo) {
        return new Promise((resolve, reject) => {
            fs.writeFile(this.filePath, connectionInfo, 'utf8', (err) => {
                if (err) {
                    reject(err.message);
                }
                else {
                    resolve();
                }
            });
        });
    }
    remove() {
        return new Promise((resolve, reject) => {
            if (!fs.existsSync(this.filePath)) {
                resolve();
                return;
            }
            fs.writeFile(this.filePath, '', 'utf8', (err) => {
                if (err) {
                    reject(err.message);
                }
                else {
                    resolve();
                }
            });
        });
    }
}
//# sourceMappingURL=FileTokenStorage.js.map