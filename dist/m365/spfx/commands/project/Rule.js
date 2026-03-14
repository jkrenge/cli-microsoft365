export class Rule {
    get supersedes() {
        return [];
    }
    addFinding(findings) {
        this.addFindingWithOccurrences([{
                file: this.file,
                resolution: this.resolution
            }], findings);
    }
    addFindingWithOccurrences(occurrences, findings) {
        this.addFindingWithCustomInfo(this.title, this.description, occurrences, findings);
    }
    addFindingWithCustomInfo(title, description, occurrences, findings) {
        findings.push({
            id: this.id,
            title: title,
            description: description,
            occurrences: occurrences,
            resolutionType: this.resolutionType,
            severity: this.severity,
            supersedes: this.supersedes
        });
    }
}
//# sourceMappingURL=Rule.js.map