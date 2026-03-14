import parse from 'json-to-ast';
import { Rule } from './Rule.js';
export class JsonRule extends Rule {
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'json';
    }
    get file() {
        return '';
    }
    addFindingWithPosition(findings, node) {
        this.addFindingWithOccurrences([{
                file: this.file,
                resolution: this.resolution,
                position: this.getPositionFromNode(node)
            }], findings);
    }
    getPositionFromNode(node) {
        if (!node || !node.loc) {
            return { line: 1, character: 1 };
        }
        return {
            line: node.loc.start.line,
            character: node.loc.start.column
        };
    }
    getAstNodeFromFile(jsonFile, jsonProperty) {
        if (!jsonFile.source) {
            return undefined;
        }
        if (!jsonFile.ast) {
            jsonFile.ast = parse(jsonFile.source);
        }
        return this.getAstNodeForProperty(jsonFile.ast, jsonProperty);
    }
    getAstNodeForProperty(node, jsonProperty) {
        if (!node.children || node.children.length === 0) {
            return node;
        }
        if (jsonProperty === '') {
            return node;
        }
        const jsonPropertyChunks = jsonProperty.split('.');
        let currentProperty = jsonPropertyChunks[0];
        currentProperty = currentProperty.replace(/;#/g, '.');
        let isArray = false;
        let arrayElement;
        if (currentProperty.endsWith(']')) {
            isArray = true;
            const pos = currentProperty.indexOf('[') + 1;
            // get array element from the property name
            arrayElement = currentProperty.substring(pos, currentProperty.length - 1);
            // remove array element from the property name
            currentProperty = currentProperty.substring(0, pos - 1);
        }
        for (let i = 0; i < node.children.length; i++) {
            let currentNode = node.children[i];
            if (currentNode.key.value !== currentProperty) {
                continue;
            }
            if (isArray) {
                const arrayIndex = parseInt(arrayElement);
                const arrayElements = currentNode.value.children;
                if (isNaN(arrayIndex)) {
                    for (let j = 0; j < arrayElements.length; j++) {
                        if (arrayElements[j].value === arrayElement) {
                            currentNode = arrayElements[j];
                            break;
                        }
                    }
                }
                else {
                    if (arrayIndex < arrayElements.length) {
                        currentNode = arrayElements[arrayIndex];
                    }
                }
            }
            // if this is the last chunk, return current node
            if (jsonPropertyChunks.length === 1) {
                return currentNode;
            }
            // more chunks left, remove current from the array, and look for child nodes
            jsonPropertyChunks.splice(0, 1);
            return this.getAstNodeForProperty((isArray ? currentNode : currentNode.value), jsonPropertyChunks.join('.'));
        }
        return node;
    }
}
//# sourceMappingURL=JsonRule.js.map