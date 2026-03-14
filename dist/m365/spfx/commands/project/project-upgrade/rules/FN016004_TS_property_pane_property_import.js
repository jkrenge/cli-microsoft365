import os from 'os';
import ts from 'typescript';
import { TsRule } from './TsRule.js';
export class FN016004_TS_property_pane_property_import extends TsRule {
    constructor() {
        super();
    }
    get id() {
        return 'FN016004';
    }
    get title() {
        return 'Property pane property import change to @microsoft/sp-property-pane';
    }
    get description() {
        return `Refactor the code to import property pane property from the @microsoft/sp-property-pane npm package instead of the @microsoft/sp-webpart-base package`;
    }
    get resolution() {
        return '';
    }
    get resolutionType() {
        return 'ts';
    }
    get severity() {
        return 'Required';
    }
    visit(project, findings) {
        if (!project.tsFiles) {
            return;
        }
        const propertyPaneObjects = ['PropertyPaneLifeCycleEvent', 'ILifeCycleEventCallback', '_PropertyPaneController', 'PropertyPaneAction', 'IPropertyPaneData', 'IPropertyPaneConfiguration', 'IPropertyPanePage', 'IPropertyPanePageHeader', 'IPropertyPaneGroup', 'IPropertyPaneField', 'PropertyPaneFieldType', 'IPropertyPaneCustomFieldProps', 'PropertyPaneCustomField', 'IPropertyPaneButtonProps', 'PropertyPaneButtonType', 'PropertyPaneButton', 'IPropertyPaneCheckboxProps', 'PropertyPaneCheckbox', 'IPropertyPaneChoiceGroupOptionIconProps', 'IPropertyPaneChoiceGroupProps', 'IPropertyPaneChoiceGroupOption', '_IPropertyPaneChoiceGroupOptionInternal', 'PropertyPaneChoiceGroup', 'IPropertyPaneDropdownProps', 'IPropertyPaneDropdownOption', 'PropertyPaneDropdownOptionType', 'IPropertyPaneDropdownCalloutProps', 'PropertyPaneDropdown', 'IPropertyPaneDynamicFieldFilters', 'IPropertyPaneDynamicFieldProps', 'PropertyPaneDynamicField', 'PropertyPaneDynamicFieldSet', 'IPropertyPaneDynamicFieldSetProps', 'PropertyPaneHorizontalRule', 'IPropertyPaneLabelProps', 'PropertyPaneLabel', 'IPropertyPaneLinkProps', 'PropertyPaneLink', 'IPropertyPaneSliderProps', 'PropertyPaneSlider', 'IPropertyPaneTextFieldProps', 'PropertyPaneTextField', 'IPropertyPaneDynamicTextFieldProps', 'IConfiguredDynamicTextFieldProps', 'PropertyPaneDynamicTextField', 'IPropertyPaneToggleProps', 'PropertyPaneToggle', 'IPropertyPaneSpinButtonProps', 'PropertyPaneSpinButton', 'IPropertyPaneConsumer', 'IDynamicConfiguration', '_IDynamicConfiguration', 'IDynamicDataSharedSourceConfiguration', 'IDynamicDataSharedPropertyConfiguration', 'DynamicDataSharedDepth', 'IDynamicDataSharedSourceFilters', 'IDynamicDataSharedPropertyFilters', 'IPropertyPaneConditionalGroup'];
        const occurrences = [];
        project.tsFiles.forEach(file => {
            const nodes = file.nodes;
            if (!nodes) {
                return;
            }
            const obj = nodes
                .filter(n => ts.isImportDeclaration(n))
                .map(n => n)
                .filter(n => n.getText().indexOf('@microsoft/sp-webpart-base') > 0);
            obj.forEach(n => {
                const resource = n.getText();
                const importsText = resource.replace(/\s/g, '').substring(resource.indexOf('{'));
                const imports = importsText.substring(0, importsText.indexOf('}')).split(',');
                const importsToStay = [];
                const importsToBeMoved = [];
                imports.forEach(importName => {
                    if (propertyPaneObjects.indexOf(importName) > -1) {
                        importsToBeMoved.push(importName);
                    }
                    else {
                        importsToStay.push(importName);
                    }
                });
                if (importsToBeMoved.length > 0) {
                    const newBaseImportDeclaration = `import { ${importsToStay.join(', ')} } from "@microsoft/sp-webpart-base";`;
                    const newPropertiesImportDeclaration = `import { ${importsToBeMoved.join(', ')} } from "@microsoft/sp-property-pane";`;
                    let resolution = newPropertiesImportDeclaration;
                    if (importsToStay.length > 0) {
                        resolution = `${newBaseImportDeclaration}${os.EOL}${resolution}`;
                    }
                    this.addOccurrence(resolution, file.path, project.path, n, occurrences);
                }
            });
        });
        if (occurrences.length > 0) {
            this.addFindingWithOccurrences(occurrences, findings);
        }
    }
}
//# sourceMappingURL=FN016004_TS_property_pane_property_import.js.map