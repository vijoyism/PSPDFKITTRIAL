/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    document: ComponentFramework.PropertyTypes.StringProperty;
    viewerwidth: ComponentFramework.PropertyTypes.StringProperty;
    viewerheight: ComponentFramework.PropertyTypes.StringProperty;
    pdfdocument: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    pdfdocument?: string;
}
