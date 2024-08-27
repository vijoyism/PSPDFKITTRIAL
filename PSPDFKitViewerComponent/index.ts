import {IInputs, IOutputs} from "./generated/ManifestTypes";
 
export class PSPDFKitViewerComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {
 
    private _notifyOutputChanged: () => void;
    private _instance: any;
    private div:any;
    private pdfBase64: string = "";
    private _context: ComponentFramework.Context<IInputs>;
    private _value1: any;
 
    /**
     * Empty constructor.
     */
    constructor()
    {
 
    }
 
    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container:HTMLDivElement)
    {
        this._notifyOutputChanged = notifyOutputChanged;
        this._context = context;    
        this.div = document.createElement("div");   
        this.div.classList.add("pspdfkit-container");
        this.div.style.height = context.parameters.viewerheight.raw + "px";
        this.div.style.width = context.parameters.viewerwidth.raw + "px";
        this.div.style.backgroundColor ="green";
        container.appendChild(this.div);
        await this.PSPDFKit(this._context);
 
    }
    public async PSPDFKit(context: ComponentFramework.Context<IInputs>)
    {
            const PSPDFKit = await import("pspdfkit/dist/modern/pspdfkit");
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            PSPDFKit.unload(".pspdfkit-container");
            // eslint-disable-next-line @typescript-eslint/ban-ts-comment
            // @ts-ignore
            PSPDFKit.load({
                            disableWebAssemblyStreaming: true,
                            baseUrl : "https://unpkg.com/pspdfkit/dist/",
                            container: ".pspdfkit-container",
                            document: this.convertBase64ToArrayBuffer(context.parameters.document.raw!)
                        }).then((instance: any) => {
                    this._instance = instance;
                    const saveButton = {
                        type: "custom",
                        id: "btnSavePdf",
                        title: "Save",
                        onPress: async (event: any) => 
                        {   
                            const pdfBuffer = await this._instance.exportPDF();
                            this.pdfBase64 = this.convertArrayBufferToBase64(pdfBuffer);
                            this._notifyOutputChanged();
                        }   
         
                    };
     
                instance.setToolbarItems((items: { type: string; id: string; title: string; onPress: (event: any) => void; }[]) => {
                items.push(saveButton);
                 
                return items;
                    });             
            })
                .catch(console.error);
    }
 
    private convertArrayBufferToBase64(buffer: ArrayBuffer): string {
         
        let binary = "";
        const bytes = new Uint8Array(buffer);
        const len = bytes.length;
        for (var i = 0; i < len; i++) {
            binary += String.fromCharCode(bytes[i]);
        }
        return btoa(binary);
    }
 
 
    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public async updateView(context: ComponentFramework.Context<IInputs>)
    {
        // Add code to update control view
         
        this.div.style.height = context.parameters.viewerheight.raw + "px";
        this.div.style.width = context.parameters.viewerwidth.raw + "px";
         
        if(context.updatedProperties.indexOf("document")> -1)
        {
        this._value1 = this.convertBase64ToArrayBuffer(context.parameters.document.raw!);
        this._context = context;
        await this.PSPDFKit(this._context);
         
        }       
    }
 
    private convertBase64ToArrayBuffer(base64String: string):ArrayBuffer {
        var base64result = base64String.toString().split(',')[1];
        const binaryString = window.atob(base64result);
        const bytes = new Uint8Array(base64String.length);
        for (var i = 0; i < base64String.length; i++) {
           bytes[i] = binaryString.charCodeAt(i);
        }
         
        return bytes.buffer;
    }
 
    /** 
     * It is called by the framework prior to a control receiving new data. 
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs
    {
        const result =  {
            pdfdocument : this.pdfBase64
        };
        return result;
    }
    public destroy(): void
    {
        // Add code to cleanup control if necessary
    }
}