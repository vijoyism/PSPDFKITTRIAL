import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as PSPDFKit from "pspdfkit";

export class PSPDFKitViewerComponent implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private _notifyOutputChanged: () => void;
    private _instance: any;
    private div: HTMLDivElement;
    private pdfBase64: string = "";
    private _context: ComponentFramework.Context<IInputs>;

    /**
     * Empty constructor.
     */
    constructor() { }

    /**
     * Used to initialize the control instance.
     * @param context The entire property bag available to control via Context Object.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public async init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        this._notifyOutputChanged = notifyOutputChanged;
        this._context = context;
        this.div = document.createElement("div");
        this.div.classList.add("pspdfkit-container");
        this.div.style.height = context.parameters.viewerheight.raw + "px";
        this.div.style.width = context.parameters.viewerwidth.raw + "px";
        this.div.style.backgroundColor = "green";
        container.appendChild(this.div);

        await this.loadPSPDFKit();
    }

    private async loadPSPDFKit() {
        try {
            // Ensure PSPDFKit is unloaded before loading a new instance
            //@ts-expect-error PSPDFKit.unload might not have type definitions for this call
            PSPDFKit.unload(".pspdfkit-container");

            // Load PSPDFKit and initialize the viewer
            //@ts-expect-error PSPDFKit.load type definitions might not be accurate or missing
            this._instance = await PSPDFKit.load({
                disableWebAssemblyStreaming: true,
                baseUrl: "https://unpkg.com/pspdfkit/dist/",
                container: this.div,
                document: this.convertBase64ToArrayBuffer(this._context.parameters.document.raw!)
            });

            const saveButton = {
                type: "custom",
                id: "btnSavePdf",
                title: "Save",
                onPress: async (event: any) => {
                    const pdfBuffer = await this._instance.exportPDF();
                    this.pdfBase64 = this.convertArrayBufferToBase64(pdfBuffer);
                    this._notifyOutputChanged();
                }
            };

            this._instance.setToolbarItems((items: { type: string; id: string; title: string; onPress: (event: any) => void; }[]) => {
                items.push(saveButton);
                return items;
            });
        } catch (error) {
            console.error(error);
        }
    }

    private convertArrayBufferToBase64(buffer: ArrayBuffer): string {
        const binaryString = Array.from(new Uint8Array(buffer))
            .map(byte => String.fromCharCode(byte))
            .join('');
        return btoa(binaryString);
    }

    /**
     * Called when any value in the property bag has changed.
     * @param context The entire property bag available to control via Context Object.
     */
    public async updateView(context: ComponentFramework.Context<IInputs>) {
        this.div.style.height = context.parameters.viewerheight.raw + "px";
        this.div.style.width = context.parameters.viewerwidth.raw + "px";

        if (context.updatedProperties.includes("document")) {
            this._context = context;
            await this.loadPSPDFKit();
        }
    }

    private convertBase64ToArrayBuffer(base64String: string): ArrayBuffer {
        const base64 = base64String.split(',')[1] || base64String;
        const binaryString = window.atob(base64);
        const len = binaryString.length;
        const bytes = new Uint8Array(len);
        for (let i = 0; i < len; i++) {
            bytes[i] = binaryString.charCodeAt(i);
        }
        return bytes.buffer;
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest.
     */
    public getOutputs(): IOutputs {
        const result = { pdfdocument: this.pdfBase64 };
        return result;
    }

    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
