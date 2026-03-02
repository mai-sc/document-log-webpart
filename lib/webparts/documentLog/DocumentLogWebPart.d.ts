import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IDocumentLogWebPartProps {
    siteUrl: string;
    listName: string;
}
export default class DocumentLogWebPart extends BaseClientSideWebPart<IDocumentLogWebPartProps> {
    private get siteUrl();
    private get listName();
    private pollInterval;
    dispose(): void;
    private timers;
    private tick;
    private fmt;
    private startTick;
    private stopTick;
    render(): void;
    private _bindEvents;
    private _printSlip;
    private _submitToSharePoint;
    private _uploadAttachments;
    private _pollForCode;
    private _showSuccess;
    private _showError;
    private _showAttachmentWarning;
    private _validate;
    private _getFormData;
    private _el;
    private _tc;
    private _setStatus;
    private _showScreen;
    private _setProgress;
    private _reset;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=DocumentLogWebPart.d.ts.map