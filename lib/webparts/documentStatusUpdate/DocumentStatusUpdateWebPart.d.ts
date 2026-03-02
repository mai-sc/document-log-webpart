import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IDocumentStatusUpdateWebPartProps {
    siteUrl: string;
    listName: string;
}
export default class DocumentStatusUpdateWebPart extends BaseClientSideWebPart<IDocumentStatusUpdateWebPartProps> {
    private get siteUrl();
    private get listName();
    private currentItem;
    render(): void;
    private _bindEvents;
    private _doSearch;
    private _showDetail;
    private _doUpdate;
    private _showSuccess;
    private _el;
    private _showScreen;
    private _escapeHtml;
    private _statusBadge;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=DocumentStatusUpdateWebPart.d.ts.map