import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
export interface IDocumentTimelineWebPartProps {
    siteUrl: string;
    listName: string;
    statusLogListName: string;
}
export default class DocumentTimelineWebPart extends BaseClientSideWebPart<IDocumentTimelineWebPartProps> {
    private get siteUrl();
    private get listName();
    private get statusLogListName();
    render(): void;
    private _bindEvents;
    private _doSearch;
    private _renderTimeline;
    private _el;
    private _showScreen;
    private _esc;
    private _statusBadge;
    private _formatDate;
    private _formatDateTime;
    private _formatDuration;
    protected get dataVersion(): Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=DocumentTimelineWebPart.d.ts.map