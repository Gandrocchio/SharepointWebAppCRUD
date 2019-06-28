import { Version } from "@microsoft/sp-core-library";
import { BaseClientSideWebPart, IPropertyPaneConfiguration } from "@microsoft/sp-webpart-base";
import "@pnp/polyfill-ie11";
import "es6-object-assign/auto";
export interface IHelloWorldWebPartProps {
    description: string;
}
export interface ISPList {
    ID?: string;
    Nome?: string;
    Cognome?: string;
    Title?: string;
}
export default class CRUDHelloWorld extends BaseClientSideWebPart<IHelloWorldWebPartProps> {
    render(): Promise<void>;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
    private AddEventListeners;
    protected AddSPItem(): void;
    private getSPItems;
    private getSPItemsAsync;
    private deleteSPItems;
    private UpdateSPItems;
    private main;
    private logConsole;
}
