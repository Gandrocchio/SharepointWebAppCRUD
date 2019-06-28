import { Version } from "@microsoft/sp-core-library";

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";

import { escape } from "@microsoft/sp-lodash-subset";

//import * as pnp from '../../../node_modules/sp-pnp-js';
import { sp, ItemAddResult, Item } from "@pnp/sp";
import "@pnp/polyfill-ie11"; 
import "es6-object-assign/auto";
import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';
import { PopupWindowPosition } from '@microsoft/sp-webpart-base/lib/propertyPane/propertyPaneFields/propertyPaneLink/IPropertyPaneLink';


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


  public async render(): Promise<void> {
    SP.SOD.executeFunc('sp.js', 'SP.ClientContext', async () => { 
      
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
            </div>
            <div>
                  <input id="Title" placeholder="Title" />          
                  <input id="Nome" placeholder="Nome" />
                  <input id="Cognome" placeholder="Cognome" />
                  <button id="AddSPItem" type="submit">Aggiungi Elemento</button>
                  <button id="UpdateSPItem" type="submit">Aggiorna Elemento</button>
                  <button id="DeleteSPItem" type="submit">Cancella Elemento</button>
            </div>
            <br>
            <div id ="DivGetItems" />
          </div>
        </div>
       </div>
      `;
    this.AddEventListeners();
    console.log("Inizio chiamata getSPItems");
    //this.getSPItems();

    let dialog = SP.UI.ModalDialog.showWaitScreenWithNoClose('Processing...', 'questa modal scomparirà tra qualche secondo...', 130, 350);
    console.log("Inizio render --> main()");
    await this.getSPItemsAsync();
    console.log("Fine render --> main()");
    dialog.close(SP.UI.DialogResult.OK);
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private AddEventListeners() : void  {
    document.getElementById("AddSPItem").addEventListener("click", () => this.AddSPItem());
    document.getElementById("DeleteSPItem").addEventListener("click", () => this.deleteSPItems());
    document.getElementById("UpdateSPItem").addEventListener("click", () => this.UpdateSPItems());
  }

  // REST

  // Insert
  protected AddSPItem() : void {
    sp.web.lists.getByTitle("prova").items.add({
      Nome: document.getElementById("Nome")["value"],
      Cognome: document.getElementById("Cognome")["value"],
      Title: document.getElementById("Title")["value"]
    }).then((result: ItemAddResult) => {
      alert("Operazone Completata");
      console.log(result.item);
    }).catch((errors: any[]) => console.log(errors));
  }

  // GET
  private getSPItems() : void {
    
    sp.web.lists.getByTitle("prova").items.getAll().then((AllItems: ISPList[]) => {
      let stringHtml: string = "<div>";
      //let array: ISPList[];
      //array = [{ Nome: "Gabriele", Cognome: "Ascione", Title: "Prova" },];
      if(AllItems.length > 0) 
      {
        AllItems.forEach((item: ISPList) => { 
          stringHtml += item.ID + " " + item.Cognome + "</br>";
          console.log(item.ID + " " + item.Cognome);
        });
      }
      else {
          stringHtml += "Non ci sono elementi</br>";
      }
      
      stringHtml += "</div>";
      const listContainer: Element = this.domElement.querySelector('#DivGetItems'); 
      listContainer.innerHTML = stringHtml;
    }).catch((error) => console.log(error));
  }

  // GET Asincrona
  private async getSPItemsAsync() : Promise<void> {
    
    let AllItems: ISPList[] = await sp.web.lists.getByTitle("prova").items.getAll(); 
    let stringHtml: string = "<div>";
      
      if(AllItems.length > 0) 
      {
        AllItems.forEach((item: ISPList) => { 
          stringHtml += item.ID + " " + item.Cognome + "</br>";
          console.log(item.ID + " " + item.Cognome);
        });
      }
      else {
          stringHtml += "Non ci sono elementi</br>";
      }
      
      stringHtml += "</div>";
      const listContainer: Element = this.domElement.querySelector('#DivGetItems'); 
      listContainer.innerHTML = stringHtml;
    //});.catch((error) => console.log(error));
  }

  // Delete
  private deleteSPItems() : void {
      var id = 1; // Ricerca tramite nome e cognome
      let list = sp.web.lists.getByTitle("prova");
      list.items.getById(1).delete().then(_ => {});
  }

  // Update
  private UpdateSPItems() : void {
      var id = 1;
      let list = sp.web.lists.getByTitle("prova");

      list.items.getById(1).update({
        Nome: document.getElementById("Nome")["value"],
        Cognome: document.getElementById("Cognome")["value"],
        Title: document.getElementById("Title")["value"]
      }).then(i => {
        console.log(i);
      });
  }

  //Async Call e Promise
  private async main() : Promise<void> {
    console.log("Inizio main()");
    await Promise.all([this.logConsole("1", 10000),this.logConsole("2", 10)]);      
    console.log("Fine main()");    
  }

  private async logConsole(s : string, timeout: number) {
    console.log("Inizio LogConsole");
    return await Promise.resolve(setTimeout(async () => { 
      console.log(s);     
      console.log("Fine LogConsole");    
    }, timeout));
  }

}
