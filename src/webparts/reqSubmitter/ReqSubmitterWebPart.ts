// module for imports begins

import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from "@microsoft/sp-webpart-base";
import { escape } from "@microsoft/sp-lodash-subset";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import styles from "./ReqSubmitter.module.scss";
import * as strings from "reqSubmitterStrings";
import { IReqSubmitterWebPartProps } from "./IReqSubmitterWebPartProps";
import {IListItem} from "./IListItem";

// main code starts here

export default class ReqSubmitterWebPart extends BaseClientSideWebPart<IReqSubmitterWebPartProps> {
private listItemEntityTypeName: string = undefined;
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.reqSubmitter}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
            <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
              <span class="ms-font-xl ms-fontColor-white">Welcome to SharePoint!</span>
              <p class="ms-font-l ms-fontColor-white">Customize SharePoint experiences using Web Parts.</p>
              <p class="ms-font-l ms-fontColor-white">${escape(this.properties.listname)}</p>
              <div class="requestInputArea" >
              <input type="text" class="inputTxtBox" name="InputListValue" />
              <button class="${styles.button} create-Button">
              <span class="${styles.label}">Create item</span>
              </button>
              </div>
              <div class="ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}">
              <div class="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
                <div class="status"></div>
                <ul class="items"><ul>
              </div>
            </div>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
      this.listItemEntityTypeName=undefined;
      this.updateStatus(this.listNotConfigured() ? "Please configure list in Web Part properties" : "Ready");
      this.setButtonsState();
      this.setButtonsEventHandlers();
  }

  // we are going to define our functions here

/**
 * This method will check if the list exist or not
 * return type boolean
 * input none
 * applied on webpart property
 */
  private listNotConfigured(): boolean {
    return this.properties.listname === undefined ||
      this.properties.listname === null ||
      this.properties.listname.length === 0;
  }

  /**
   * 
   * @param status this represents the status whether list name has been updated in webpart property pane 
   * @param items might be removed in final version
   */
  private updateStatus(status: string, items: IListItem[] = []): void {
    this.domElement.querySelector(".status").innerHTML = status;
    // well I do not feel this call is require and in final version this will be
    // removed from here
    // this.updateItemsHtml(items);
  }

/**
 * set button state
 */
private setButtonsState(): void {
    const buttons: NodeListOf<Element> = this.domElement.querySelectorAll(`button.${styles.button}`);
    const listNotConfigured: boolean = this.listNotConfigured();

    for (let i: number = 0; i < buttons.length; i++) {
      const button: Element = buttons.item(i);
      if (listNotConfigured) {
        button.setAttribute("disabled", "disabled");
      }
      // tslint:disable-next-line:one-line
      else {
        button.removeAttribute("disabled");
      }
    }
  }

  private setButtonsEventHandlers(): void {
    const webPart: ReqSubmitterWebPart = this;
    this.domElement.querySelector("button.create-Button").addEventListener("click", () => { webPart.createItem(); });
  }

  /**
   * This method creates list items entered in the textbox
   */
  private createItem()
{
  this.updateStatus("Creating item");
  // notice that we have typecasted html input element
  let textboxvalue : string = (<HTMLInputElement>document.getElementsByClassName("inputTxtBox")[0]).value;
  this.getListItemEntityTypeName()
      .then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          "__metadata": {
            "type": listItemEntityTypeName
          },
          "Title":  textboxvalue
        });
        return this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listname}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              "Accept": "application/json;odata=nometadata",
              "Content-type": "application/json;odata=verbose",
              "odata-version": ""
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): Promise<IListItem> => {
        return response.json();
      })
      .then((item: IListItem): void => {
        this.updateStatus(`Item '${item.Title}' (ID: ${item.Id}) successfully created`);
      }, (error: any): void => {
        this.updateStatus("Error while creating the item: " + error);
      });
  }

   private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeName) {
        resolve(this.listItemEntityTypeName);
        return;
      }

      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.properties.listname}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "odata-version": ""
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this.listItemEntityTypeName);
        });
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
                PropertyPaneTextField("listname", {
                  label:  "List"// strings.listnameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
