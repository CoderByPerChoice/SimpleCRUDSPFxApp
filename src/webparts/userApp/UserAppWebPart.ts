import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './UserAppWebPart.module.scss';
import * as strings from 'UserAppWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface IUserAppWebPartProps {
  description: string;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  Address: string;
  Phone_x0020_Number: string;
  City: string;
  ID: number;
}

export default class UserAppWebPart extends BaseClientSideWebPart<IUserAppWebPartProps> {
  private Listname: string = "Users";

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.userApp}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to user management!</span>
            </div>
          </div>
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.label}">Full Name</span>
            </div>
            <div class="${styles.column}">
              <input type="text" id="txtFullName" placeholder="Enter full name."/>
            </div>
          </div>
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.label}">Address</span>
            </div>
            <div class="${styles.column}">
              <textarea id="txtAddress" placeholder="Enter address." rows="4" cols="50"></textarea>
            </div>
          </div>
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.label}">Phone Number</span>
            </div>
            <div class="${styles.column}">
              <input type="text" id="txtPhoneNumber" placeholder="Enter phone number."/>
            </div>
          </div>
           <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.label}">City</span>
            </div>
            <div class="${styles.column}">
              <select id="ddlCity">
              <option value="Select">~Select City ~</option>
              <option value="Bangalore">Bangalore</option>
              <option value="Hyderabad">Hyderabad</option>
              <option value="New Delhi">New Delhi</option>
              <option value="Mumbai">Mumbai</option>
            </select>
            </div>
          </div>
           <div class="${styles.row}">
            <div class="${styles.column}">
              <button type="button" class="${styles.button}" id="btnSubmit">Save</button>
              <button type="button" class="${styles.button}" id="btnCancel">Cancel</button>
            </div>
          </div>
           <div class="${styles.row}">
            <div class="${styles.column}">
              <table border=1 width=100% style="border-collapse: collapse;" id="userinfo"></table
            </div>
          </div>
        </div>
      </div>`;

    this._bindEvents();
    this._populateUsersList();
  }

  private _bindEvents(): any {
    document.getElementById('btnSubmit').addEventListener('click', () => this.submitRequest());
    document.getElementById('btnCancel').addEventListener('click', () => this.cancelRequest());
    document.getElementById('btnCancel').style.visibility = "hidden";
  }

  private _populateUsersList() {
    this._getListData()
      .then((response) => {
        this._renderList(response.value);
      });
  }

  private _renderList(items: ISPList[]): void {

    var table = document.getElementById("userinfo");
    table.innerHTML = "";
    var thead, tr, td, th;
    table.appendChild(thead = document.createElement("thead"));
    thead.appendChild(tr = document.createElement("tr"));
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "Full Name";
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "Address";
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "Phone Number";
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "City";
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "Delete";
    tr.appendChild(th = document.createElement("th"));
    th.innerHTML = "Edit";

    //Loop through each item of users list.
    items.forEach((item: ISPList) => {

      tr = document.createElement("tr");
      table.appendChild(tr);
      tr.appendChild(td = document.createElement("td"));
      td.innerHTML = item.Title;
      tr.appendChild(td = document.createElement("td"));
      td.innerHTML = item.Address;
      tr.appendChild(td = document.createElement("td"));
      td.innerHTML = item.Phone_x0020_Number;
      tr.appendChild(td = document.createElement("td"));
      td.innerHTML = item.City;

      //Delete button render.
      tr.appendChild(td = document.createElement("td"));
      var btn = document.createElement('input');
      btn.type = "button";
      btn.className = styles.button;
      btn.value = "Delete";
      btn.addEventListener('click', () => this.deleteUser(item.ID));
      td.appendChild(btn);

      //Update button render.
      tr.appendChild(td = document.createElement("td"));
      var btn1 = document.createElement('input');
      btn1.type = "button";
      btn1.className = styles.button;
      btn1.value = "Edit";
      btn1.addEventListener('click', () => this.updateUser(item.ID));
      td.appendChild(btn1);
    });
  }

  public deleteUser(ID) {
    //alert('Deleting item -> ' + ID);
    if (!window.confirm('Are you sure you want to delete the latest item?')) {
      return;
    }

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${ID})`,

      SPHttpClient.configurations.v1,
      {

        headers: {

          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }

      })

      .then((response: SPHttpClientResponse): void => {

        alert(`Item with ID: ${ID} successfully Deleted`);
        this._populateUsersList();

      }, (error: any): void => {
        alert(`${error}`);
      });
  }

  public updateUser(ID) {
    //alert('Updating item -> ' + ID);
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${ID})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'odata-version': ''
        }
      })
      .then((response: SPHttpClientResponse) => {
        return response.json();
      })
      .then((item): void => {

        document.getElementById('txtFullName')["value"] = item.Title;
        document.getElementById('txtAddress')["value"] = item.Address;
        document.getElementById('txtPhoneNumber')["value"] = item.Phone_x0020_Number;
        document.getElementById('ddlCity')["value"] = item.City;

        localStorage.setItem('ItemId', item.Id);
        document.getElementById('btnSubmit').innerText = "Update";
        document.getElementById('btnCancel').style.visibility = "visible";
      }, (error: any): void => {
        alert(error);
      });
  }

  private cancelRequest() {
    this.ClearForm();
    localStorage.removeItem('ItemId');
    document.getElementById('btnSubmit').innerText = "Save";
    document.getElementById('btnCancel').style.visibility = "hidden";
  }

  private submitRequest() {

    if (document.getElementById('txtFullName')["value"] === "") {
      alert('Name cannot be blank!');
      return;
    }

    if (document.getElementById('txtAddress')["value"] === "") {
      alert('Address cannot be blank!');
      return;
    }

    if (document.getElementById('txtPhoneNumber')["value"] === "") {
      alert('Phone number cannot be blank!');
      return;
    }

    if (document.getElementById('ddlCity')["value"] === "Select") {
      alert('City cannot be blank!');
      return;
    }

    const body: string = JSON.stringify({

      'Title': document.getElementById('txtFullName')["value"],
      'Address': document.getElementById('txtAddress')["value"],
      'Phone_x0020_Number': document.getElementById('txtPhoneNumber')["value"],
      'City': document.getElementById('ddlCity')["value"]

    });

    if (document.getElementById('btnSubmit').innerText === "Update") {
      //Update user.
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items(${localStorage.getItem('ItemId')})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'IF-MATCH': '*',
            'X-HTTP-Method': 'MERGE'
          },
          body: body
        })
        .then((response: SPHttpClientResponse): void => {
          alert(`Item with ID: ${localStorage.getItem('ItemId')} successfully updated`);
          this.ClearForm();
          localStorage.removeItem('ItemId');
          localStorage.clear();
          document.getElementById('btnSubmit').innerText = "Save";
          document.getElementById('btnCancel').style.visibility = "hidden";
          this._populateUsersList();
        }, (error: any): void => {
          alert(`${error}`);
        });
    }
    else {
      //New user.
      this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.Listname}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: body
        })
        .then((response: SPHttpClientResponse): Promise<ISPList> => {
          return response.json();
        }).then((item: ISPList): void => {
          this.ClearForm();
          alert('Item has been successfully Saved ');
          this._populateUsersList();
        }, (error: any): void => {
          alert(`${error}`);
        });
    }
  }

  private ClearForm() {
    document.getElementById('txtFullName')["value"] = "";
    document.getElementById('txtAddress')["value"] = "";
    document.getElementById('txtPhoneNumber')["value"] = "";
    document.getElementById('ddlCity')["value"] = "Select";
  }

  private _getListData(): Promise<ISPLists> {

    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + this.Listname + "')/Items", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {

        return response.json();

      });

  }
  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
}
