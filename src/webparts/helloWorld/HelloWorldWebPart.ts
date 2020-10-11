import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
}

//*********************Following is required for SP List Data***********************/

import { 
  SPHttpClient,
  SPHttpClientResponse
 } from "@microsoft/sp-http";

export interface SPList {
  value: SPListItem[];
}

export interface SPListItem {
  ID: number;
  Title: string;  
  FirstName: string;
  LastName: string;
  EmpID: number;
}

//***************************************END****************************************/

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  //method to get and convert list data into json format for rendering by api call
  private _getListData(): Promise<SPList>{
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl +
      "/_api/web/lists/GetByTitle('EmployeeList')/Items", SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => { return response.json(); });
  }

  //method to render the data came from sharepoint api call
  private _renderList(): void {
    this._getListData()
      .then((response) => {
        let html: string = `<head>
          <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
          <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
          <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
        </head>
        <body>
        
          &nbsp;
          
          </br></br>
          <table class="table table-hover">
            <th>Title</th><th>Emp ID</th><th>First Name</th><th>Last Name</th>`;
      
      response.value.forEach((item: SPListItem) => {
        html += `<tr>
          <td>${item.Title}</td>
          <td>${item.EmpID}</td>
          <td>${item.FirstName}</td>
          <td>${item.LastName}</td>
          <td><a href="#" name="editBtn" id="${item.ID}" class="btn btn-success btn-xs">Edit</td>
          <td><a href="#" name="deleteBtn" id="${item.ID}" class="btn btn-danger btn-xs">Delete</td>
        </tr>`;
      });

      html += `</table>
        <br/><br/>
        
      </body>`;

      const listContainer: Element= this.domElement.querySelector('#spListContainer');  //getting the <div> area to display our data
      listContainer.innerHTML = html;  //filling the targeted <div> area

      //Binding Delete links with the event
      let listItems = document.getElementsByName('deleteBtn');
      for(let j:number=0; j<listItems.length; j++){
        listItems[j].addEventListener('click', (event) =>{
          this._deleteListItem(event);
        });
      }

      //Binding Edit links with the event
      listItems = document.getElementsByName('editBtn');
      for(let j:number=0; j<listItems.length; j++){
        listItems[j].addEventListener('click', (event) =>{
          this._updateListItem(event);
        });
      }

      });
  }

  //resetting all the textboxes to blank
  private ClearMethod(): void {
    document.getElementById('txtTitle')["value"]="";
    document.getElementById('txtEmpID')["value"]="";
    document.getElementById('txtFirstName')["value"]="";
    document.getElementById('txtLastName')["value"]="";
  }

  //Setting onClick methods of the buttons
  private _setButtonsEventHandlers(): void {
    //this.domElement.querySelector('btnUpdate').disabled = true;
    let x = <HTMLButtonElement>document.getElementById("btnUpdate");
    x.disabled = true;

    
    //const webPart: SpListWebPartWebPart = this;
    this.domElement.querySelector('#btnAdd').addEventListener('click', () => { this._createListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.FinalupdateItem(); });
    //this.domElement.querySelector('#btnDelete').addEventListener('click', () => { webPart._deleteListItem(); });
  }

  //Create new item method by onClick event
  private _createListItem(): void {
    //document.getElementById('myForm').style.display = 'block';
    //alert("You Reached here!");

    debugger;
    if(document.getElementById('txtTitle')["value"]=="") {
      alert('Required the Title !!!');
      return;
    }
    if(document.getElementById('txtEmpID')["value"]=="") {
      alert('Required the Email Id !!!');
      return;
    }
    if(document.getElementById('txtFirstName')["value"]=="") {
      alert('Required the First Name !!!');
      return;
    }
    if(document.getElementById('txtLastName')["value"]=="") {
      alert('Required the Last name !!!');
      return;
    }

    const body: string = JSON.stringify({ 
      'Title': document.getElementById('txtTitle')["value"],
      'EmpID': Number(document.getElementById('txtEmpID')["value"]),
      'FirstName': document.getElementById('txtFirstName')["value"],
      'LastName': document.getElementById('txtLastName')["value"]
    });

    //alert('Title:'+document.getElementById('txtTitle')["value"]+' Id:'+Number(document.getElementById('txtEmpID')["value"])+
    //' FName:'+document.getElementById('txtFirstName')["value"]);

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeList')/items`, 
      SPHttpClient.configurations.v1, 
      { 
	    headers: { 
		    'Accept': 'application/json;odata=nometadata', 
	      'Content-type': 'application/json;odata=nometadata', 	
	      'odata-version': ''
    	}, 
	    body: body 
      }) 
    .then((response: SPHttpClientResponse): Promise<SPList>=> { 
	    return response.json(); 
    }) 
    .then((item: SPList): void => {
	      this.ClearMethod();
	      alert('Item has been successfully Saved ');
	      localStorage.removeItem('ItemId');
	      localStorage.clear();
	      this._renderList();
      }, 
      (error: any): void => { 
		    alert(`${error}`); 
      });

  }

  //Edit item method by onClick event
  private _updateListItem(clickedEvent): void {
    //making Update btn clickable
    let x = <HTMLButtonElement>document.getElementById('btnUpdate');
    x.disabled = false;

    x = <HTMLButtonElement>document.getElementById('btnAdd');
    x.disabled = true;

    let me:any = clickedEvent.target;
    this.FillTargetDataWithId(me.id);  //calling Update list item after getting list id
  }

  //Filling the elements the data which we want to edit or update
  private FillTargetDataWithId(Id: number){
    this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeList')/items(${Id})`, 
      SPHttpClient.configurations.v1, 
      { 
	      headers: { 
	        'Accept': 'application/json;odata=nometadata', 
	        'odata-version': ''
       	} 
      })
      .then((response: SPHttpClientResponse)=> { 
	      return response.json(); 
      }) 
     .then((item):void => {
      	document.getElementById('txtTitle')["value"]= item.Title;
	      document.getElementById('txtFirstName')["value"]=item.FirstName;
	      document.getElementById('txtEmpID')["value"]=item.EmpID;
	      document.getElementById('txtLastName')["value"]=item.LastName;
	      localStorage.setItem('ItemId', item.Id);
      }, 
      (error: any): void => { 
	      alert(error);
      }); 
  }

  //Actually Updating the data by onClick event
  private FinalupdateItem(){
    const body: string = JSON.stringify({ 
      'Title': document.getElementById('txtTitle')["value"],
      'EmpID': document.getElementById('txtEmpID')["value"],
      'FirstName': document.getElementById('txtFirstName')["value"],
      'LastName': document.getElementById('txtLastName')["value"]  
    }); 

    this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeList')/items(${localStorage.getItem('ItemId')})`, 
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
        this.ClearMethod();
        localStorage.removeItem('ItemId');
        localStorage.clear();
        let x = <HTMLButtonElement>document.getElementById('btnAdd');
        x.disabled = false;
        this._renderList();
      }, 
      (error: any): void => { 
        alert(`${error}`); 
      }); 
  
    }

  //Delete item method by onClick event
  private _deleteListItem(clickedEvent): void {
    let me:any = clickedEvent.target;
    this.DeleteListItem(me.id);  //calling delete list item after getting list id
  }

  private DeleteListItem(Id:number) {
    if(!window.confirm('Are you sure you want to delete this list item?')) {
      return;
    }

    let etag: string = undefined; 
	  this.context.spHttpClient.post(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('EmployeeList')/items(${Id})`, 
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
			  alert(`Item with ID: ${Id} successfully Deleted`);

			  this._renderList();
      }, 
		  (error: any): void => { 
			  alert(`${error}`); 
      });
  }

  //***********************************execution starts from here like Main*********************************
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.helloWorld }">
        
      <div id="spListContainer">          
      </div>
      <div>
        <form >    
          <br>     
          <div>    
            <div > 
              <input type="text" id="txtTitle" name="Title" placeholder="Title"/>    
              <input type="number" id="txtEmpID" name="EmpID" placeholder="Employee ID"/> 
              <input type="text" id="txtFirstName" name="FirstName" placeholder="First Name"/> 
              <input type="text" id="txtLastName" name="LastName" placeholder="Last Name"/> 
                              
              <br/><br/>
              <button id="btnAdd"  type="submit" class="btn btn-primary">+New Item</button>
              &nbsp;&nbsp;&nbsp;&nbsp;
              <button id="btnUpdate"  type="submit" class="btn btn-primary">Update Item</button>
            </div>    
          </div>    
        </form>    
      </div>

      </div>`;
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
