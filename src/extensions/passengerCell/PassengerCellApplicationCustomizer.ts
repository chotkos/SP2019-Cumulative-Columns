import { override } from '@microsoft/decorators'; 
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base'; 
 
 /**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IPassengerCellApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

interface ArrayConstructor {
  from(arrayLike: any, mapFn?, thisArg?): Array<any>;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class PassengerCellApplicationCustomizer
  extends BaseApplicationCustomizer<IPassengerCellApplicationCustomizerProperties> {

  public passengerCellTemplate :string = `
  <table width="300">
    <tr>
      <td width="50">
        <div style="font-size:17px;line-height:48px;text-align:center!center;background-color:#6BA5E7;width:50px;height:50px;border-radius:50px;color:white;display:block;">
          <center>{{imageInitials}}</center>
        </div>
      </td>
      <td> 
        <table>
          <tr><td width="250"><b>{{firstName}} {{lastName}}</b><td></tr>		
          <tr><td width="250">{{ageGroup}}, {{nationality}}</td></tr>
        </table>
      </td>
    </tr>	
  </table>`;



  @override
  public onInit(): Promise<void> { 

    this.renderPassengerCells();
    setInterval(()=>{
      this.renderPassengerCells();
    },500);
 
    return Promise.resolve();
  }

  public renderPassengerCells(){
    let elements = document.querySelectorAll(".od-FieldRenderer-text>span")  as any as Array<HTMLElement>;

    for(var i=0;i<elements.length;i++){
      let e = elements[i];
      let l = e as HTMLDivElement;
      if(l.innerText.indexOf("###passengerCell###")!=-1){
      var data = JSON.parse(l.innerText.split('###')[2]); 
      e.innerHTML   = this.passengerCellTemplate.replace("{{imageInitials}}", data.imageInitials)
                    .replace("{{firstName}}", data.firstName)
                    .replace("{{lastName}}", data.lastName)
                    .replace("{{nationality}}", data.nationality)
                    .replace("{{ageGroup}}", data.ageGroup);
      }
    }
  }

}
