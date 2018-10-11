import * as pnp from 'sp-pnp-js';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './QuickLinksWebpartWebPart.module.scss';
import * as strings from 'QuickLinksWebpartWebPartStrings';

export interface IQuickLinksWebpartWebPartProps {
  description: string;
}

require('./app/style.css');
export default class QuickLinksWebpartWebPart extends BaseClientSideWebPart<IQuickLinksWebpartWebPartProps> {

  public constructor() {
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css');

    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/jquery/3.1.1/jquery.min.js', { globalExportsName: 'jQuery' }).then((jQuery: any): void => {
      SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js',  { globalExportsName: 'jQuery' }).then((): void => {        
      });
    });
  }

  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      pnp.setup({
        spfxContext: this.context
      });
      
    });
  }

  public getDataFromList():void {
    var mythis =this;
    pnp.sp.web.lists.getByTitle('QuickLinks').items.get().then(function(result){
      console.log("Got List Data:"+JSON.stringify(result));
      mythis.displayData(result);
    },function(er){
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:"+er);
    });
  }


  public displayData(data):void{
    data.forEach(function(val){

      var url = val.URL?val.URL.Url:"#";
      var myHtml = '<h4>'+
			'<li>'+
				'<i class="fa fa-link"></i>'+ 
				'<a href="'+url+'"  target="_blank"> '+val.Title+'</a>'+
			'</li>'+
		'</h4>';
        var div = document.getElementById("quickLinks");
        div.innerHTML+=myHtml;
    }); 
    
  }

  public render(): void {
    this.domElement.innerHTML = `
    <div class="col-xs-12 col-sm-12 col-md-12 col-lg-12">
      <div class="card card-stats events-news">
          <div class="card-header" style="background-color: #da3b01!important;">
            Quick Links
          </div>
          <div class="card-content panel-body rowtop">
            <ul class="QuickLinks1 list-unstyled" id="quickLinks">
            </ul>
          </div>
          <div class="panel-footer" style="text-align:center color:#337ab7;">
            <a href="/sites/Intranet/SPFX/Lists/QuickLinks/AllItems.aspx" target="_blank">Read more</a>
          </div>
      </div>
    </div>`;

    this.getDataFromList();
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
