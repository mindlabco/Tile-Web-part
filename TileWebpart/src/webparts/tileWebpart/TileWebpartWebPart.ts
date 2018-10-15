import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TileWebpartWebPart.module.scss';
import * as strings from 'TileWebpartWebPartStrings';

import * as pnp from "sp-pnp-js";
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as $ from "jquery";
window["$"] = $;
window["jQuery"] = $;
import {  UrlQueryParameterCollection } from '@microsoft/sp-core-library';
export interface ITileWebpartWebPartProps {
  description: string;
}
var mythis: any;
var tileid;
var KeyColor;
var SecColor;
var BarColor;
require("./app/CSS/icommon.css");
require("./app/CSS/style.css");
require("./app/rateyo.js");
export default class TileWebpartWebPart extends BaseClientSideWebPart<ITileWebpartWebPartProps> {
  private imgUrl = require("./app/images/shopping-cart.png");
  private tileData;
  private tileTitle: string;
  private tileImg;
  private tileDescp;
  private tileMobility;
  private clickCount;
  private subscCount;
  private rating;
  private category;
  private tileType;
  private attach;
  private avgRating:number;
  public constructor() {
    super();
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/rateYo/2.3.2/jquery.rateyo.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css');
    SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/css/bootstrap.min.css');
    SPComponentLoader.loadScript('https://cdnjs.cloudflare.com/ajax/libs/twitter-bootstrap/3.3.7/js/bootstrap.min.js');
  }
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {

      pnp.setup({
        spfxContext: this.context
      });

    });
  }

  public getDataFromList(): void {
    pnp.sp.web.lists.getByTitle("LPTiles").items.filter("ID eq '"+tileid+"'").select("Title", "TileType/Tile_x0020_Type", "*").expand("TileType", "AttachmentFiles").get().then(function (result) {

      //console.log("Got List Data:" + JSON.stringify(result));
      if(result.length > 0)
      {
        mythis.displayTile(result);
      }
    }, function (er) {
      alert("Oops, Something went wrong, Please try after sometime");
      console.log("Error:" + er);
    });
  }

  public getThemeColors(): void {
    pnp.sp.web.lists.getByTitle("Site Branding").items.select("Title","Color").get().then(function (result) {
      result.forEach(function(item,index){
            item.Title == "Key Color"?KeyColor = item.Color:item.Title == "Bar Color"?BarColor = item.Color:SecColor = item.Color;
      })
      mythis.getDataFromList();
      
    }, function (er) {
      console.log("Error:" + er);
      mythis.getDataFromList();
    });
  }

  public displayTile(data) {
    this.tileTitle = data[0].Title;
    this.tileImg = data[0].TileImageURL;
    this.tileDescp = data[0].Description?data[0].Description:'';
    this.tileMobility = data[0].Mobile;
    this.clickCount = data[0].Click_x0020_Count;
    this.subscCount = data[0].TileAssignedTo;
    this.subscCount = this.subscCount?this.subscCount.split(','):0;
    this.avgRating = data[0].AverageRating?data[0].AverageRating:0;
    this.rating = data[0].RatingCount?data[0].RatingCount:0;
    this.category = data[0].Metadata;
    this.tileType = data[0].TileType.Tile_x0020_Type;

    this.getTileTypeImg(this.tileType);

    if (this.tileMobility) {
      this.tileMobility = "Yes"
    } else {
      this.tileMobility = "No"
    }
    var html = `<div class="row tile-header">

    <div class="col-xs-2">
      <span class="request" style="background:${KeyColor?KeyColor:""};"><span class="icon-cart"></span></span>
    </div>
    <div class="col-xs-10">
        <p class="tileTitle" style="color:${BarColor?BarColor:""};">${this.tileTitle}</p>
    </div>
</div>

<div class="row tile-middle">
    <div class="col-sm-3">        
      <img class="tile-img" src="${this.tileImg}" alt="tile-img">      
    </div>
    
      <div class="col-sm-9">
          <div class="row tileMiddleRight">
            <div class="col-sm-9"><p class="tileDesc">${this.tileDescp}</p>${this.tileDescp.length > 70?'<button class="btnFullDesc" type="button">Full Description</button>':''}</div>
            <div class="col-sm-3">        
              <img id="tileTypeImg" src="" alt="Tile Type Logo">
              <span class="tile-type" style="color:${BarColor?BarColor:""};">${this.tileType}</span>        
            </div>
          </div>
          <div class="row tileMiddleRight">
            <div class="col-sm-3">
                <p class="mobility" style="color:${BarColor?BarColor:""};">Mobility: ${this.tileMobility}</p>
            </div>
            <div class="col-sm-2">
              <span class="icon-eye"></span>
              <span>${this.clickCount}</span>          
            </div>
            <div class="col-sm-2">
              <i class="fa fa-users"></i>
              <span>(${this.subscCount.length?this.subscCount.length - 1:0})</span>
            </div>
            <div class="col-sm-5">
              <span id="avgRating"></span>
              <span>${(this).avgRating.toFixed(1)} (${this.rating})</span>
            </div>  
          </div>      
      </div>
</div>

<div class="row tile-footer">
    <div class="col-xs-6">
      <ul class="card-title" id="category">      
        Tags
      </ul>
    </div>
    <div class="col-xs-6">
    </div>
</div>`
    $('#tile').append(html);

    $('#avgRating').rateYo({
      rating: data[0].AverageRating?data[0].AverageRating.toFixed(1):0, 
      starWidth:"17px",
      readOnly: true,
      ratedFill:KeyColor?KeyColor:""
    });

    $(document).click(function(e){
      if($(e.target).hasClass('btnFullDesc'))
      {
        $('.tileDesc').removeClass('descExpand');
        $(e.target).siblings('.tileDesc').addClass('descExpand');
      }
      else
      {
        $('.tileDesc').removeClass('descExpand');
      }
  })

    this.category.forEach(function (val) {
      $('#category').append('<li class="category">' + val + '</li>');
    })
  }



  public async getTileTypeImg(tileType) {
    console.log("function");
    this.getTileTypeData(tileType);
  }

  public getTileTypeData(tileType) {
    var context = this;
    try {
      pnp.sp.web.lists.getByTitle("Type List").items.filter("Tile_x0020_Type eq '" + tileType + "'").expand("AttachmentFiles").get().then(function (result) {

        //console.log("Got List Data:" + JSON.stringify(result));
        context.attach = result[0].AttachmentFiles[0].ServerRelativeUrl;
        $('#tileTypeImg').attr('src', context.attach);
      }, function (er) {
        alert("Oops, Something went wrong, Please try after sometime");
        console.log("Error:" + er);

      });
    }
    catch (err) {
      console.log('Error: ', err.message);
    }
  }


  public render(): void {
    mythis = this;
    var queryParameters = new UrlQueryParameterCollection(window.location.href);
    tileid = queryParameters.getValue("tileid");
    this.domElement.innerHTML = `<div class="container-fluide tile" id="tile"></div>`;
    this.getThemeColors();
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
