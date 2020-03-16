import * as React from 'react';
import styles from './ReactSharepoint.module.scss';
import { IReactSharepointProps } from './IReactSharepointProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IRequestSharepointForm } from './IReactSharepointFormModel';
import {
  IComboBoxOption, ILabelStyles, FontWeights, Label, TooltipHost, IconButton, TextField, ComboBox, Toggle, PrimaryButton
} from "office-ui-fabric-react";
import { getGUID } from "@pnp/common"; 
import { PeoplePicker,PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CommunicationColors, FontSizes } from "@uifabric/fluent-theme";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "@pnp/sp/presets/all";
import * as $ from 'jquery';
import { TaxonomyPicker, IPickerTerms } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';

export default class ReactSharepoint extends React.Component<IReactSharepointProps, IRequestSharepointForm> {
  constructor(IRequestSharePointSitePnp) {
    super(IRequestSharePointSitePnp);
    this.state = {
      SiteTitle: "",
      SiteDescription: "",
      SiteClassification: "",
      CreatorDivision: undefined,
      CreatorDepartment: undefined,
      SiteLocation: undefined,
      SiteUsage: "",
      SiteOwner:"",
      SiteDeputy:[],
      SiteMembers:[],
      SiteVisitors:[],
      ConfirmationGP: false,
      DuplicationCheck: false,
      NDAConfirmation: false
    };
  }
  private check: boolean = true;
  private siteCollectionOptions: IComboBoxOption[] = [];
  private siteUsageOptions: IComboBoxOption[] = [];
  public render(): React.ReactElement<IReactSharepointProps> {
    const headerLabel: ILabelStyles = {
      root: {
        fontSize: FontSizes.size20,
        backgroundColor: CommunicationColors.primary,
        fontWeight: FontWeights.semibold,
        textAlign: "center"
      }
    };
    this.getSiteClassification();
    this.getSiteUsageOptions();
    this.handleSiteName = this.handleSiteName.bind(this);
    this.handleDescription = this.handleDescription.bind(this);
    this.handleDepartment=this.handleDepartment.bind(this);
    this.handleDivision=this.handleDivision.bind(this);
    this.handleLocation=this.handleLocation.bind(this);
    this.handleSubmit = this.handleSubmit.bind(this);
    this.handleOwner = this.handleOwner.bind(this);
    this.handleDeputy = this.handleDeputy.bind(this);
    this.handleMembers = this.handleMembers.bind(this);
    this.handleVisitors = this.handleVisitors.bind(this);
    this.handleDuplicationCheck = this.handleDuplicationCheck.bind(this);
    this.handleGuidelinesConformation = this.handleGuidelinesConformation.bind(
      this
    );
    this.handleNdaConfirmation = this.handleNdaConfirmation.bind(this);
    return (
      <div>
        <div>
          <Label styles={headerLabel}>Request New Site</Label>
        </div>
        <div>
          <form>
            <div className="ms-Grid" dir="ltr">
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>Name of your site</Label>
                  <TooltipHost content="Use only English characters. Limited to 256 characters.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <TextField
                    value={this.state.SiteTitle}
                    onChanged={this.handleSiteName}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    Description of the use of this site
                  </Label>
                  <TooltipHost content="Description will  show up on the landing page of your new site.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <TextField
                    value={this.state.SiteDescription}
                    onChanged={this.handleDescription}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label>Classify the usage of your site</Label>
                  <TooltipHost content="One option only. Note: Sites for internal use only cannot be changed into site for external use later.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <ComboBox
                    value={this.state.SiteClassification}
                    onChanged={this.handleUsage}
                    options={this.siteCollectionOptions}
                  ></ComboBox>
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    For which function is the collaboration site?
                  </Label>
                  <TooltipHost content="If it is more than one department, please select 'Cross functional'.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <TaxonomyPicker 
                    termsetNameOrID="b7462828-8671-4aad-844a-6ad5c463bb5e"  
                    panelTitle="Select Department"  
                    label=""  
                    context={this.props.context}  
                    onChange={this.handleDepartment}  
                    isTermSetSelectable={false} /> 
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label>For which divison is the collaboration site?</Label>
                  <TooltipHost content="If it is more than one division please select 'Cross Divisional'.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TaxonomyPicker 
                    termsetNameOrID="7d36593e-22cf-476a-a5e5-232427d08707"  
                    panelTitle="Select Division"  
                    label=""  
                    context={this.props.context}  
                    onChange={this.handleDivision}  
                    isTermSetSelectable={false} /> 
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    For which location is the collaboration site?
                  </Label>
                  <TooltipHost content="If it is more than one location please select 'Global'.">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                <TaxonomyPicker 
                    termsetNameOrID="2b198f7e-2824-481e-919e-e838f54891ca"  
                    panelTitle="Select Location"  
                    label=""  
                    context={this.props.context}  
                    onChange={this.handleLocation}  
                    isTermSetSelectable={false} /> 
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    Do you plan to use OneNote with your team?
                  </Label>
                  <TooltipHost content="Note: You can activate it later on your own if needed">
                    <IconButton iconProps={{ iconName: "Info" }} />
                  </TooltipHost>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <ComboBox
                    value={this.state.SiteUsage}
                    options={this.siteUsageOptions}
                    onChanged={this.handleOneNote}
                  ></ComboBox>
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>Owner (Full Access)</Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <PeoplePicker
                    selectedItems={this.handleOwner}
                    ensureUser={true}
                    context={this.props.context}
                    personSelectionLimit={1}
                    showtooltip={true}
                    isRequired={true}
                    principalTypes={[PrincipalType.User]}
                    disabled={false}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>Deputy (Full Access)</Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <PeoplePicker
                    selectedItems={this.handleDeputy}
                    ensureUser={true}
                    context={this.props.context}
                    personSelectionLimit={1}
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label>Members (Contribute rights)</Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <PeoplePicker
                    selectedItems={this.handleMembers}
                    context={this.props.context}
                    ensureUser={true}
                    personSelectionLimit={1}
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label>Visitors (Read Only)</Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <PeoplePicker
                    selectedItems={this.handleVisitors}
                    ensureUser={true}
                    context={this.props.context}
                    personSelectionLimit={1}
                    showtooltip={true}
                    isRequired={true}
                    disabled={false}
                  />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    Have you checked with your colleagues that there is no
                    similar site in place?
                  </Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <Toggle onChanged={this.handleDuplicationCheck} />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label>
                    Have you read and understood the Guiding principles?
                  </Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <Toggle onChanged={this.handleGuidelinesConformation} />
                </div>
              </div>
              <div
                className="ms-Grid-row"
                style={{ padding: "10px" }}
                hidden={this.check}
              >
                <div
                  className="ms-Grid-col ms-sm6 ms-md6 ms-lg6"
                  style={{ display: "inline-flex" }}
                >
                  <Label required={true}>
                    You are about to create a collaboration site to share
                    content With externals. Do you have Or plan a
                    “Non-Disclosure Agreement” with these externals you want to
                    invite to this collaboration site?
                  </Label>
                </div>
                <div className="ms-Grid-col ms-sm6 ms-md6 ms-lg6">
                  <Toggle onChanged={this.handleNdaConfirmation} />
                </div>
              </div>
              <div className="ms-Grid-row" style={{ padding: "10px" }}>
                <div className="ms-Grid-col">
                  <PrimaryButton text="Submit" onClick={this.handleSubmit} />
                </div>
              </div>
            </div>
          </form>
        </div>
      </div>
    );
  }
 
  private handleSiteName = (value: string): void => {
    return this.setState({ SiteTitle: value });
  }
  private handleDescription = (value: string): void => {
    return this.setState({ SiteDescription: value });
  }
  private handleUsage = (item: IComboBoxOption): void => {
    this.setState({ SiteClassification: item.text });
    if (item.text == "Work with externals") {
      this.check = false;
    } else {
      this.check = true;
    }
  }
  private handleDepartment = (item : IPickerTerms): void => {
    console.log(item)
    sp.web.lists.getByTitle("Provisioning Process Governance")
      .items.select("Creator Department" , "Title").
      expand("Creator Department").get()
      .then((answer: any) => {
        console.log( answer);
      })
    let it={
      __metadata: { "type": "Collection(SP.Taxonomy.TaxonomyFieldValue)" },
      Label: "1",
      TermGuid: item[0].key,
      WssId: -1
    }
    this.setState({ CreatorDepartment: item });
  }
  private handleDivision = (item:IPickerTerms): void => {
    this.setState({ CreatorDivision: item[0].key.toString() });
  }
  private handleLocation = (item:IPickerTerms): void => {
    this.setState({ SiteLocation: item[0].key.toString() });
  }
  private handleOneNote = (item: IComboBoxOption): void => {
    this.setState({ SiteUsage: item.text });
  }

  private handleOwner = (user: any[]): void => {
    this.setState({ SiteOwner: user[0].id.toString() });
  }
  private handleDeputy = (user: any[]): void => {
    let x:string[]=[];
    user.forEach(element => {
      x.push(element.id.toString());
    });
    this.setState({ SiteDeputy: x });
  }
  private handleMembers = (user: any[]): void => {
    let x:string[]=[];
    user.forEach(element => {
      x.push(element.id.toString());
    });
    this.setState({ SiteMembers: x });
  }
  private handleVisitors = (user: any[]): void => {
    let x:string[]=[];
    user.forEach(element => {
      x.push(element.id.toString());
    });
    this.setState({ SiteVisitors: x });
  }
  private handleDuplicationCheck = (checked: boolean): void => {
    this.setState({ DuplicationCheck: checked });
  }
  private handleGuidelinesConformation = (checked: boolean): void => {
    this.setState({ ConfirmationGP: checked });
  }
  private handleNdaConfirmation = (checked: boolean): void => {
    this.setState({ NDAConfirmation: checked });
  }
  
  private getSiteClassification = (): void => {
    this.siteCollectionOptions = [];
    sp.web.lists
      .getByTitle("Provisioning Process Governance")
      .fields.getByInternalNameOrTitle("Site Classification")
      .get()
      .then(result => {
        result["Choices"].forEach((element, index) => {
          this.siteCollectionOptions.push({ key: index, text: element });
        });
      });
  }
  private getSiteUsageOptions = (): void => {
    
    this.siteUsageOptions = [];
    sp.web.lists
      .getByTitle("Provisioning Process Governance")
      .fields.getByInternalNameOrTitle("SiteUsage")
      .get()
      .then(result => {
        result["Choices"].forEach((element, index) => {
          this.siteUsageOptions.push({ key: index, text: element });
        });
      });
  }
  private cleanGuid(guid: string): string {
    if (guid !== undefined) {
        return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
    } else {
        return '';
    }
}
  private handleSubmit = (): void => {
    let id:any;
    console.log(this.state);
    console.log(sp.web.lists.getByTitle("Provisioning Process Governance").fields.get());
    console.log(sp.web.lists.getByTitle("Provisioning Process Governance").fields.filter('ReadOnlyField eq false and Hidden eq false').get());
    console.log(sp.web.lists.getByTitle("Provisioning Process Governance").items.get()); 

    console.log(this.state.CreatorDepartment,typeof(this.state.CreatorDepartment))
    //sp.web.lists.getByTitle("Provisioning Process Governance").items.add({
    //  SiteTitle: this.state.SiteTitle,
    //  SiteDescription: this.state.SiteDescription,
    //  SiteClassification: this.state.SiteClassification,

    //  SiteUsage: this.state.SiteUsage,

    //  ConfirmationGP: this.state.ConfirmationGP,
    //  DuplicationCheck: this.state.DuplicationCheck,
    //  NDAConfirmation: this.state.NDAConfirmation
    //}).then((x)=>{
    //  id=x.data.id;
    //})
    var webUrl = this.props.siteUrl;
    var listTitle = 'Provisioning Process Governance';
    var itemId = 1;
    var itemPayload = {
        CreatorDepartment : {
          '__metadata' : {'type': 'SP.Taxonomy.TaxonomyFieldValue'},
          'Label': this.state.CreatorDepartment.Label,  //<-set term label here
          'TermGuid':this.state.CreatorDepartment.key,  //<- set term guid here 
          'WssId': -1
        }
      }; 
    this.updateListItem(webUrl,listTitle,itemId,itemPayload)
   .done(function(){
       console.log('Tax field valued has been updated');
   })
   .fail(function(error){
       console.log(JSON.stringify(error));
   });

  }
  private executeJson(url,method,headers,payload) 
{
    method = method || 'GET';
    headers = headers || {};
    headers["Accept"] = "application/json;odata=verbose";
    if(method == "POST") {
        headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
    }      
    var ajaxOptions = 
    {       
       url: url,   
       type: method,  
       contentType: "application/json;odata=verbose",
       headers: headers,
       data:JSON.stringify(payload)
    };
 
    return $.ajax(ajaxOptions);
}

private updateListItem(webUrl, listTitle,itemId, itemPayload) {
    var endpointUrl = webUrl + "/_api/web/lists/GetByTitle('" + listTitle +  "')";
    var headers = {};
    headers["X-HTTP-Method"] = "MERGE";
    headers["If-Match"] = "*";
    return this.executeJson(endpointUrl,'POST',headers,itemPayload);
}
}
