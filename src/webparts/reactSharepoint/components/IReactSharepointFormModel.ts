import { IPickerTerms } from "@pnp/spfx-controls-react/lib/TaxonomyPicker";
export interface IRequestSharepointForm {
    SiteTitle:string;
    SiteDescription:string;
    SiteClassification:string;
    CreatorDivision: any;
    CreatorDepartment: any;
    SiteLocation: any;
    SiteUsage:string;
    SiteOwner:string;
    SiteDeputy:string[];
    SiteMembers:string[];
    SiteVisitors:string[];
    ConfirmationGP:boolean;
    DuplicationCheck:boolean;
    NDAConfirmation:boolean;
}