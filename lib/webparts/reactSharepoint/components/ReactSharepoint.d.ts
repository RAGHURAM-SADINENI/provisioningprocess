import * as React from 'react';
import { IReactSharepointProps } from './IReactSharepointProps';
import { IRequestSharepointForm } from './IReactSharepointFormModel';
export default class ReactSharepoint extends React.Component<IReactSharepointProps, IRequestSharepointForm> {
    constructor(IRequestSharePointSitePnp: any);
    private check;
    private siteCollectionOptions;
    private siteUsageOptions;
    render(): React.ReactElement<IReactSharepointProps>;
    private handleSiteName;
    private handleDescription;
    private handleUsage;
    private handleDepartment;
    private handleDivision;
    private handleLocation;
    private handleOneNote;
    private handleOwner;
    private handleDeputy;
    private handleMembers;
    private handleVisitors;
    private handleDuplicationCheck;
    private handleGuidelinesConformation;
    private handleNdaConfirmation;
    private getSiteClassification;
    private getSiteUsageOptions;
    private cleanGuid;
    private handleSubmit;
    private executeJson;
    private updateListItem;
}
//# sourceMappingURL=ReactSharepoint.d.ts.map