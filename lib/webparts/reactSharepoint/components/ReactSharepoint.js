var __extends = (this && this.__extends) || (function () {
    var extendStatics = function (d, b) {
        extendStatics = Object.setPrototypeOf ||
            ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
            function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
        return extendStatics(d, b);
    };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import { FontWeights, Label, TooltipHost, IconButton, TextField, ComboBox, Toggle, PrimaryButton } from "office-ui-fabric-react";
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { CommunicationColors, FontSizes } from "@uifabric/fluent-theme";
import { sp } from "@pnp/sp/presets/all";
import * as $ from 'jquery';
import { TaxonomyPicker } from '@pnp/spfx-controls-react/lib/TaxonomyPicker';
var ReactSharepoint = /** @class */ (function (_super) {
    __extends(ReactSharepoint, _super);
    function ReactSharepoint(IRequestSharePointSitePnp) {
        var _this = _super.call(this, IRequestSharePointSitePnp) || this;
        _this.check = true;
        _this.siteCollectionOptions = [];
        _this.siteUsageOptions = [];
        _this.handleSiteName = function (value) {
            return _this.setState({ SiteTitle: value });
        };
        _this.handleDescription = function (value) {
            return _this.setState({ SiteDescription: value });
        };
        _this.handleUsage = function (item) {
            _this.setState({ SiteClassification: item.text });
            if (item.text == "Work with externals") {
                _this.check = false;
            }
            else {
                _this.check = true;
            }
        };
        _this.handleDepartment = function (item) {
            console.log(item);
            sp.web.lists.getByTitle("Provisioning Process Governance")
                .items.select("Creator Department", "Title").
                expand("Creator Department").get()
                .then(function (answer) {
                console.log(answer);
            });
            var it = {
                __metadata: { "type": "Collection(SP.Taxonomy.TaxonomyFieldValue)" },
                Label: "1",
                TermGuid: item[0].key,
                WssId: -1
            };
            _this.setState({ CreatorDepartment: item });
        };
        _this.handleDivision = function (item) {
            _this.setState({ CreatorDivision: item[0].key.toString() });
        };
        _this.handleLocation = function (item) {
            _this.setState({ SiteLocation: item[0].key.toString() });
        };
        _this.handleOneNote = function (item) {
            _this.setState({ SiteUsage: item.text });
        };
        _this.handleOwner = function (user) {
            _this.setState({ SiteOwner: user[0].id.toString() });
        };
        _this.handleDeputy = function (user) {
            var x = [];
            user.forEach(function (element) {
                x.push(element.id.toString());
            });
            _this.setState({ SiteDeputy: x });
        };
        _this.handleMembers = function (user) {
            var x = [];
            user.forEach(function (element) {
                x.push(element.id.toString());
            });
            _this.setState({ SiteMembers: x });
        };
        _this.handleVisitors = function (user) {
            var x = [];
            user.forEach(function (element) {
                x.push(element.id.toString());
            });
            _this.setState({ SiteVisitors: x });
        };
        _this.handleDuplicationCheck = function (checked) {
            _this.setState({ DuplicationCheck: checked });
        };
        _this.handleGuidelinesConformation = function (checked) {
            _this.setState({ ConfirmationGP: checked });
        };
        _this.handleNdaConfirmation = function (checked) {
            _this.setState({ NDAConfirmation: checked });
        };
        _this.getSiteClassification = function () {
            _this.siteCollectionOptions = [];
            sp.web.lists
                .getByTitle("Provisioning Process Governance")
                .fields.getByInternalNameOrTitle("Site Classification")
                .get()
                .then(function (result) {
                result["Choices"].forEach(function (element, index) {
                    _this.siteCollectionOptions.push({ key: index, text: element });
                });
            });
        };
        _this.getSiteUsageOptions = function () {
            _this.siteUsageOptions = [];
            sp.web.lists
                .getByTitle("Provisioning Process Governance")
                .fields.getByInternalNameOrTitle("SiteUsage")
                .get()
                .then(function (result) {
                result["Choices"].forEach(function (element, index) {
                    _this.siteUsageOptions.push({ key: index, text: element });
                });
            });
        };
        _this.handleSubmit = function () {
            var id;
            console.log(_this.state);
            console.log(sp.web.lists.getByTitle("Provisioning Process Governance").fields.get());
            console.log(sp.web.lists.getByTitle("Provisioning Process Governance").fields.filter('ReadOnlyField eq false and Hidden eq false').get());
            console.log(sp.web.lists.getByTitle("Provisioning Process Governance").items.get());
            console.log(_this.state.CreatorDepartment, typeof (_this.state.CreatorDepartment));
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
            var webUrl = _this.props.siteUrl;
            var listTitle = 'Provisioning Process Governance';
            var itemId = 1;
            var itemPayload = {
                CreatorDepartment: {
                    '__metadata': { 'type': 'SP.Taxonomy.TaxonomyFieldValue' },
                    'Label': _this.state.CreatorDepartment.Label,
                    'TermGuid': _this.state.CreatorDepartment.key,
                    'WssId': -1
                }
            };
            _this.updateListItem(webUrl, listTitle, itemId, itemPayload)
                .done(function () {
                console.log('Tax field valued has been updated');
            })
                .fail(function (error) {
                console.log(JSON.stringify(error));
            });
        };
        _this.state = {
            SiteTitle: "",
            SiteDescription: "",
            SiteClassification: "",
            CreatorDivision: undefined,
            CreatorDepartment: undefined,
            SiteLocation: undefined,
            SiteUsage: "",
            SiteOwner: "",
            SiteDeputy: [],
            SiteMembers: [],
            SiteVisitors: [],
            ConfirmationGP: false,
            DuplicationCheck: false,
            NDAConfirmation: false
        };
        return _this;
    }
    ReactSharepoint.prototype.render = function () {
        var headerLabel = {
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
        this.handleDepartment = this.handleDepartment.bind(this);
        this.handleDivision = this.handleDivision.bind(this);
        this.handleLocation = this.handleLocation.bind(this);
        this.handleSubmit = this.handleSubmit.bind(this);
        this.handleOwner = this.handleOwner.bind(this);
        this.handleDeputy = this.handleDeputy.bind(this);
        this.handleMembers = this.handleMembers.bind(this);
        this.handleVisitors = this.handleVisitors.bind(this);
        this.handleDuplicationCheck = this.handleDuplicationCheck.bind(this);
        this.handleGuidelinesConformation = this.handleGuidelinesConformation.bind(this);
        this.handleNdaConfirmation = this.handleNdaConfirmation.bind(this);
        return (React.createElement("div", null,
            React.createElement("div", null,
                React.createElement(Label, { styles: headerLabel }, "Request New Site")),
            React.createElement("div", null,
                React.createElement("form", null,
                    React.createElement("div", { className: "ms-Grid", dir: "ltr" },
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Name of your site"),
                                React.createElement(TooltipHost, { content: "Use only English characters. Limited to 256 characters." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(TextField, { value: this.state.SiteTitle, onChanged: this.handleSiteName }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Description of the use of this site"),
                                React.createElement(TooltipHost, { content: "Description will  show up on the landing page of your new site." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(TextField, { value: this.state.SiteDescription, onChanged: this.handleDescription }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, null, "Classify the usage of your site"),
                                React.createElement(TooltipHost, { content: "One option only. Note: Sites for internal use only cannot be changed into site for external use later." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(ComboBox, { value: this.state.SiteClassification, onChanged: this.handleUsage, options: this.siteCollectionOptions }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "For which function is the collaboration site?"),
                                React.createElement(TooltipHost, { content: "If it is more than one department, please select 'Cross functional'." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(TaxonomyPicker, { termsetNameOrID: "b7462828-8671-4aad-844a-6ad5c463bb5e", panelTitle: "Select Department", label: "", context: this.props.context, onChange: this.handleDepartment, isTermSetSelectable: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, null, "For which divison is the collaboration site?"),
                                React.createElement(TooltipHost, { content: "If it is more than one division please select 'Cross Divisional'." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(TaxonomyPicker, { termsetNameOrID: "7d36593e-22cf-476a-a5e5-232427d08707", panelTitle: "Select Division", label: "", context: this.props.context, onChange: this.handleDivision, isTermSetSelectable: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "For which location is the collaboration site?"),
                                React.createElement(TooltipHost, { content: "If it is more than one location please select 'Global'." },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(TaxonomyPicker, { termsetNameOrID: "2b198f7e-2824-481e-919e-e838f54891ca", panelTitle: "Select Location", label: "", context: this.props.context, onChange: this.handleLocation, isTermSetSelectable: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Do you plan to use OneNote with your team?"),
                                React.createElement(TooltipHost, { content: "Note: You can activate it later on your own if needed" },
                                    React.createElement(IconButton, { iconProps: { iconName: "Info" } }))),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(ComboBox, { value: this.state.SiteUsage, options: this.siteUsageOptions, onChanged: this.handleOneNote }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Owner (Full Access)")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(PeoplePicker, { selectedItems: this.handleOwner, ensureUser: true, context: this.props.context, personSelectionLimit: 1, showtooltip: true, isRequired: true, principalTypes: [PrincipalType.User], disabled: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Deputy (Full Access)")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(PeoplePicker, { selectedItems: this.handleDeputy, ensureUser: true, context: this.props.context, personSelectionLimit: 1, showtooltip: true, isRequired: true, disabled: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, null, "Members (Contribute rights)")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(PeoplePicker, { selectedItems: this.handleMembers, context: this.props.context, ensureUser: true, personSelectionLimit: 1, showtooltip: true, isRequired: true, disabled: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, null, "Visitors (Read Only)")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(PeoplePicker, { selectedItems: this.handleVisitors, ensureUser: true, context: this.props.context, personSelectionLimit: 1, showtooltip: true, isRequired: true, disabled: false }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "Have you checked with your colleagues that there is no similar site in place?")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(Toggle, { onChanged: this.handleDuplicationCheck }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, null, "Have you read and understood the Guiding principles?")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(Toggle, { onChanged: this.handleGuidelinesConformation }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" }, hidden: this.check },
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6", style: { display: "inline-flex" } },
                                React.createElement(Label, { required: true }, "You are about to create a collaboration site to share content With externals. Do you have Or plan a \u201CNon-Disclosure Agreement\u201D with these externals you want to invite to this collaboration site?")),
                            React.createElement("div", { className: "ms-Grid-col ms-sm6 ms-md6 ms-lg6" },
                                React.createElement(Toggle, { onChanged: this.handleNdaConfirmation }))),
                        React.createElement("div", { className: "ms-Grid-row", style: { padding: "10px" } },
                            React.createElement("div", { className: "ms-Grid-col" },
                                React.createElement(PrimaryButton, { text: "Submit", onClick: this.handleSubmit }))))))));
    };
    ReactSharepoint.prototype.cleanGuid = function (guid) {
        if (guid !== undefined) {
            return guid.replace('/Guid(', '').replace('/', '').replace(')', '');
        }
        else {
            return '';
        }
    };
    ReactSharepoint.prototype.executeJson = function (url, method, headers, payload) {
        method = method || 'GET';
        headers = headers || {};
        headers["Accept"] = "application/json;odata=verbose";
        if (method == "POST") {
            headers["X-RequestDigest"] = $("#__REQUESTDIGEST").val();
        }
        var ajaxOptions = {
            url: url,
            type: method,
            contentType: "application/json;odata=verbose",
            headers: headers,
            data: JSON.stringify(payload)
        };
        return $.ajax(ajaxOptions);
    };
    ReactSharepoint.prototype.updateListItem = function (webUrl, listTitle, itemId, itemPayload) {
        var endpointUrl = webUrl + "/_api/web/lists/GetByTitle('" + listTitle + "')";
        var headers = {};
        headers["X-HTTP-Method"] = "MERGE";
        headers["If-Match"] = "*";
        return this.executeJson(endpointUrl, 'POST', headers, itemPayload);
    };
    return ReactSharepoint;
}(React.Component));
export default ReactSharepoint;
//# sourceMappingURL=ReactSharepoint.js.map