import { IInputs, IOutputs } from './generated/ManifestTypes';
import DataSetInterfaces = ComponentFramework.PropertyHelper.DataSetApi;
type DataSet = ComponentFramework.PropertyTypes.DataSet;
import React, { Component } from 'react';
import ReactDOM from 'react-dom';
import { GridComponent, IGridComponentProps } from './CustomGridComponent';
import { IColumn } from '@fluentui/react';

export class FluentUIAppointmentPicker
    implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;

    private _availableTestSites: any[];

    /**
     * Empty constructor.
     */
    constructor() {}

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ) {
        // Add control initialization code
        this.container = container;

        // Add the initial filter to only show sessions on Today at first load
        context.parameters.gridDataSet.filtering.clearFilter();
        
        let condition:DataSetInterfaces.ConditionExpression = {
            attributeName: context.parameters.gridDataSet.columns.find(e => e.alias === "gridFilterByDate")?.name || "createdon", // Get the date search field from the property-set in the manifest
            conditionOperator: 25, //on
            value: context.formatting.formatDateAsFilterStringInUTC(new Date()) // Filter to show sessions on Today's date
        };

        let conditionsArray: any = [];
        conditionsArray.push(condition);

        context.parameters.gridDataSet.filtering.setFilter({
            conditions: conditionsArray,
            filterOperator: 1
        });

        // Use another data set input to build the site filter drop down options
        this._availableTestSites = this.buildSiteDropdownItems(context);
   
        context.parameters.gridDataSet.refresh();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {

        // Add code to update control view
        const newColumns: IColumn[] = context?.parameters?.gridDataSet?.columns?.map(
            (x) => {
                const column: IColumn = {
                    name: x.displayName,
                    key: x.alias,
                    fieldName: x.name,
                    isResizable: true,
                    minWidth: 100,
                    maxWidth: 200,
                };
                return column;
            }
        );

        // The props that are passed to the grid component
        let props: IGridComponentProps = {
            
            bookAction: (bookingSlotId?:string) => {
                // Call the webapi here to create the new booking record
                this.makeABooking(context, bookingSlotId);
            },

            filterAction: (siteIds?:string[], selectedDate?:Date) => {
                this.buildFilter(context, siteIds, selectedDate);
            },

            availableSites: this._availableTestSites,

            columns: newColumns,
            items: this.buildItems(context, newColumns)
        };

        this.renderControl(context, props);
    }

    private renderControl(context: ComponentFramework.Context<IInputs>, theProps: IGridComponentProps) {
        let gridComponent = React.createElement(GridComponent, theProps, {});

        ReactDOM.render(gridComponent, this.container);

    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {};
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
        ReactDOM.unmountComponentAtNode(this.container);
    }

    public makeABooking(context: ComponentFramework.Context<IInputs>, bookingSlotId?:string, ): void {

        // This would have to be amended based on whatever entity was being created
        console.log(`Make a booking for: ${bookingSlotId}`);
        /*
        // @ts-ignore
        context.mode.contextInfo.entityId;
        // @ts-ignore
        context.mode.contextInfo.entityTypeName
        */
        let recordData: any = {};
        recordData["tim_AppointmentSlot@odata.bind"] = "/tim_appointmentslots(" + bookingSlotId + ")"
        // @ts-ignore - context.mode.contextInfo is unsupported (for now). This is the id of the record that the PCF has been placed on.
        recordData["tim_Contact@odata.bind"] = "/contacts(" + context.mode.contextInfo.entityId + ")"

        context.webAPI.createRecord("tim_appointmentbooking", recordData).then(
            function(response: (ComponentFramework.EntityReference)) {
                console.log(response);

                // navigateTo isn't officially supported by PCF controls yet so using this, once navigateTo support is added then this could be a nice modal dialog
                context.navigation.openForm({ 
                    entityName: "tim_appointmentbooking", 
                    entityId: response.id as unknown as string,
                    openInNewWindow: false,
                    windowPosition: 1
                  });


            }, function(error: any) {
                console.log(error);
          });
    }

    // Pages through the data set to get all the items to show
    private buildItems(context: ComponentFramework.Context<IInputs>, newColumns: IColumn[]): any[] {


        if (context.parameters.gridDataSet.loading === false && context.parameters.gridDataSet.paging.hasNextPage === true) {
            console.log("going to next page");
            console.log("count:" + context.parameters.gridDataSet.sortedRecordIds.length);
            context.parameters.gridDataSet.paging.loadNextPage();
        }


        let newItems: any[] = context?.parameters?.gridDataSet?.sortedRecordIds.map(
            (id) => {
                const record = context.parameters.gridDataSet.records[id];
                const x: any = {};
                for (let index = 0; index < newColumns.length; index++) {
                    const fieldName = newColumns[index].fieldName || '';
                    x[fieldName] = record.getFormattedValue(fieldName);
                    x.recordId = record.getRecordId();
                }
                return x;
                }
            );

        return newItems;
    }


    private buildSiteDropdownItems(context: ComponentFramework.Context<IInputs>): any[] {
        let testingSites: any[] = [];
        context.parameters.sitesDataSet.sortedRecordIds.forEach(r => {
            const record = context.parameters.sitesDataSet.records[r];
            //testingSites.push({key: record.getRecordId(), text: record.getFormattedValue('tim_name')})
            testingSites.push({key: record.getRecordId(), text: record.getFormattedValue(context.parameters.sitesDataSet.columns.find(e => e.alias === "siteDropdownDisplayValue")?.name || "unknown")})
        });

        return testingSites;
    }

    // Use the selected site(s) and date to filter the dataset
    private buildFilter(context: ComponentFramework.Context<IInputs>, siteIds?: string[], selectedDate?: Date) {

        console.log("Entered buildFilter");
        console.log(siteIds);
        console.log(selectedDate);

        context.parameters.gridDataSet.filtering.clearFilter();

        let conditionsArray: any = [];

        if (selectedDate) {
            let dateCondition:DataSetInterfaces.ConditionExpression = {
                attributeName: context.parameters.gridDataSet.columns.find(e => e.alias === "gridFilterByDate")?.name || "createdon", // Get the date search field from the property-set in the manifest
                conditionOperator: 25, //on
                value: context.formatting.formatDateAsFilterStringInUTC(selectedDate)
            };

            conditionsArray.push(dateCondition);
        }

        if (siteIds && siteIds.length > 0) {
            let siteCondition: DataSetInterfaces.ConditionExpression = {
                attributeName: "tim_site", // This needs updating to the attribute name on the gridDataSet that contains the link to the records in sitesDataSet
                conditionOperator: 8, //in
                value: siteIds
            }

            conditionsArray.push(siteCondition);
        }

        let filterExp: DataSetInterfaces.FilterExpression = {
            conditions: conditionsArray,
            filterOperator: 0 // And
        }

        context.parameters.gridDataSet.filtering.setFilter(filterExp);
        context.parameters.gridDataSet.refresh();
    }
}
