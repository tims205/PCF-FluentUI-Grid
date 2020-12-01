import React, { FunctionComponent, useState } from "react";
import { Fabric, IColumn, Selection, DetailsList, initializeIcons, ITextFieldStyles, DetailsListLayoutMode, PrimaryButton, Stack, SelectionMode, DatePicker, ScrollablePane, ScrollbarVisibility, Sticky, StickyPositionType, IRenderFunction, IDetailsHeaderProps, Dropdown, IDropdownOption, IStackTokens, IDropdownStyles } from "@fluentui/react";
import {calendarStrings} from './CalendarStrings'
import { DayOfWeek } from 'office-ui-fabric-react/lib/Calendar';


initializeIcons();

export interface IGridComponentProps {
  columns: IColumn[],
  items: any[],
  availableSites: any[],
  bookAction: (bookingSlotId?:string) => void;
  filterAction: (siteIds:string[],selectedDate:Date) => void;  
}

interface IEntity {
  recordId: string;
  key: string;
  name: string;
  value: string;
}

const horizontalGapStackTokens: IStackTokens = {
  childrenGap: 10,
  padding: 10,
};

const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 },
};

export const GridComponent: React.FunctionComponent<IGridComponentProps> = (props: IGridComponentProps) => { 
  const [message, setMessage] = useState('Data unavailable');
  const [selectionDetails, setSelectionDetails] = useState<string>();
  const [anythingSelected, setAnythingSelected] = useState<boolean>(false);
  const [selectedDate, setSelectedDate] = useState<Date>(new Date());
  const [selectedSites, setSelectedSites] = React.useState<string[]>([]);

  // Used to set the detailslist header as 'Sticky' 
  const onRenderDetailsHeader: IRenderFunction<IDetailsHeaderProps> = (props, defaultRender) => {
    if (!props) {
      return null;
    }
    return (
      <Sticky stickyPosition={StickyPositionType.Header} isScrollSynced>
        {defaultRender!({
          ...props
        })}
      </Sticky>
    );
  };


  const selection: Selection = new Selection({
    onSelectionChanged: () => {
      setSelectionDetails(getSelectionDetails(selection))
      setAnythingSelected(selection.count === 1 ? true : false)
    }
  });

  const bookAction = () => {
    props.bookAction(selectionDetails);
  }

  
  const filterDateAction = (selectedDate: Date) => {
    setSelectedDate(selectedDate);
    props.filterAction(selectedSites, selectedDate);
  }

  const filterSiteAction = (event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption) => {
    debugger;
    if (item) {
      setSelectedSites(
        item.selected ? [...selectedSites, item.key as string] : selectedSites.filter(key => key !== item.key),
      );
      props.filterAction(item.selected ? [...selectedSites, item.key as string] : selectedSites.filter(key => key !== item.key), selectedDate);
    } 
  }
  
  return (
    <>
      <Stack>
       <ScrollablePane scrollbarVisibility={ScrollbarVisibility.auto}>
         <Sticky stickyPosition={StickyPositionType.Header}>
          <Stack horizontal disableShrink tokens={horizontalGapStackTokens}>
            <span>
              <DatePicker
                firstDayOfWeek={DayOfWeek.Sunday}
                strings={calendarStrings}
                value={selectedDate}
                showGoToToday={true}
                showMonthPickerAsOverlay={true}
                placeholder="Search by date..."
                ariaLabel="Select a date"
                onSelectDate={date => date == null ? console.log('no date') : filterDateAction(date) }
              />
            </span>
            <span>
            <Dropdown
                placeholder="Select Sites"
                // eslint-disable-next-line react/jsx-no-bind
                selectedKeys={selectedSites}
                multiSelect
                options={props.availableSites}
                onChange={(event,item) => {filterSiteAction(event,item)}} 
                styles={dropdownStyles}
              />
            </span>
            <Stack.Item align="end">
              <span>
                <PrimaryButton
                  key= 'bookButton'
                  text="Book"
                  disabled={!anythingSelected}
                  onClick={bookAction}
                />
              </span>
            </Stack.Item>
          </Stack>
        </Sticky>
          <DetailsList
            items={props.items}
            columns={props.columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            selection={selection}
            selectionMode={SelectionMode.single}
            selectionPreservedOnEmptyClick={true}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            onRenderDetailsHeader={onRenderDetailsHeader}
          />
        </ScrollablePane>
      </Stack>
    </>
  )
}

const getSelectionDetails = (selection: Selection) => {
  switch (selection.count) {
    case 0:
      return 'No items selected';
    case 1:
      return (selection.getSelection()[0] as IEntity).recordId;
    default:
      return `${selection.count} items selected`;
  }
}


