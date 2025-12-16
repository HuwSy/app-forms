export interface SharepointChoiceForm {
  [key: string]: any;
  Attachments?: {
    results: SharepointChoiceAttachment[];
  };
}

export interface SharepointChoiceAttachment {
  FileName: string;
  ServerRelativeUrl?: string;
  Data?: ArrayBuffer;
  Length?: number;
  Checked?: boolean;
  Deleted?: boolean;
  TimeCreated?: Date;
  ListItemAllFields?: SharepointChoiceForm;
  OldListItemAllFields?: SharepointChoiceForm;
}

export interface SharepointChoiceList {
  [fieldName: string]: SharepointChoiceField;
}

export interface SharepointChoiceField {
  TypeAsString?: 'Text' | 'Note' | 'DateTime' | 'Number' | 'Integer' | 'Boolean' | 'Choice' | 'MultiChoice' | 'User' | 'UserMulti' | 'URL' | 'Attachments' | 'Lookup' | 'LookupMulti';
  InternalName?: string;
  Scope?: string;
  Title?: string;
  Required?: boolean;
  ReadOnlyField?: boolean;
  Description?: string;
  DefaultValue?: string;
  MaxLength?: number;
  Min?: number;
  Max?: number;
  DisplayFormat?: number; // 0 = date only, 1 = date and time
  Choices?: string[];
  SelectionGroup?: number;
  RichText?: boolean;
  Format?: string;
  FillInChoice?: string;
  AppendOnly?: boolean;
}

export interface SharepointChoiceUser {
  Id: number;
  Title: string;
  LoginName: string;
}

export interface SharepointChoicePermission {
  userId: number;
  perms: {
    [key: string]: boolean;
  };
}

export interface SharepointChoiceTabs {
  [tabName: string]: SharepointChoiceRow[];
}

export interface SharepointChoiceRow {
  _selected?: boolean;
  _editing?: string;
  [key: string]: string | number | boolean | Date | SharepointChoiceRowChild | null | undefined;
}

export interface SharepointChoiceRowChild {
  [key: string]: string | number | boolean | Date | null | undefined;
}

export interface SharepointChoiceSort {
  [tabName: string]: {
    direction: 'asc' | 'desc';
    field: string;
  }[];
}

export interface SharepointChoiceFilter {
  [tabName: string]: {
    equals: string | number | boolean | Date | null;
    contains: string | null;
    greater: number | Date | null;
    less: number | Date | null;
  }[];
}

export interface SharepointChoiceHide {
  [tabName: string]: string[];
}

export interface SharepointChoiceColumn {
  headerName?: string; // display name of the column
  field?: string; // datafield name in the data, supports dot notation for nested fields
  headerTooltip?: string; // tooltip for the header
  nowrap?: boolean; // enforce nowrap in cell content
  cellClicked?: (row: SharepointChoiceRow, target: HTMLElement | EventTarget | undefined) => boolean | Promise<boolean>; // on click of the cell, will override row click
  /*
    // example: edit in modal etc
    {
      private modalResolve?: (value: boolean) => void;
      private modalReject?: (reason?: any) => void;

      cellClicked = async (rowData: any, target: any): Promise<boolean> => {
        return new Promise<boolean>(async (resolve, reject) => {
          this.modalResolve = resolve;
          this.modalReject = reject;

          try {
            //load modal component
          } catch (e) {
            //reject reload data in grid
            this.modalReject?.(e);
            //close the modal
          }
        });
      }

      async saveClicked() {
        try {
          // save data
          this.modalResolve?.(true); // true to trigger cache rebuild
        } catch (e) {
          //dont reject the prmoise as the modal save may get clicked again
          //this.modalReject?.(e);
        }
      }

      calcelClicked() {
        this.modalResolve?.(false); // false to prevent cache rebuild
      }
    }

    // example: on cell change via app-choice
    {
      cellClicked = async (rowData: any, target: any): Promise<boolean> => {
        // if not from app-choice edit return
        if (!target || target.tagName != "APP-CHOICE")
          return false;
        // await any actions on the changed data
        await this.someSaveFunction(row);
        // if editing ended/on change due to clicking a different cell then dont continue to click next cell
        if (!rowData._editing?.endsWith(target.getAttribute('field'))
          return;
        // trigger the next cell to be editable
        target = target.parentNode.nextElementSibling;
        while (target && !target.className.includes('editable')) {
          target = target.nextElementSibling;
        }
        // if there is another on the row make it editable
        if (target) {
          target.click();
          // dont trigger cache rebuild as this may wipe the editable state above
          return false;
        }
        // no more editable in this row then trigger cache rebuild
        return true;
      }
    }
  */
  cellRenderer?: (val: any, row: SharepointChoiceRow, index: number) => string; // must be string template not HTMLElement, if HTML input etc needed then use cellClicked to update or to generate HTML element dynamically
  /*
    return `
      ${
        val == 'Yes'
        ? '✔'
        : '<span style="font-size: 20px;">☐</span>'
      }
      ${val}
    `;
  */
  filter?: 'text' | 'number' | 'date' | 'select' | 'none'; // filter type
  width?: number; // minwidth of the column
  children?: SharepointChoiceColumn[] // children columns, only 1 layer deep supported
  center?: boolean; // center align the column
  sortable?: boolean; // disable sorting on this column
  hide?: boolean | ((tab: string) => boolean); // hide column, or function to determine hide state based on selected tab
  spec?: SharepointChoiceField; // make the cell editable using app-choice, onchange will trigger cell/row clicked for any save actions etc with target tagname of app-choice as only distinguishing factor that its emitted post edit
  _filtervisible?: boolean; // internal use to track filter visibility

}
