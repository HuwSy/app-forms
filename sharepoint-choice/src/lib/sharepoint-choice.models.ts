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

// export excel icon as base64 string
export const ExcelIcon: string = "data:image/png;base64,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";
