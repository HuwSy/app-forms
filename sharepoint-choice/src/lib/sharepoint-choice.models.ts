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
  'odata.context'?: {
    web: any; // PnP Web context
  };
  [fieldName: string]: SharepointChoiceField | { web: any } | undefined;
}

export interface SharepointChoiceField {
  TypeAsString?: 'Text' | 'Note' | 'DateTime' | 'Number' | 'Integer' | 'Boolean' | 'Choice' | 'MultiChoice' | 'User' | 'UserMulti' | 'URL' | 'Attachments' | 'Lookup' | 'LookupMulti';
  InternalName?: string;
  Title?: string;
  Required?: boolean;
  ReadOnlyField?: boolean;
  DefaultValue?: string;
  MaxLength?: number;
  Min?: number;
  Max?: number;
  DisplayFormat?: number; // 0 = date only, 1 = date and time
  Choices?: string[];
  SelectionGroup?: number;
  RichText?: boolean;
  Description?: string;
  Format?: string;
  FillInChoice?: string;
  AppendOnly?: boolean;
}

export interface SharepointChoiceUser {
  Key: string;
  Id: number;
  DisplayText: string;
  Title: string;
  LoginName: string;
  EntityData? : {
    Email?: string;
  }
}

export interface SharepointChoicePermission {
  [key: string]: boolean;
}

export interface SharepointChoiceTabs {
  [tabName: string]: SharepointChoiceRow[];
}

export interface SharepointChoiceRow {
  selected?: boolean;
  [key: string]: string | number | boolean | Date | SharepointChoiceRowChild | null | undefined;
}

export interface SharepointChoiceRowChild {
  [key: string]: string | number | boolean | Date | null | undefined;
}

export interface SharepointChoiceFilter {
  equals: string | number | boolean | Date | null;
  contains: string | null;
  greater: number | Date | null;
  less: number | Date | null;
}

export interface SharepointChoiceColumn {
  headerName?: string; // display name of the column
  field?: string; // datafield name in the data, supports dot notation for nested fields
  headerTooltip?: string; // tooltip for the header
  nowrap?: boolean; // enforce nowrap in cell content
  cellClicked?: (row: SharepointChoiceRow, target: HTMLElement) => void | Promise<void>; // on click of the cell
  /*
        // example: toggle selection and add text input to edit title
        {
          var input = document.createElement('input');
          input.type = 'text';
          input.value = row['title'];
          input.onclick = (e) => {
            e.stopPropagation();
          }
          input.onchange = async (e) => {
            row['title'] = input.value;
            await this.someSaveFunction(row);
            input.remove();
            target.innerHTML = row['title'];
          }
          target.innerHTML = '';
          target.appendChild(input);
          input.focus();
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
  filter?: 'text' | 'number' | 'date' | 'select'; // filter type: 'text' | 'number' | 'date' | 'select'
  width?: number; // minwidth of the column
  children?: SharepointChoiceColumn[] // children columns, only 1 layer deep supported
  center?: boolean; // center align the column
  sortable?: boolean; // disable sorting on this column
  hide?: boolean | ((tab: string) => boolean); // hide column, or function to determine hide state based on selected tab
  _filtervisible?: boolean; // internal use to track filter visibility
}
