import { spfi, SPFI, SPBrowser } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import "@pnp/sp/site-groups";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/attachments";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/profiles";
import "@pnp/sp/search";
import {
  PublicClientApplication,
  createNestablePublicClientApplication,
  type AuthenticationResult,
} from "@azure/msal-browser";
import { PermissionKind } from "@pnp/sp/security";
import { ISearchQuery, ISort, SearchResults } from "@pnp/sp/search";
import { App } from "./sharepoint-choice.settings";
import {
  SharepointChoicePermission,
  SharepointChoiceForm,
  SharepointChoiceList,
  SharepointChoiceField,
  SharepointChoiceAttachment,
  MAX_PASTED_IMAGE_SIZE_BYTES,
  MAX_PASTED_IMAGE_SIZE_MB,
} from "./sharepoint-choice.models";

interface InlineImageUpload {
  field: string;
  fileName: string;
  data: ArrayBuffer;
  placeholder: string;
}

interface InlineImageExtraction {
  uploads: InlineImageUpload[];
  htmlByField: Map<string, string>;
  oversizedCount: number;
}

interface SharePointPageContextInfo {
  webAbsoluteUrl?: string;
  userLoginName?: string;
  userDisplayName?: string;
  userEmail?: string;
  userId?: number;
  webTitle?: string;
}

interface SharePointChoiceWindow extends Window {
  _spPageContextInfo?: SharePointPageContextInfo;
  Office?: {
    context?: {
      requirements?: {
        isSetSupported?: (requirementSet: string, version: string) => boolean;
      };
    };
  };
}

interface SharePointSecurable {
  getCurrentUserEffectivePermissions(): Promise<{ Low: number; High: number }>;
}

interface SharePointVersionChange {
  field: string;
  from: unknown;
  to: unknown;
}

interface SharePointItemVersion extends SharepointChoiceForm {
  Created: string;
  Modified: string;
  ChangedFields?: SharePointVersionChange[];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

interface ListItemValidationResult {
  FieldName?: string;
  FieldValue?: string;
  ItemId?: number;
}

interface ListItemValidationResponse {
  value?: ListItemValidationResult[];
}

export function MapErrors(error: unknown, context?: string): string {
  interface AlertMessageCandidate {
    message: string;
    score: number;
    kind: "extracted" | "guidance";
    details?: string;
  }

  const alertMessageKeyScores: Record<string, number> = {
    usermessage: 110,
    errordescription: 105,
    errormessage: 100,
    message: 95,
    detail: 90,
    details: 85,
    title: 80,
    reason: 75,
    description: 70,
    messages: 70,
    errors: 65,
    error: 60,
    body: 35,
    responsetext: 35,
    returned: 20,
    value: 10,
  };

  function stripStackTrace(details: string): string {
    const lines = details.split(/\r?\n/);
    const stackStart = lines.findIndex((line) => /^\s*(?:at\s+\S+|stack trace\s*:|--- End of stack trace)/i.test(line));
    return lines
      .slice(0, stackStart < 0 ? lines.length : stackStart)
      .join("\n")
      .trim();
  }

  function errorDetails(error: unknown): string {
    if (typeof error === "string") return stripStackTrace(error);
    if (error instanceof Error) return `${error.name}: ${error.message}`;

    try {
      return JSON.stringify(error, (key, value: unknown) => (key === "stack" ? undefined : value)) ?? String(error);
    } catch (serializationError) {
      return stripStackTrace(String(error));
    }
  }

  function errorCodes(error: unknown): string {
    if (typeof error !== "object" || error === null) return "";

    const record = error as Record<string, unknown>;
    return ["errorCode", "code", "errorMessage", "message"]
      .map((key) => record[key])
      .filter((value): value is string => typeof value === "string")
      .join("\n");
  }

  function embeddedJsonValues(text: string): unknown[] {
    const values: unknown[] = [];

    const raw = text.slice(0, 100_000);
    const decoded = raw.replace(/\\"/g, '"').replace(/\\\\u([0-9a-f]{4})/gi, (_match, code: string) => `\\u${code}`);

    for (const searchable of raw === decoded ? [raw] : [raw, decoded]) {
      for (let start = 0; start < searchable.length && values.length < 20; start++) {
        const opening = searchable[start];
        if (opening !== "{" && opening !== "[") continue;

        const closings: string[] = [opening === "{" ? "}" : "]"];
        let inString = false;
        let escaped = false;

        for (let end = start + 1; end < searchable.length; end++) {
          const character = searchable[end]!;
          if (inString) {
            if (escaped) escaped = false;
            else if (character === "\\") escaped = true;
            else if (character === '"') inString = false;
            continue;
          }

          if (character === '"') inString = true;
          else if (character === "{") closings.push("}");
          else if (character === "[") closings.push("]");
          else if (character === "}" || character === "]") {
            if (closings.pop() !== character) break;
            if (closings.length > 0) continue;

            const segment = searchable.slice(start, end + 1);
            let parsed = false;
            let value: unknown;
            try {
              value = JSON.parse(segment);
              parsed = true;
            } catch {}

            if (!parsed) {
              try {
                const decoded = JSON.parse(`"${segment}"`);
                if (typeof decoded === "string") {
                  value = JSON.parse(decoded);
                  parsed = true;
                }
              } catch {}
            }

            if (parsed) {
              values.push(value);
              start = end;
            }
            break;
          }
        }
      }
    }

    return values;
  }

  function alertMessageKeyScore(key: string): number {
    return alertMessageKeyScores[key.toLowerCase().replace(/[-_\s]/g, "")] ?? 0;
  }

  function collectAlertMessageCandidates(error: unknown): AlertMessageCandidate[] {
    const candidates: AlertMessageCandidate[] = [];
    const visited = new WeakSet<object>();
    let visitedCount = 0;

    const addCandidate = (value: string, relevance: number): void => {
      const message = stripStackTrace(value);
      if (!message) return;

      let score = relevance + Math.min(message.length, 240) / 80;
      if (
        /^(?:bad request|internal server error|request failed(?: with status code \d+)?|validation failed|one or more validation errors occurred|the request is invalid|something went wrong|an error occurred|unknown error)[.!:]?$/i.test(
          message,
        )
      )
        score -= 50;
      if (/^(?:exception getting API data|error making HttpClient request|http failure response)/i.test(message))
        score -= 60;
      if (/\b(?:stack trace|at\s+[\w.$]+\s*\()/i.test(message)) score -= 30;

      candidates.push({ message, score, kind: "extracted" });
    };

    const visit = (value: unknown, depth: number, relevance: number, allowPlainText: boolean): void => {
      if (depth > 10 || visitedCount > 300) return;

      if (typeof value === "string") {
        const embedded = embeddedJsonValues(value);
        if (embedded.length > 0) {
          embedded.forEach((item) => visit(item, depth + 1, relevance, false));
        } else if (allowPlainText || relevance > 0) {
          addCandidate(value, relevance);
        }
        return;
      }

      if (typeof value !== "object" || value === null || visited.has(value)) return;
      visited.add(value);
      visitedCount++;

      if (Array.isArray(value)) {
        value.forEach((item) => visit(item, depth + 1, relevance, relevance > 0));
        return;
      }

      if (value instanceof Error) {
        visit(value.message, depth + 1, Math.max(relevance, alertMessageKeyScore("message")), true);
        const cause = (value as Error & { cause?: unknown }).cause;
        if (cause !== undefined) visit(cause, depth + 1, Math.max(relevance, alertMessageKeyScore("error")), false);
        if (typeof value.stack === "string") visit(value.stack, depth + 1, relevance, false);
      }

      for (const key of Object.keys(value)) {
        if (value instanceof Error && (key === "message" || key === "stack" || key === "cause")) continue;
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (!descriptor || !("value" in descriptor)) continue;

        const keyScore = alertMessageKeyScore(key);
        visit(descriptor.value, depth + 1, Math.max(relevance, keyScore), keyScore > 0 || relevance > 0);
      }
    };

    visit(error, 0, 0, true);
    return candidates;
  }

  function alertHeading(context: string | undefined, details: string): string | undefined {
    if (context?.trim()) return context.trim();
    const firstLine = details.trim().split(/\r?\n/, 1)[0]?.trim();
    return firstLine &&
      /^Error\s+(?:saving|loading|retrieving|retreiving|deleting|adding|extracting|importing)[^:\r\n]*:\s*$/i.test(
        firstLine,
      )
      ? firstLine
      : undefined;
  }

  const alertMessageRules: Array<{
    matches: RegExp;
    message: (details: string) => string;
  }> = [
    {
      matches: /locked for shared use/i,
      message: (details) => {
        const fileName = details.match(/Error saving file (.+?) in folder /i)?.[1];
        return `${fileName ? `File "${fileName}"` : "This file"} is already open for shared use, blocking the save. Please close any open tabs or clients using this file before trying again.`;
      },
    },
    {
      matches: /interaction_in_progress/i,
      message: () => "A sign-in is already in progress. Complete the open sign-in flow before trying again.",
    },
    {
      matches: /consent_required/i,
      message: () =>
        "This app needs additional permissions before it can continue. Ask your administrator to grant consent.",
    },
    {
      matches: /login_required|interaction_required/i,
      message: () => "Your sign-in needs to be refreshed. Sign in again, then retry.",
    },
    {
      matches: /\[(?:401|403)\]|access denied|unauthorized|forbidden/i,
      message: () =>
        "You do not have permission to complete this SharePoint action. Sign in with an account that has access, then try again.",
    },
    {
      matches: /\[404\]|file not found|does not exist/i,
      message: () =>
        "The requested SharePoint item could not be found. It may have been moved or deleted; refresh and try again.",
    },
    {
      matches: /failed to fetch|network error|network request failed/i,
      message: () => "The request could not reach the service. Check your connection and try again.",
    },
  ];

  const rawDetails = errorDetails(error);
  const details = context ? `${context}\n\n${rawDetails}` : rawDetails;
  const candidates = collectAlertMessageCandidates(error);
  const bestExtractedCandidate = candidates.reduce<AlertMessageCandidate | undefined>(
    (best, candidate) => (best === undefined || candidate.score > best.score ? candidate : best),
    undefined,
  );
  const searchable = `${details}\n${errorCodes(error)}\n${candidates.map((candidate) => candidate.message).join("\n")}`;
  const matchingRule = alertMessageRules.find((rule) => rule.matches.test(searchable));
  if (matchingRule) {
    candidates.push({
      message: matchingRule.message(details),
      score: 120,
      kind: "guidance",
      details: bestExtractedCandidate?.message,
    });
  }

  const bestCandidate = candidates.reduce<AlertMessageCandidate | undefined>(
    (best, candidate) => (best === undefined || candidate.score > best.score ? candidate : best),
    undefined,
  );
  if (bestCandidate?.kind === "guidance") {
    const heading = alertHeading(context, rawDetails);
    return [
      heading,
      bestCandidate.message,
      bestCandidate.details ? `Error details:\n${bestCandidate.details}` : undefined,
      `Debug details:\n${details}`,
    ]
      .filter((section): section is string => !!section)
      .join("\n\n");
  }
  if (bestCandidate && bestCandidate.message !== rawDetails) {
    const heading = alertHeading(context, rawDetails);
    return heading ? `${heading}\n\n${bestCandidate.message}` : bestCandidate.message;
  }
  return details;
}

/**
 * Candidate future seams, subject to consumer-usage review: browser context/watermark,
 * SharePoint reads and permissions, data transforms/version diffs, MSAL/API calls,
 * and list/folder/file persistence. Keep this facade intact until consumers are checked.
 */
export class SharepointChoiceUtils {
  public readonly context: string;
  public sp: SPFI;

  // attempt to establish correct context url for the site from one of the available sources then setup logging for this class
  public constructor(context?: string) {
    const pageContextWindow = window as SharePointChoiceWindow;
    // if no mock or null or empty or incomplete context try to get this from the current page url
    this.context =
      context ||
      pageContextWindow._spPageContextInfo?.webAbsoluteUrl ||
      document.location.href.replace(/(\/SitePages\/|\/Pages\/|\/_layouts\/|\/Lists\/|#|\?).*$/i, "");

    // ensure full url
    if (!this.context.startsWith("https://"))
      this.context = document.location.origin + "/" + this.context.replace(/^\//, "");

    this.context = this.context.replace(/\/$/, "");

    this.sp = spfi().using(SPBrowser({ baseUrl: this.context }));

    // Best effort for consumers that may use classic context later; this is not a readiness signal.
    void this.mockClassicContext(!!context).catch((error: unknown) => {
      console.error("SharepointChoiceUtils.mockClassicContext initialization error", error);
    });

    this.watermark();
  }

  private watermark() {
    if (App.Release != "LIVE" && document.getElementById("sp-environment-watermark") == null) {
      const watermarkStyle = document.createElement("style");
      watermarkStyle.id = "sp-environment-watermark";
      watermarkStyle.innerHTML = `[ng-version]::before{content: '${Array(11).fill(App.Release).join(" - ")}';}`;
      document.head.appendChild(watermarkStyle);
    }
  }

  private async mockClassicContext(overwrite: boolean = false) {
    const pageContextWindow = window as SharePointChoiceWindow;
    // no classic sp context then mock one up, or if a sprecific context passed in then overwrite if different
    if (
      !pageContextWindow._spPageContextInfo ||
      (overwrite && pageContextWindow._spPageContextInfo.webAbsoluteUrl != this.context)
    )
      pageContextWindow._spPageContextInfo = {
        webAbsoluteUrl: this.context,
      };
    // no user in context
    if (!pageContextWindow._spPageContextInfo.userLoginName) {
      const web = await this.sp.web();
      const currentUser = await this.sp.web.currentUser();
      pageContextWindow._spPageContextInfo = {
        userLoginName: currentUser.LoginName,
        userDisplayName: currentUser.Title,
        userEmail: currentUser.Email,
        userId: currentUser.Id,
        webAbsoluteUrl: this.context,
        webTitle: web.Title,
      };
    }
  }

  // Get the current user and permissions to a flat object for easier use in [disabled]="permission['']" etc.
  // NOTE: this will only detect direct assignments or users added to a mail enabled global security group.
  // Any lookup failure returns an invalid user id and no permissions so consumers fail closed.
  public async permissions(): Promise<SharepointChoicePermission> {
    const pageContextWindow = window as SharePointChoiceWindow;

    const userPermissions: SharepointChoicePermission = {
      userId: 0,
      perms: {},
    };

    let webTitle: string;
    try {
      webTitle = (await this.sp.web()).Title;
      const contextUserId = pageContextWindow._spPageContextInfo?.userId;
      userPermissions.userId =
        typeof contextUserId == "number" && contextUserId > 0 ? contextUserId : (await this.sp.web.currentUser()).Id;
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions web/user lookup error", e);
      return userPermissions;
    }
    if (userPermissions.userId <= 0) return userPermissions;

    try {
      // get any directly assigned groups
      // this doesnt work well with ad and aad groups assignments
      const userGroups = await this.sp.web.currentUser.groups();
      userGroups.forEach((group) => {
        userPermissions.perms[group.LoginName] = true;
        if (group.LoginName.startsWith(`${webTitle} `))
          userPermissions.perms[group.LoginName.replace(`${webTitle} `, "")] = true;
      });
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions group lookup error", e);
      return { userId: 0, perms: {} };
    }

    // ad and aad groups within sp groups dont always expose groups above
    // this depends on hidden, no crawl list with specific permissions assigned to the same list item title and created by SHAREPOINT\System Account
    try {
      const securityList = await this.sp.web.lists.getByTitle("Security")();
      if (securityList.Hidden && securityList.IsApplicationList) {
        const securityEntries = await this.sp.web.lists.getByTitle("Security").items.select("Title").top(5000)();
        securityEntries.forEach((entry) => {
          userPermissions.perms[entry.Title] = true;
          if (entry.Title.startsWith(`${webTitle} `))
            userPermissions.perms[entry.Title.replace(`${webTitle} `, "")] = true;
        });
      }
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions security list lookup error", e);
      return { userId: 0, perms: {} };
    }

    return userPermissions;
  }

  // check permission against object
  public async hasPermission(securableObject: SharePointSecurable, permissions: PermissionKind[]): Promise<boolean> {
    try {
      const effectivePermissions = await securableObject.getCurrentUserEffectivePermissions();
      for (const requiredPermission of permissions) {
        if (this.sp.web.hasPermissions(effectivePermissions, requiredPermission)) return true;
      }
    } catch (e) {
      // permission request failed so return false below
      console.error("SharepointChoiceUtils.hasPermission lookup error", e);
    }
    return false;
  }

  // search
  public async search(
    query: string,
    limit: number = 1000,
    page: number = 1,
    sort?: ISort[],
    select?: string[],
    detail?: string[],
    filter?: string[],
  ): Promise<SearchResults> {
    return await this.sp.search(<ISearchQuery>{
      Querytext: query,
      RowLimit: limit,
      StartRow: (page - 1) * limit,
      EnableInterleaving: true,
      TrimDuplicates: false,
      HitHighlightedProperties: detail ? detail : ["Title", "Path", "Content"],
      SelectProperties: select ? select : ["Title", "Path", "Author", "Editor", "Created", "LastModifiedTime"],
      EnableSorting: true,
      SortList: sort ? sort : [{ Property: "LastModifiedTime", Direction: 1 }],
      Refiners: (detail ? detail : ["Title", "Path"]).join(","),
      RefinementFilters: filter ? filter : [],
    });
  }

  // get list fields in the appropriate format for use in <sharepoint-choice spec=""> attributes
  public async fields(listTitle: string): Promise<SharepointChoiceList> {
    const fieldDefinitions: SharepointChoiceList = {};

    try {
      // even though the main fields are in the selection not all are returned such as Format, so parse the SchemaXml for the rest
      const selectedFieldNames =
        "Title,InternalName,TypeAsString,Scope,Required,Choices,MaxLength,Description,DisplayFormat,AppendOnly,SelectionGroup,Format,FillInChoice,RichText,ReadOnlyField,DefaultValue,SchemaXml".split(
          ",",
        );
      const listFields = await this.sp.web.lists.getByTitle(listTitle).fields.select(...selectedFieldNames)();
      listFields.forEach((field) => {
        if (field.SchemaXml) {
          const fieldValues = field as typeof field & Record<string, unknown>;
          const schemaAttributes = new DOMParser().parseFromString(field.SchemaXml, "text/xml").documentElement
            .attributes;
          Array.from(schemaAttributes).forEach((attribute) => {
            if (!fieldValues[attribute.name]) {
              const value = attribute.value;
              fieldValues[attribute.name] =
                value == "TRUE"
                  ? true
                  : value == "FALSE"
                    ? false
                    : !isNaN(parseFloat(value)) && isFinite(Number(value))
                      ? parseFloat(value)
                      : value;
            }
          });
          // prevent reparsing anywhere else
          field.SchemaXml = "";
        }
        // override scope to current context as it will be used for cross site apps
        field.Scope = this.context;
        fieldDefinitions[field.InternalName] = field as SharepointChoiceField;
      });
    } catch (e) {
      // This mock fallback is required by consuming-app unit tests without a SharePoint list.
      console.error("SharepointChoiceUtils.fields schema lookup error", e);
      fieldDefinitions["Title"] = {
        TypeAsString: "Text",
        InternalName: "Title",
        MaxLength: 16,
        Description: "Tooltip",
        Scope: this.context,
      };
    }

    return fieldDefinitions;
  }

  private async cleanLoadKeys(itemData: SharepointChoiceForm, listTitle?: string, itemId?: number) {
    // exclude these fields from data entirely
    const excludedFields = [
      "IsCurrentVersion",
      "VersionId",
      "File_x005f_x0020_x005f_Type",
      "owshiddenversion",
      "Order",
      "Last_x005f_x0020_x005f_Modified",
      "Created_x005f_x0020_x005f_Date",
      "ParentUniqueId",
      "ScopeId",
      "NoExecute",
      "ContentVersion",
      "SMLastModifiedDate",
      "SMTotalFileStreamSize",
      "WorkflowVersion",
      "ProgId",
      "MetaInfo",
      "Restricted",
      "OriginatorId",
      "AccessPolicy",
      "SortBehavior",
    ];

    for (const fieldName in itemData) {
      // excludedFields
      if (excludedFields.includes(fieldName)) {
        delete itemData[fieldName];
        continue;
      }

      // people fields return twice
      if (
        fieldName.endsWith("StringId") &&
        (itemData[fieldName.replace(/StringId$/, "Id")] || itemData[fieldName.replace(/StringId$/, "Id")] === null)
      ) {
        delete itemData[fieldName];
        continue;
      }

      // remove odata. prefixed
      if (fieldName.startsWith("odata.") || fieldName == "__metadata") {
        delete itemData[fieldName];
        continue;
      }

      // if there are attachments start loading
      if (fieldName == "Attachments" && listTitle && itemId) {
        if (itemData[fieldName])
          // this will be a boolean off the sp api so coerce truethiness and get results if true
          itemData[fieldName] = {
            results: await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).attachmentFiles(),
          };
        else itemData[fieldName] = { results: [] };
        continue;
      }

      // blank is null
      if (itemData[fieldName] === "") itemData[fieldName] = null;

      // dont process nulls or blanks
      if (itemData[fieldName] == null) continue;

      // return multifields to results, old behaviour for old people fields and to prevent json paring clashing
      const fieldValue = itemData[fieldName];
      if (Array.isArray(fieldValue) && fieldValue.length > 0) {
        itemData[fieldName] = {
          results: fieldValue.filter((value) => value && !String(value).includes("\n")),
          __metadata: {
            type: typeof fieldValue[0] == "number" ? "Collection(Edm.Int32)" : "Collection(Edm.String)",
          },
        };
        continue;
      }

      // version history objects have name fields as this
      if (isRecord(fieldValue) && fieldValue["LookupValue"]) {
        itemData[fieldName] = fieldValue["LookupValue"];
        continue;
      }

      // dates and date times
      const dateValue = String(fieldValue);
      if (
        /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9](:[0-6][0-9](\.[0-9]+)?)?(Z|(\+|-)[0-1]?[0-9](:?[0-5][0-9])?)?$/.test(
          dateValue,
        ) ||
        /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/.test(dateValue)
      ) {
        itemData[fieldName] = new Date(dateValue);
        continue;
      }

      // parse objects within text fields for looped data
      try {
        if (typeof fieldValue !== "string") continue;
        const firstCharacter = fieldValue.trim().substring(0, 1);
        if (
          (firstCharacter == "{" || firstCharacter == "[") &&
          fieldValue.trim().endsWith(firstCharacter == "{" ? "}" : "]")
        ) {
          itemData[fieldName] = JSON.parse(fieldValue);
          itemData[fieldName] = this.parseLoop(itemData[fieldName]);
          continue;
        }
      } catch (e) {
        // no further parsable dates
        console.error("SharepointChoiceUtils.cleanLoadKeys parse fallback", e, fieldName);
      }
    }
  }

  // load list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async data(id: number, listTitle: string): Promise<SharepointChoiceForm> {
    try {
      const response: unknown = await this.sp.web.lists.getByTitle(listTitle).items.getById(id)();
      if (!isRecord(response)) throw new TypeError(`SharePoint returned an invalid item for list ${listTitle}.`);
      const itemData: SharepointChoiceForm = response;
      await this.cleanLoadKeys(itemData, listTitle, id);
      return itemData;
    } catch (e) {
      alert(MapErrors(e, "Error loading:"));
      throw e;
    }
  }

  public async version(
    id: number,
    listTitle: string,
    spec: SharepointChoiceList | null = null,
  ): Promise<SharepointChoiceForm[]> {
    // SharePoint returns versions in reverse order; this endpoint request is capped at 5,000.
    const response: unknown = await this.sp.web.lists.getByTitle(listTitle).items.getById(id).versions.top(5000)();
    if (
      !Array.isArray(response) ||
      !response.every(
        (version: unknown): version is SharePointItemVersion =>
          isRecord(version) && typeof version["Created"] === "string" && typeof version["Modified"] === "string",
      )
    ) {
      throw new TypeError(`SharePoint returned invalid version data for list ${listTitle}, item ${id}.`);
    }
    const itemVersions: SharePointItemVersion[] = response;

    // exclude these fields
    const excludedFields = [
      "ID",
      "Created",
      "Author",
      "Editor",
      "Modified",
      "FileRef",
      "FileDirRef",
      "FileLeafRef",
      "ItemChildCount",
      "FolderChildCount",
      "Attachments",
      "FSObjType",
      "VersionLabel",
      "GUID",
      "UniqueId",
    ];

    const normalize = (value: unknown): unknown => {
      if (value instanceof Date) return value.toJSON();
      if (Array.isArray(value)) return value.map(normalize);
      if (value !== null && typeof value == "object") {
        const record = value as Record<string, unknown>;
        return Object.keys(record)
          .sort()
          .reduce<Record<string, unknown>>((normalized, key) => {
            normalized[key] = normalize(record[key]);
            return normalized;
          }, {});
      }
      return value;
    };

    const format = (value: unknown): unknown => {
      if (value === null || typeof value != "object") return value;
      if (value instanceof Date) return value.toLocaleString();
      const results = (value as Record<string, unknown>)["results"];
      if (Array.isArray(results))
        return results
          .map(format)
          .sort((a, b) => String(a).localeCompare(String(b)))
          .join(", ");
      return JSON.stringify(normalize(value));
    };

    // versions return reverse order, so process from oldest to newest
    let versionIndex = itemVersions.length - 1;
    let previousVersion: SharePointItemVersion | null = null;
    // SharePoint includes the current item as a version, so the returned collection is non-empty.
    do {
      const currentVersion = itemVersions[versionIndex]!;
      // these dates are always utc when they lack z or +
      currentVersion.Created += currentVersion.Created.includes("Z") || currentVersion.Created.includes("+") ? "" : "Z";
      currentVersion.Modified +=
        currentVersion.Modified.includes("Z") || currentVersion.Modified.includes("+") ? "" : "Z";

      await this.cleanLoadKeys(currentVersion);
      currentVersion.ChangedFields = [];

      for (const fieldName of Object.keys(currentVersion)) {
        if ((spec && !spec[fieldName]) || excludedFields.includes(fieldName)) continue;

        const currentValue = format(currentVersion[fieldName]);
        const previousValue = previousVersion ? format(previousVersion[fieldName]) : null;

        if (fieldName == "ChangedFields") continue;
        if (currentValue == previousValue) continue;

        currentVersion.ChangedFields.push({
          field: spec ? spec[fieldName]?.Title || fieldName : fieldName,
          from: previousValue,
          to: currentValue,
        });
      }

      previousVersion = currentVersion;
      versionIndex--;
    } while (versionIndex >= 0);

    // output history with exisiting version info plus the changes for each version for use in history tabs
    return itemVersions;
  }

  private parseLoop(value: unknown): unknown {
    try {
      if (typeof value == "string") {
        if (
          /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9]:*[0-9]*\.*[0-9]*Z*$/.test(value) ||
          /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/.test(value)
        ) {
          return new Date(value);
        }
      } else if (value !== null && typeof value == "object") {
        try {
          const record = value as Record<string, unknown>;
          for (const propertyName of Object.keys(record)) {
            record[propertyName] = this.parseLoop(record[propertyName]);
          }
        } catch (e) {
          console.error("SharepointChoiceUtils.parseLoop nested value error", e);
        }
      }
    } catch (e) {
      // silently continue
    }
    return value;
  }

  // use mappings to determine the api to call, then call it with the correct parameters, with the ability to override to localhost api
  public async msalApi<TResponse = unknown>(
    serverRelativeEndPoint: string,
    tokenRole: string,
    httpMethod: string = "GET",
    jsonPostData: unknown = null,
    dataType: string = "json",
    environment: string = App.Release,
  ): Promise<TResponse | null> {
    // use mappings to determine the api to call
    const endpoint = serverRelativeEndPoint.replace(/^\//, "");
    const api = (App.ApiMap || {})[endpoint.split("/")[0]!.toLowerCase()];
    const apiServer = api?.server;
    const configuredClientId = api?.[environment];
    const fallbackClientId = api?.["DEV"];
    const clientId =
      typeof configuredClientId === "string" && configuredClientId
        ? configuredClientId
        : typeof fallbackClientId === "string" && fallbackClientId
          ? fallbackClientId
          : undefined;
    const apiToken = apiServer
      ? (App.ApiToken[apiServer]?.[environment] ?? App.ApiToken[apiServer]?.["DEV"] ?? "")
      : "";
    const apiBaseUrl = apiServer ? App.ApiServers[apiServer]?.[environment] : undefined;

    return (await this.callApi(
      App.Tenancy,
      clientId,
      `${apiToken}${api?.name}/${tokenRole}`,
      endpoint.split("/").length == 1
        ? undefined
        : environment == "LOCAL" || !apiBaseUrl
          ? `https://localhost:${api?.port || 44301}/${endpoint}`
          : `${apiBaseUrl}/${endpoint}`,
      httpMethod,
      jsonPostData,
      dataType,
    )) as TResponse | null;
  }

  // calls an api more generically, or graph api if no parameters passed
  public callApi(
    tenancyOnMicrosoft?: string,
    clientId?: string,
    permissionScope?: string,
    apiUrl?: string,
    httpMethod?: string,
    jsonPostData?: unknown,
    dataType?: "json",
  ): Promise<unknown | null>;
  public callApi(
    tenancyOnMicrosoft: string | undefined,
    clientId: string | undefined,
    permissionScope: string | undefined,
    apiUrl: string | undefined,
    httpMethod: string | undefined,
    jsonPostData: unknown,
    dataType: "text",
  ): Promise<string | null>;
  public callApi(
    tenancyOnMicrosoft: string | undefined,
    clientId: string | undefined,
    permissionScope: string | undefined,
    apiUrl: string | undefined,
    httpMethod: string | undefined,
    jsonPostData: unknown,
    dataType: "buffer",
  ): Promise<ArrayBuffer | null>;
  public callApi(
    tenancyOnMicrosoft: string | undefined,
    clientId: string | undefined,
    permissionScope: string | undefined,
    apiUrl: string | undefined,
    httpMethod: string | undefined,
    jsonPostData: unknown,
    dataType: "none",
  ): Promise<Response | null>;
  public callApi(
    tenancyOnMicrosoft: string | undefined,
    clientId: string | undefined,
    permissionScope: string | undefined,
    apiUrl: string | undefined,
    httpMethod: string | undefined,
    jsonPostData: unknown,
    dataType: string,
  ): Promise<unknown>;
  public async callApi(
    tenancyOnMicrosoft?: string,
    clientId?: string,
    permissionScope?: string,
    apiUrl?: string,
    httpMethod?: string,
    jsonPostData?: unknown,
    dataType: string = "json",
  ): Promise<unknown> {
    const office = (window as SharePointChoiceWindow).Office;
    const isNestedAppAuthSupported = !!office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1");

    const msalOptions = {
      auth: {
        clientId: clientId || App.GraphClient,
        authority: `https://login.microsoftonline.com/${tenancyOnMicrosoft || App.Tenancy}.onmicrosoft.com`,
        redirectUri: this.context?.replace(/\/$/, ""),
      },
      cache: {
        cacheLocation: "localStorage",
        cacheRetentionDays: 1,
      },
      system: {
        allowRedirectInIframe: true,
        iframeBridgeTimeout: 45000, // this needs to load the sp page so will be slower in bridge mode
        popupBridgeTimeout: 180000, // this needs to allow for user interaction in some cases
        redirectNavigationTimeout: 240000, // this needs to allow for user interaction in some cases
      },
    };

    // init client
    const msalApplication = isNestedAppAuthSupported
      ? await createNestablePublicClientApplication(msalOptions)
      : new PublicClientApplication(msalOptions);

    let tokenResponse: AuthenticationResult;

    if (!isNestedAppAuthSupported) {
      await msalApplication.initialize();
      // Completing any pending redirect flow clears stale temporary auth state.
      try {
        await msalApplication.handleRedirectPromise();
      } catch (e) {
        // await may have nothing to do
        console.error("SharepointChoiceUtils.callApi handleRedirectPromise error", e);
      }
    }

    // permission settings
    const tokenRequest = {
      scopes: permissionScope
        ? [permissionScope]
        : apiUrl?.includes("/messages/") || apiUrl?.includes("/attachments/") || apiUrl?.includes("/mailfolders/")
          ? ["Mail.Read", "Mail.Read.Shared"]
          : ["User.Read"],
      account: msalApplication.getAllAccounts()[0],
    };

    // attempt to get token or login and get token
    try {
      tokenResponse = await msalApplication.acquireTokenSilent(tokenRequest);
    } catch (e) {
      try {
        if (isNestedAppAuthSupported) {
          tokenResponse = await msalApplication.acquireTokenPopup({ scopes: tokenRequest.scopes });
          tokenRequest.account = tokenResponse.account || msalApplication.getAllAccounts()[0];
          if (!tokenResponse.accessToken && tokenRequest.account) {
            tokenResponse = await msalApplication.acquireTokenSilent(tokenRequest);
          }
        } else {
          const loginResult = await msalApplication.loginPopup(tokenRequest);
          tokenRequest.account = loginResult.account || msalApplication.getAllAccounts()[0];
          tokenResponse = await msalApplication.acquireTokenSilent(tokenRequest);
        }
      } catch (inner) {
        if (this.isMsalInteractionInProgress(inner)) {
          const staleKeys = this.getMsalInteractionKeys();
          // Intentionally aggressive recovery: remove stale interaction state before asking the user to retry.
          this.clearMsalInteractionState();
          alert(MapErrors(inner, "Error logging in:"));
          throw `MSAL interaction_in_progress detected. Cleared stale keys (${staleKeys.length}${staleKeys.length > 0 ? `: ${staleKeys.join(", ")}` : ""})`;
        }

        alert(MapErrors(inner, "Error logging in:"));
        throw `Exception acquiring token silently or via popup for tenant ${tenancyOnMicrosoft || App.Tenancy} with error ${inner}`;
      }
    }

    // if no url, login only, then return
    if (!apiUrl) return null;

    // query api
    let response: Response | undefined;
    try {
      response = await fetch(apiUrl, {
        method: httpMethod || "GET",
        headers: {
          Authorization: `Bearer ${tokenResponse.accessToken}`,
          "Content-Type": dataType == "json" ? "application/json" : "",
        },
        body: jsonPostData ? (JSON.stringify(jsonPostData) ?? null) : null,
      });

      // return formatted data for 2xx, 4xx and 5xx will not return
      if (response.status == 204) return null;
      if (response.status < 200 || response.status > 299) throw "Exception";

      if (dataType == "json") {
        // Endpoint response schemas vary; keep JSON unknown until the consuming feature validates it.
        const responseData: unknown = await response.json();
        return responseData;
      }
      if (dataType == "text") return await response.text();
      if (dataType == "buffer") return await response.arrayBuffer();
      if (dataType && dataType != "none") return new Blob([await response.arrayBuffer()], { type: dataType });

      return response;
    } catch (e) {
      throw `Exception getting API data with status ${response?.status ?? "no response"} response ${e} and body ${response && typeof response.text === "function" ? await response.text().catch(() => "<unreadable body>") : (response?.body ?? "")}`;
    }
  }

  private isMsalInteractionInProgress(error: unknown): boolean {
    const errorRecord = typeof error === "object" && error !== null ? (error as Record<string, unknown>) : undefined;
    const code = String(errorRecord?.["errorCode"] || errorRecord?.["code"] || "").toLowerCase();
    const message = String(errorRecord?.["errorMessage"] || errorRecord?.["message"] || error || "").toLowerCase();
    return code.includes("interaction_in_progress") || message.includes("interaction_in_progress");
  }

  private getMsalInteractionKeys(): string[] {
    const keys: string[] = [];
    const collect = (storage?: Storage) => {
      if (!storage) return;
      for (let storageIndex = 0; storageIndex < storage.length; storageIndex++) {
        const key = storage.key(storageIndex);
        if (!key) continue;
        if (/^msal\./i.test(key) && /interaction|request|urlhash|state|nonce/i.test(key)) keys.push(key);
      }
    };

    try {
      collect(window.sessionStorage);
    } catch (e) {
      // no cleanup here needed
      console.error("SharepointChoiceUtils.getMsalInteractionKeys sessionStorage access error", e);
    }
    try {
      collect(window.localStorage);
    } catch (e) {
      // no cleanup here needed
      console.error("SharepointChoiceUtils.getMsalInteractionKeys localStorage access error", e);
    }

    return keys.filter((k, i, a) => a.indexOf(k) === i);
  }

  private clearMsalInteractionState(): void {
    const clear = (storage?: Storage) => {
      if (!storage) return;

      const remove: string[] = [];
      for (let storageIndex = 0; storageIndex < storage.length; storageIndex++) {
        const key = storage.key(storageIndex);
        if (!key) continue;
        if (/^msal\./i.test(key) && /interaction|request|urlhash|state|nonce/i.test(key)) remove.push(key);
      }

      remove.forEach((key) => storage.removeItem(key));
    };

    try {
      clear(window.sessionStorage);
    } catch (e) {
      // no cleanup here needed
      console.error("SharepointChoiceUtils.clearMsalInteractionState sessionStorage access error", e);
    }
    try {
      clear(window.localStorage);
    } catch (e) {
      // no cleanup here needed
      console.error("SharepointChoiceUtils.clearMsalInteractionState localStorage access error", e);
    }
  }

  private cleanSaveKeys(saveData: SharepointChoiceForm, unchangedData?: SharepointChoiceForm): void {
    const originalData = unchangedData ?? {};

    // mutate the save object directly
    delete saveData["$$hashKey"];
    delete saveData["Id"];
    delete saveData["ID"];
    delete saveData["Attachments"];

    for (const fieldName of Object.keys(saveData)) {
      if (saveData[fieldName] === "") saveData[fieldName] = null;

      // keep fields becoming blank
      if (saveData[fieldName] === null && originalData[fieldName] !== null) continue;

      // remove and unedited, including internal fields
      if (
        originalData[fieldName] !== undefined &&
        JSON.stringify(originalData[fieldName]) == JSON.stringify(saveData[fieldName])
      ) {
        delete saveData[fieldName];
        continue;
      }

      const fieldValue = saveData[fieldName];
      if (fieldValue === null || fieldValue === undefined) continue;

      // convert dates
      if (fieldValue instanceof Date) {
        saveData[fieldName] = fieldValue.toJSON();
        if (saveData[fieldName] == "Invalid Date") saveData[fieldName] = null;

        if (saveData[fieldName] == null && !originalData[fieldName]) delete saveData[fieldName];
        continue;
      }

      if (Array.isArray(fieldValue)) {
        saveData[fieldName] = JSON.stringify(fieldValue);
        continue;
      }

      if (!isRecord(fieldValue)) continue;

      // if Url
      if (
        Object.prototype.hasOwnProperty.call(fieldValue, "Url") ||
        Object.prototype.hasOwnProperty.call(fieldValue, "Description")
      ) {
        if (!fieldValue["Description"]) fieldValue["Description"] = fieldValue["Url"];
        if (!fieldValue["Url"]) saveData[fieldName] = null;

        if (saveData[fieldName] == null && !originalData[fieldName]) delete saveData[fieldName];
        continue;
      }

      // if Geolocation
      if (
        Object.prototype.hasOwnProperty.call(fieldValue, "Latitude") ||
        Object.prototype.hasOwnProperty.call(fieldValue, "Longitude")
      ) {
        const parseCoordinate = (coordinate: unknown): number | null => {
          if (coordinate === "" || coordinate === undefined || coordinate === null) return null;
          const parsed = typeof coordinate === "number" ? coordinate : Number.parseFloat(String(coordinate));
          return Number.isFinite(parsed) ? parsed : null;
        };
        const latitude = parseCoordinate(fieldValue["Latitude"]);
        const longitude = parseCoordinate(fieldValue["Longitude"]);

        if (latitude === null || longitude === null) saveData[fieldName] = null;
        else saveData[fieldName] = { Latitude: latitude, Longitude: longitude };

        if (saveData[fieldName] == null && !originalData[fieldName]) delete saveData[fieldName];
        continue;
      }

      // convert back to direct array and ensure no nulls selected, should never occur but does on some browsers? and deduplicate data
      if (Object.prototype.hasOwnProperty.call(fieldValue, "results")) {
        const results = Array.isArray(fieldValue["results"]) ? fieldValue["results"] : [];
        const selectedValues = results.filter(
          (selectedValue): selectedValue is string | number =>
            !!selectedValue && (typeof selectedValue === "string" || typeof selectedValue === "number"),
        );
        const uniqueValues = [...new Set(selectedValues)];
        saveData[fieldName] = uniqueValues.length > 0 ? uniqueValues : null;

        if (saveData[fieldName] == null && !originalData[fieldName]) delete saveData[fieldName];
        continue;
      }

      // convert JSON
      saveData[fieldName] = JSON.stringify(saveData[fieldName]);
      continue;
    }
  }

  private hasData(itemData: SharepointChoiceForm): boolean {
    for (const fieldName in itemData) if (fieldName != "Id") return true;

    return false;
  }

  private normalizePath(path: string): string {
    if (!path) return "";
    const protocolSeparatorIndex = path.indexOf("://");
    if (protocolSeparatorIndex >= 0) path = path.substring(path.indexOf("/", protocolSeparatorIndex + 4));
    return decodeURIComponent(path).replace(/\/$/, "");
  }

  private extractInlineImages(save: SharepointChoiceForm): InlineImageExtraction {
    const uploads: InlineImageUpload[] = [];
    const htmlByField = new Map<string, string>();
    let oversizedCount = 0;

    for (const [field, value] of Object.entries(save)) {
      if (typeof value !== "string" || !/<img\b/i.test(value)) continue;

      const fieldUploadCount = uploads.length;
      const html = value.replace(/<img\b[^>]*>/gi, (imageTag) => {
        const source = /\bsrc\s*=\s*(["'])(data:image\/[^"']+)\1/i.exec(imageTag);
        if (!source) return imageTag;

        const data = /^data:(image\/[a-z0-9.+-]+);base64,([a-z0-9+/]*={0,2})$/i.exec(source[2]!);
        if (!data || data[1]!.toLowerCase() === "image/svg+xml") return imageTag;

        const mimeType = data[1]!.toLowerCase();
        const base64 = data[2]!;
        const padding = base64.endsWith("==") ? 2 : base64.endsWith("=") ? 1 : 0;
        const byteLength = Math.floor((base64.length * 3) / 4) - padding;
        if (byteLength > MAX_PASTED_IMAGE_SIZE_BYTES) {
          oversizedCount++;
          return "";
        }

        let binary: string;
        try {
          binary = atob(base64);
        } catch {
          throw new Error(`Unable to decode pasted image data in field ${field}.`);
        }
        const bytes = new Uint8Array(binary.length);
        for (let index = 0; index < binary.length; index++) bytes[index] = binary.charCodeAt(index);

        const extension =
          mimeType === "image/jpeg"
            ? "jpg"
            : mimeType
                .substring("image/".length)
                .split("+")[0]!
                .replace(/[^a-z0-9]/g, "") || "img";
        const placeholder = `sharepoint-choice-inline-image-${crypto.randomUUID()}`;
        uploads.push({
          field,
          fileName: `pasted-image-${crypto.randomUUID()}.${extension}`,
          data: bytes.buffer,
          placeholder,
        });

        return imageTag.replace(source[0], `src="${placeholder}"`);
      });

      if (uploads.length > fieldUploadCount) {
        htmlByField.set(field, html);
        delete save[field];
      } else if (html !== value) {
        save[field] = html;
      }
    }

    return { uploads, htmlByField, oversizedCount };
  }

  private escapeHtmlAttribute(value: string): string {
    return value.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }

  // patch save list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async save(
    formDataIncIdToUpdate: SharepointChoiceForm,
    uneditedDataToBuildPatch: SharepointChoiceForm,
    listTitle: string,
  ): Promise<number> {
    const saveData = JSON.parse(JSON.stringify(formDataIncIdToUpdate)) as SharepointChoiceForm;
    const initialItemId = saveData["Id"];
    let itemId: number | undefined = undefined;
    if (typeof initialItemId === "number") itemId = initialItemId;
    else if (typeof initialItemId === "string" && initialItemId.trim()) {
      const parsedItemId = Number(initialItemId);
      if (Number.isFinite(parsedItemId)) itemId = parsedItemId;
    }
    let inlineImages: InlineImageExtraction;

    try {
      inlineImages = this.extractInlineImages(saveData);
      if (inlineImages.oversizedCount) {
        alert(
          `${inlineImages.oversizedCount === 1 ? "A pasted image was" : `${inlineImages.oversizedCount} pasted images were`} skipped because each image must be ${MAX_PASTED_IMAGE_SIZE_MB} MB or smaller.`,
        );
      }

      this.cleanSaveKeys(saveData, uneditedDataToBuildPatch);

      // save/update the item
      if (!itemId) {
        const createdItem = await this.sp.web.lists.getByTitle(listTitle).items.add(saveData);
        itemId = createdItem.Id;
      } else if (this.hasData(saveData)) {
        await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update(saveData);
      }
    } catch (e) {
      alert(MapErrors(e, "Error saving item:"));
      throw e;
    }

    const errors: string[] = [];
    // process attachments as deletes then uploads
    if (
      itemId &&
      formDataIncIdToUpdate.Attachments &&
      formDataIncIdToUpdate.Attachments.results &&
      formDataIncIdToUpdate.Attachments.results.length > 0
    ) {
      const attachmentNamesToDelete = formDataIncIdToUpdate.Attachments.results
        .filter((attachment: SharepointChoiceAttachment) => attachment.Deleted)
        .map((attachment: SharepointChoiceAttachment) => attachment.FileName);

      for (const attachmentName of attachmentNamesToDelete)
        try {
          await this.sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemId)
            .attachmentFiles.getByName(attachmentName)
            .delete();
        } catch (e) {
          errors.push(
            `Error deleting attachment ${attachmentName} for item ${itemId} in list ${listTitle} with error ${MapErrors(e)}`,
          );
        }

      const attachmentsToAdd = formDataIncIdToUpdate.Attachments.results.filter(
        (attachment: SharepointChoiceAttachment) =>
          !attachment.Deleted && !attachment.ServerRelativeUrl && !attachment.Uploaded,
      );

      for (const attachment of attachmentsToAdd) {
        if (attachment.Data == undefined) continue;
        try {
          const addedAttachment = await this.sp.web.lists
            .getByTitle(listTitle)
            .items.getById(itemId)
            .attachmentFiles.add(attachment.FileName, attachment.Data);
          attachment.Uploaded = true;
          delete attachment.Data;
          attachment.ServerRelativeUrl = (await addedAttachment.file()).ServerRelativeUrl;
        } catch (e) {
          errors.push(
            `Error adding attachment ${attachment.FileName} for item ${itemId} in list ${listTitle} with error ${MapErrors(e)}`,
          );
        }
      }
    }

    const imageErrors: string[] = [];
    const imageHtmlByField = new Map(inlineImages.htmlByField);
    for (const image of inlineImages.uploads) {
      try {
        if (!itemId) throw new Error("Cannot upload pasted images without a saved list item.");

        const added = await this.sp.web.lists
          .getByTitle(listTitle)
          .items.getById(itemId)
          .attachmentFiles.add(image.fileName, image.data);
        const serverRelativeUrl = (await added.file()).ServerRelativeUrl;
        if (!serverRelativeUrl) throw new Error(`SharePoint did not return a URL for ${image.fileName}.`);

        const html = imageHtmlByField.get(image.field);
        if (html === undefined) throw new Error(`Unable to locate the rich-text content for ${image.field}.`);
        imageHtmlByField.set(image.field, html.replace(image.placeholder, this.escapeHtmlAttribute(serverRelativeUrl)));
      } catch (e) {
        imageErrors.push(
          `Error adding pasted image ${image.fileName} for item ${itemId} in list ${listTitle} with error ${MapErrors(e)}`,
        );
      }
    }

    if (inlineImages.uploads.length > 0 && imageErrors.length === 0) {
      const imageFieldValues: SharepointChoiceForm = {};
      for (const [field, html] of imageHtmlByField) imageFieldValues[field] = html;
      try {
        if (itemId === undefined) throw new Error("Cannot save pasted image links without a saved list item.");
        await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).update(imageFieldValues);
      } catch (e) {
        alert(MapErrors(e, "Error saving pasted image links:"));
        throw e;
      }
    }
    errors.push(...imageErrors);

    if (errors.length > 0) {
      const errorMessages = errors.join("\n");
      alert(`Error saving attachments:\n\n${errorMessages}`);
      throw errorMessages;
    }

    if (itemId === undefined) throw new Error(`Unable to determine item id for list ${listTitle}`);
    return itemId;
  }

  public async saveFolderListItem(
    formDataIncIdToUpdate: SharepointChoiceForm,
    uneditedDataToBuildPatch: SharepointChoiceForm,
    listTitle: string,
    folderPath: string,
    additionalProps?: { leafName?: string; objectType?: 0 | 1 | 2 },
  ): Promise<{ id: number; folderPath: string }> {
    folderPath = this.normalizePath(folderPath);

    if (!formDataIncIdToUpdate["Id"]) {
      const folderItemData = JSON.parse(JSON.stringify(formDataIncIdToUpdate)) as SharepointChoiceForm;

      try {
        this.cleanSaveKeys(folderItemData);

        // return the data in the format required for folder/child items
        const simpleFieldUpdates = Object.keys(folderItemData)
          .map((fieldName) => {
            let fieldValue = folderItemData[fieldName];
            // exclude complex objects for 2nd update to catch
            if (typeof fieldValue == "object") return null;
            else if (typeof fieldValue == "string" && /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]/.test(fieldValue))
              return null;
            else if (fieldName.endsWith("Id")) return null;
            // define values as string required for the api
            else if (fieldValue === null) fieldValue = "";
            else if (typeof fieldValue == "boolean") fieldValue = fieldValue ? "1" : "0";
            else if (typeof fieldValue == "number") fieldValue = fieldValue.toString();
            else if (typeof fieldValue == "string") fieldValue = fieldValue.trim();
            if (typeof fieldValue !== "string") return null;
            return {
              FieldName: fieldName,
              FieldValue: fieldValue,
            };
          })
          .filter((item) => item !== undefined && item !== null);

        const validationResults = (await this.sp.web.lists
          .getByTitle(listTitle)
          .addValidateUpdateItemUsingPath(simpleFieldUpdates, folderPath, false, undefined, additionalProps)) as
          | ListItemValidationResult[]
          | ListItemValidationResponse;

        const fieldUpdateResults = Array.isArray(validationResults)
          ? validationResults
          : (validationResults as ListItemValidationResponse).value || [];
        let itemId = fieldUpdateResults.find((result) => result.ItemId)?.ItemId;

        if (!itemId) {
          const idFieldValue = fieldUpdateResults.find((result) => result.FieldName == "Id")?.FieldValue;
          if (idFieldValue !== undefined && idFieldValue !== "") itemId = parseInt(idFieldValue, 10);
        }

        formDataIncIdToUpdate["Id"] = itemId;
      } catch (e: unknown) {
        if (!(e instanceof Error ? e.message : String(e)).includes("already exists") || !additionalProps?.leafName) {
          alert(MapErrors(e, "Error saving folder:"));
          throw e;
        }

        const existingItemPath = `${folderPath}/${additionalProps.leafName}`;
        const existingItems = await this.sp.web.lists
          .getByTitle(listTitle)
          .items.filter(`FileRef eq '${existingItemPath.replace(/'/g, "''")}'`)
          .select("Id")
          .top(1)();

        formDataIncIdToUpdate["Id"] = existingItems[0]?.Id;
      }

      // purposly blank the unedited here so all fields are rewritten later
      uneditedDataToBuildPatch = {};
    }

    if (!formDataIncIdToUpdate["Id"]) {
      const error = new Error(`Unable to determine folder item id for list ${listTitle}`);
      console.error("SharepointChoiceUtils.saveFolderListItem missing item id", error);
      throw error;
    }
    const itemId = await this.save(formDataIncIdToUpdate, uneditedDataToBuildPatch, listTitle);
    const savedItem = await this.sp.web.lists.getByTitle(listTitle).items.getById(itemId).select("FileRef")();

    return {
      id: itemId,
      folderPath: this.normalizePath(savedItem?.FileRef),
    };
  }

  // get query parameters, not strictly sharepoint but reused a lot
  public param(parameterToReturn: string): string | undefined {
    const parameterPattern = new RegExp(`[?&]${parameterToReturn}=([^&]+).*$`);
    const matchedParameter = document.location.search.match(parameterPattern);
    return !matchedParameter ? undefined : decodeURIComponent(matchedParameter[1]!).replace(/\+/g, " ");
  }

  public async ensurePath(path: string, start: number): Promise<void> {
    path = this.normalizePath(path);

    const currentFolderPath = path
      .split("/")
      .slice(0, start + 1)
      .join("/");
    const folderQuery = this.sp.web.getFolderByServerRelativePath(currentFolderPath);
    try {
      const folderInfo = await folderQuery();
      if (!folderInfo.Exists)
        await this.sp.web
          .getFolderByServerRelativePath(path.split("/").slice(0, start).join("/"))
          .addSubFolderUsingPath(path.split("/").slice(start)[0]!);
    } catch (e) {
      console.warn("SharepointChoiceUtils.ensurePath folder lookup failed; attempting to create it", e);
      await this.sp.web
        .getFolderByServerRelativePath(path.split("/").slice(0, start).join("/"))
        .addSubFolderUsingPath(path.split("/").slice(start)[0]!);
    }
    if (currentFolderPath != path) await this.ensurePath(path, start + 1);
  }

  public async getRoot(listTitle: string): Promise<string> {
    const rootFolder = await this.sp.web.lists.getByTitle(listTitle).rootFolder();
    return rootFolder.ServerRelativeUrl;
  }

  public async getFiles(path: string, additionalPath: string | undefined): Promise<SharepointChoiceAttachment[]> {
    path = this.normalizePath(path);

    const folderFiles = await this.sp.web
      .getFolderByServerRelativePath(path + (additionalPath ? "/" + additionalPath : ""))
      .files.orderBy("TimeCreated")
      .expand("ListItemAllFields")();

    return await Promise.all(
      folderFiles.map(async (file) => {
        const expandedFile = file as typeof file & { ListItemAllFields: SharepointChoiceForm };
        const listItemAllFields = expandedFile.ListItemAllFields;
        await this.cleanLoadKeys(listItemAllFields);

        return {
          FileName: file.Name,
          TimeCreated: new Date(file.TimeCreated),
          ServerRelativeUrl: file.ServerRelativeUrl,
          ListItemAllFields: listItemAllFields,
          OldListItemAllFields: JSON.parse(JSON.stringify(listItemAllFields)),
        };
      }),
    );
  }

  public async relocateFolder(source: string, destination: string): Promise<string | null> {
    // ensure these are server relative paths
    const destinationPath = this.normalizePath(destination);
    const sourcePath = this.normalizePath(source);

    // if the destination folder is the same as the current then return null
    if (sourcePath.toLowerCase().replace(/\/$/, "") == destinationPath.toLowerCase().replace(/\/$/, "")) return null;

    // move the files to the new folder by making the parent then moving the folder directly
    const parentPath = destinationPath.substring(0, destinationPath.lastIndexOf("/"));
    await this.ensurePath(parentPath, !this.context || this.context.length < 2 ? 2 : 4);

    // move files, keep both may result in a renamed folder if the destination already exists
    const movedFolder = await this.sp.web.getFolderByServerRelativePath(sourcePath).moveByPath(destinationPath, {
      KeepBoth: true,
      RetainEditorAndModifiedOnMove: true,
      ShouldBypassSharedLocks: true,
    });

    // return where its been moved
    return decodeURIComponent((await movedFolder()).ServerRelativeUrl || destination);
  }

  public async saveFiles(
    path: string,
    additionalPath: string | undefined,
    requestUrl: { Url: string; Description: string } | undefined,
    attachments: { results: SharepointChoiceAttachment[] },
    metadata: SharepointChoiceForm | undefined,
  ): Promise<void> {
    path = this.normalizePath(path);

    // Common metadata is intentionally applied after per-file values and overrides matching fields.
    const commonMetadata: SharepointChoiceForm = metadata
      ? (JSON.parse(JSON.stringify(metadata)) as SharepointChoiceForm)
      : {};
    if (requestUrl && requestUrl.Url) {
      if (!requestUrl.Description) requestUrl.Description = requestUrl.Url;
      commonMetadata["Request"] = requestUrl;
    }

    const errors: string[] = [];
    try {
      if (metadata || requestUrl) {
        const folderItem = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folderItem.update(commonMetadata);
      }

      // subfolders for these
      if (additionalPath) {
        path += "/" + additionalPath;
        const additionalFolderItem = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await additionalFolderItem.update(commonMetadata);
      }
    } catch (e) {
      errors.push("Error saving folder metadata " + MapErrors(e));
    }

    // process saves and deletes
    for (const file of attachments.results) {
      try {
        if (!file.ListItemAllFields) file.ListItemAllFields = {};

        // clone common metadata for files
        for (const fieldName of Object.keys(commonMetadata)) {
          file.ListItemAllFields[fieldName] = commonMetadata[fieldName];
        }

        // basic list item fields cleanup
        delete file.ListItemAllFields["$$hashKey"];
        delete file.ListItemAllFields["Id"];
        delete file.ListItemAllFields["ID"];

        this.cleanSaveKeys(file.ListItemAllFields, file.OldListItemAllFields);

        if (file.Deleted) {
          // file to delete
          await this.sp.web.getFileByServerRelativePath(file.ServerRelativeUrl || path + "/" + file.FileName).recycle();
        } else if (file.Data) {
          // file to upload
          const uploadedFile = await this.sp.web
            .getFolderByServerRelativePath(path)
            .files.addUsingPath(file.FileName, file.Data, {
              EnsureUniqueFileName: true,
            });
          // Record upload success before metadata updates so a retry never uploads the file again.
          file.TimeCreated = new Date();
          file.ServerRelativeUrl = uploadedFile.ServerRelativeUrl;
          delete file.Data;
          if (this.hasData(file.ListItemAllFields)) {
            const uploadedFileItem = await this.sp.web
              .getFileByServerRelativePath(uploadedFile.ServerRelativeUrl)
              .getItem();
            await uploadedFileItem.update(file.ListItemAllFields);
          }
        } else if (this.hasData(file.ListItemAllFields)) {
          // get current item and check for changes
          const existingFileItem = await this.sp.web
            .getFileByServerRelativePath(file.ServerRelativeUrl || path + "/" + file.FileName)
            .getItem();
          await existingFileItem.update(file.ListItemAllFields);
        }
      } catch (e) {
        errors.push(`Error saving file ${file.FileName} in folder ${path} with error ${MapErrors(e)}`);
      }
    }

    if (errors.length > 0) {
      const errorMessages = errors.join("\n");
      alert(`Error saving file:\n\n${errorMessages}`);
      throw errorMessages;
    }
  }
}
