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
import { PublicClientApplication, createNestablePublicClientApplication } from "@azure/msal-browser";
import { PermissionKind } from "@pnp/sp/security";
import { ISearchQuery, ISort, SearchResults } from "@pnp/sp/search";
import { App } from "./sharepoint-choice.settings";
import {
  SharepointChoicePermission,
  SharepointChoiceForm,
  SharepointChoiceList,
  SharepointChoiceField,
  SharepointChoiceAttachment,
} from "./sharepoint-choice.models";

export function MapErrors(error: unknown, context?: string): string {
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
  constructor(context?: string) {
    let w: any = window;
    // if no mock or null or empty or incomplete context try to get this from the current page url
    this.context =
      context ||
      (w._spPageContextInfo ? w._spPageContextInfo.webAbsoluteUrl : null) ||
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
      let environment = document.createElement("style");
      environment.id = "sp-environment-watermark";
      environment.innerHTML = `[ng-version]::before{content: '${Array(11).fill(App.Release).join(" - ")}';}`;
      document.head.appendChild(environment);
    }
  }

  private async mockClassicContext(overwrite: boolean = false) {
    let w: any = window;
    // no classic sp context then mock one up, or if a sprecific context passed in then overwrite if different
    if (!w._spPageContextInfo || (overwrite && w._spPageContextInfo.webAbsoluteUrl != this.context))
      w._spPageContextInfo = {
        webAbsoluteUrl: this.context,
      };
    // no user in context
    if (!w._spPageContextInfo.userLoginName) {
      let web = await this.sp.web();
      let user = await this.sp.web.currentUser();
      w._spPageContextInfo = {
        userLoginName: user.LoginName,
        userDisplayName: user.Title,
        userEmail: user.Email,
        userId: user.Id,
        webAbsoluteUrl: this.context,
        webTitle: web.Title,
      };
    }
  }

  // Get the current user and permissions to a flat object for easier use in [disabled]="permission['']" etc.
  // NOTE: this will only detect direct assignments or users added to a mail enabled global security group.
  // Any lookup failure returns an invalid user id and no permissions so consumers fail closed.
  public async permissions(): Promise<SharepointChoicePermission> {
    let w: any = window;

    let permission: SharepointChoicePermission = {
      userId: 0,
      perms: {},
    };

    let web: string = "";
    try {
      web = (await this.sp.web()).Title;
      let contextUserId = w._spPageContextInfo?.userId;
      permission.userId =
        typeof contextUserId == "number" && contextUserId > 0 ? contextUserId : (await this.sp.web.currentUser()).Id;
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions web/user lookup error", e);
      return permission;
    }
    if (permission.userId <= 0) return permission;

    try {
      // get any directly assigned groups
      // this doesnt work well with ad and aad groups assignments
      let perm = await this.sp.web.currentUser.groups();
      perm.forEach((x) => {
        permission.perms[x.LoginName] = true;
        if (x.LoginName.startsWith(`${web} `)) permission.perms[x.LoginName.replace(`${web} `, "")] = true;
      });
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions group lookup error", e);
      return { userId: 0, perms: {} };
    }

    // ad and aad groups within sp groups dont always expose groups above
    // this depends on hidden, no crawl list with specific permissions assigned to the same list item title and created by SHAREPOINT\System Account
    try {
      let sec = await this.sp.web.lists.getByTitle("Security")();
      if (sec.Hidden && sec.IsApplicationList) {
        let per = await this.sp.web.lists.getByTitle("Security").items.select("Title").top(5000)();
        per.forEach((s) => {
          permission.perms[s.Title] = true;
          if (s.Title.startsWith(`${web} `)) permission.perms[s.Title.replace(`${web} `, "")] = true;
        });
      }
    } catch (e) {
      console.error("SharepointChoiceUtils.permissions security list lookup error", e);
      return { userId: 0, perms: {} };
    }

    return permission;
  }

  // check permission against object
  public async hasPermission(object: any, permissions: PermissionKind[]): Promise<boolean> {
    try {
      let perm = await object.getCurrentUserEffectivePermissions();
      for (let permission of permissions) {
        if (this.sp.web.hasPermissions(perm, permission)) return true;
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
    let spec: SharepointChoiceList = {};

    try {
      // even though the main fields are in the selection not all are returned such as Format, so parse the SchemaXml for the rest
      let selectFields =
        "Title,InternalName,TypeAsString,Scope,Required,Choices,MaxLength,Description,DisplayFormat,AppendOnly,SelectionGroup,Format,FillInChoice,RichText,ReadOnlyField,DefaultValue,SchemaXml".split(
          ",",
        );
      let arr = await this.sp.web.lists.getByTitle(listTitle).fields.select(...selectFields)();
      arr.forEach((x) => {
        if (x.SchemaXml) {
          let s = new DOMParser().parseFromString(x.SchemaXml, "text/xml").documentElement.attributes;
          Array.from(s).reduce((acc, attr) => {
            if (!x[attr.name]) {
              x[attr.name] = attr.value;
              x[attr.name] = x[attr.name] == "TRUE" ? true : x[attr.name] == "FALSE" ? false : x[attr.name];
              x[attr.name] =
                x[attr.name] != null && !isNaN(parseFloat(x[attr.name])) && isFinite(x[attr.name])
                  ? parseFloat(x[attr.name])
                  : x[attr.name];
            }
            return acc;
          }, {});
          // prevent reparsing anywhere else
          x.SchemaXml = "";
        }
        // override scope to current context as it will be used for cross site apps
        x.Scope = this.context;
        spec[x.InternalName] = x as SharepointChoiceField;
      });
    } catch (e) {
      // This mock fallback is required by consuming-app unit tests without a SharePoint list.
      console.error("SharepointChoiceUtils.fields schema lookup error", e);
      spec["Title"] = {
        TypeAsString: "Text",
        InternalName: "Title",
        MaxLength: 16,
        Description: "Tooltip",
        Scope: this.context,
      };
    }

    return spec;
  }

  private async cleanLoadKeys(d: SharepointChoiceForm, listTitle?: string, id?: number) {
    // exclude these fields from data entirely
    let excludedFields = [
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

    for (var key in d) {
      // excludedFields
      if (excludedFields.includes(key)) {
        delete d[key];
        continue;
      }

      // people fields return twice
      if (
        key.endsWith("StringId") &&
        (d[key.replace(/StringId$/, "Id")] || d[key.replace(/StringId$/, "Id")] === null)
      ) {
        delete d[key];
        continue;
      }

      // remove odata. prefixed
      if (key.startsWith("odata.") || key == "__metadata") {
        delete d[key];
        continue;
      }

      // if there are attachments start loading
      if (key == "Attachments" && listTitle && id) {
        if (d[key])
          // this will be a boolean off the sp api so coerce truethiness and get results if true
          d[key] = {
            results: await this.sp.web.lists.getByTitle(listTitle).items.getById(id).attachmentFiles(),
          };
        else d[key] = { results: [] };
        continue;
      }

      // blank is null
      if (d[key] === "") d[key] = null;

      // dont process nulls or blanks
      if (d[key] == null) continue;

      // return multifields to results, old behaviour for old people fields and to prevent json paring clashing
      if (typeof d[key] == "object" && !d[key].results && d[key].length > 0) {
        d[key] = {
          results: d[key].filter((k) => k && !k.toString().includes("\n")),
          __metadata: {
            type: typeof d[key][0] == "number" ? "Collection(Edm.Int32)" : "Collection(Edm.String)",
          },
        };
        continue;
      }

      // version history objects have name fields as this
      if (typeof d[key] == "object" && d[key].LookupValue) {
        d[key] = d[key].LookupValue;
        continue;
      }

      // dates and date times
      let i = d[key].toString();
      if (
        /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9](:[0-6][0-9](\.[0-9]+)?)?(Z|(\+|-)[0-1]?[0-9](:?[0-5][0-9])?)?$/.test(
          i,
        ) ||
        /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/.test(i)
      ) {
        d[key] = new Date(d[key]);
        continue;
      }

      // parse objects within text fields for looped data
      try {
        let f = d[key].toString().trim().substring(0, 1);
        if (
          (f == "{" || f == "[") &&
          d[key]
            .toString()
            .trim()
            .endsWith(f == "{" ? "}" : "]")
        ) {
          d[key] = JSON.parse(d[key]);
          d[key] = this.parseLoop(d[key]);
          continue;
        }
      } catch (e) {
        // no further parsable dates
        console.error("SharepointChoiceUtils.cleanLoadKeys parse fallback", e, key);
      }
    }
  }

  // load list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async data(id: number, listTitle: string): Promise<SharepointChoiceForm> {
    let d: SharepointChoiceForm = {};

    try {
      d = await this.sp.web.lists.getByTitle(listTitle).items.getById(id)();
      await this.cleanLoadKeys(d, listTitle, id);
    } catch (e) {
      alert(MapErrors(e, "Error loading:"));
      throw e;
    }

    return d;
  }

  public async version(
    id: number,
    listTitle: string,
    spec: SharepointChoiceList | null = null,
  ): Promise<SharepointChoiceForm[]> {
    // SharePoint returns versions in reverse order; this endpoint request is capped at 5,000.
    let d = await this.sp.web.lists.getByTitle(listTitle).items.getById(id).versions.top(5000)();

    // exclude these fields
    let excludedFields = [
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

    let normalize = (v: any): any => {
      if (v instanceof Date) return v.toJSON();
      if (Array.isArray(v)) return v.map(normalize);
      if (v && typeof v == "object") {
        return Object.keys(v)
          .sort()
          .reduce(
            (normalized, key) => {
              normalized[key] = normalize(v[key]);
              return normalized;
            },
            {} as Record<string, any>,
          );
      }
      return v;
    };

    let format = (v: any) => {
      if (!v || typeof v != "object") return v;
      if (v instanceof Date) return v.toLocaleString();
      if (Array.isArray(v.results))
        return v.results
          .map(format)
          .sort((a, b) => String(a).localeCompare(String(b)))
          .join(", ");
      return JSON.stringify(normalize(v));
    };

    // versions return reverse order, so process from oldest to newest
    let i = d.length - 1;
    let lastVersion: any = null;
    // SharePoint includes the current item as a version, so the returned collection is non-empty.
    do {
      // these dates are always utc when they lack z or +
      d[i].Created += d[i].Created.includes("Z") || d[i].Created.includes("+") ? "" : "Z";
      d[i].Modified += d[i].Modified.includes("Z") || d[i].Modified.includes("+") ? "" : "Z";

      await this.cleanLoadKeys(d[i]);
      d[i].ChangedFields = [];

      for (const k of Object.keys(d[i])) {
        if ((spec && !spec[k]) || excludedFields.includes(k)) continue;

        let c = format(d[i][k]);
        let f = lastVersion ? format(lastVersion[k]) : null;

        if (c == f) continue;

        d[i].ChangedFields.push({
          field: spec ? spec[k]?.Title || k : k,
          from: f,
          to: c,
        });
      }

      lastVersion = d[i];
      i--;
    } while (i >= 0);

    // output history with exisiting version info plus the changes for each version for use in history tabs
    return d;
  }

  private parseLoop(i: any): any {
    try {
      if (typeof i == "string") {
        if (
          /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9][ T][0-2][0-9]:[0-5][0-9]:*[0-9]*\.*[0-9]*Z*$/.test(i) ||
          /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]$/.test(i)
        ) {
          return new Date(i);
        }
      } else if (typeof i == "object") {
        try {
          for (let a of Object.keys(i)) i[a] = this.parseLoop(i[a]);
        } catch (e) {
          console.error("SharepointChoiceUtils.parseLoop nested value error", e);
        }
      }
    } catch (e) {
      // silently continue
    }
    return i;
  }

  // use mappings to determine the api to call, then call it with the correct parameters, with the ability to override to localhost api
  public async msalApi(
    serverRelativeEndPoint: string,
    tokenRole: string,
    httpMethod: string = "GET",
    jsonPostData: any = null,
    dataType: string = "json",
    environment: string = App.Release,
  ): Promise<any> {
    // use mappings to determine the api to call
    var endPoint = serverRelativeEndPoint.replace(/^\//, "");
    var api = (App.ApiMap || {})[endPoint.split("/")[0]!.toLowerCase()];

    return this.callApi(
      App.Tenancy,
      api?.[environment] || api?.["DEV"],
      `${App.ApiToken?.[api?.server]?.[environment] ?? App.ApiToken?.[api?.server]?.["DEV"] ?? ""}${api?.name}/${tokenRole}`,
      endPoint.split("/").length == 1
        ? undefined
        : environment == "LOCAL" || !App.ApiServers?.[api?.server]?.[environment]
          ? `https://localhost:${api?.port || 44301}/${endPoint}`
          : `${App.ApiServers[api?.server][environment]}/${endPoint}`,
      httpMethod,
      jsonPostData,
      dataType,
    );
  }

  // calls an api more generically, or graph api if no parameters passed
  public async callApi(
    tenancyOnMicrosoft?: string,
    clientId?: string,
    permissionScope?: string,
    apiUrl?: string,
    httpMethod?: string,
    jsonPostData?: any,
    dataType: string = "json",
  ): Promise<any> {
    let office = (window as any).Office;
    let nested = !!office?.context?.requirements?.isSetSupported?.("NestedAppAuth", "1.1");

    var msalOpts = {
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
    var msal = nested ? await createNestablePublicClientApplication(msalOpts) : new PublicClientApplication(msalOpts);

    var login;

    if (!nested) {
      await msal.initialize();
      // Completing any pending redirect flow clears stale temporary auth state.
      try {
        await msal.handleRedirectPromise();
      } catch (e) {
        // await may have nothing to do
        console.error("SharepointChoiceUtils.callApi handleRedirectPromise error", e);
      }
    }

    // permission settings
    var params = {
      scopes: permissionScope
        ? [permissionScope]
        : apiUrl?.includes("/messages/") || apiUrl?.includes("/attachments/") || apiUrl?.includes("/mailfolders/")
          ? ["Mail.Read", "Mail.Read.Shared"]
          : ["User.Read"],
      account: msal.getAllAccounts()[0],
    };

    // attempt to get token or login and get token
    try {
      login = await msal.acquireTokenSilent(params);
    } catch (e) {
      try {
        if (nested) {
          login = await msal.acquireTokenPopup({ scopes: params.scopes });
          params.account = login.account || msal.getAllAccounts()[0];
          if (!login.accessToken && params.account) login = await msal.acquireTokenSilent(params);
        } else {
          let loginResult = await msal.loginPopup(params);
          params.account = loginResult.account || msal.getAllAccounts()[0];
          login = await msal.acquireTokenSilent(params);
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
    var r;
    try {
      r = await fetch(apiUrl, {
        method: httpMethod || "GET",
        headers: {
          Authorization: `Bearer ${login.accessToken}`,
          "Content-Type": dataType == "json" ? "application/json" : "",
        },
        body: jsonPostData ? JSON.stringify(jsonPostData) : null,
      });

      // return formatted data for 2xx, 4xx and 5xx will not return
      if (r.status == 204) return null;
      if (r.status < 200 || r.status > 299) throw "Exception";

      if (dataType == "json") return await r.json();
      if (dataType == "text") return await r.text();
      if (dataType == "buffer") return await r.arrayBuffer();
      if (dataType && dataType != "none") return new Blob([await r.arrayBuffer()], { type: dataType });

      return r;
    } catch (e) {
      throw `Exception getting API data with status ${r?.status ?? "no response"} response ${e} and body ${r && typeof r.text === "function" ? await r.text().catch(() => "<unreadable body>") : (r?.body ?? "")}`;
    }
  }

  private isMsalInteractionInProgress(error: unknown): boolean {
    const raw = error as any;
    const code = `${raw?.errorCode || raw?.code || ""}`.toLowerCase();
    const message = `${raw?.errorMessage || raw?.message || raw || ""}`.toLowerCase();
    return code.includes("interaction_in_progress") || message.includes("interaction_in_progress");
  }

  private getMsalInteractionKeys(): string[] {
    const keys: string[] = [];
    const collect = (storage?: Storage) => {
      if (!storage) return;
      for (let i = 0; i < storage.length; i++) {
        const key = storage.key(i);
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
      for (let i = 0; i < storage.length; i++) {
        const key = storage.key(i);
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

  private cleanSaveKeys(save: SharepointChoiceForm, uned?: SharepointChoiceForm): void {
    if (!uned) uned = {};

    // mutate the save object directly
    delete save["$$hashKey"];
    delete save["Id"];
    delete save["ID"];
    delete save["Attachments"];

    for (let key of Object.keys(save)) {
      if (save[key] === "") save[key] = null;

      // keep fields becoming blank
      if (save[key] === null && uned[key] !== null) continue;

      // remove and unedited, including internal fields
      if (uned[key] !== undefined && JSON.stringify(uned[key]) == JSON.stringify(save[key])) {
        delete save[key];
        continue;
      }

      // prevent errors on nulls or non objects
      if (save[key] === null || save[key] === undefined || typeof save[key] != "object") continue;

      // convert dates
      if (save[key].toJSON) {
        save[key] = save[key].toJSON();
        if (save[key] == "Invalid Date") save[key] = null;

        if (save[key] == null && !uned[key]) delete save[key];
        continue;
      }

      // if Url
      if (
        Object.prototype.hasOwnProperty.call(save[key], "Url") ||
        Object.prototype.hasOwnProperty.call(save[key], "Description")
      ) {
        if (!save[key].Description) save[key].Description = save[key].Url;
        if (!save[key].Url) save[key] = null;

        if (save[key] == null && !uned[key]) delete save[key];
        continue;
      }

      // if Geolocation
      if (
        Object.prototype.hasOwnProperty.call(save[key], "Latitude") ||
        Object.prototype.hasOwnProperty.call(save[key], "Longitude")
      ) {
        let latitude = save[key].Latitude;
        let longitude = save[key].Longitude;

        if (latitude === "" || latitude === undefined) latitude = null;
        else if (typeof latitude == "string") latitude = parseFloat(latitude);

        if (longitude === "" || longitude === undefined) longitude = null;
        else if (typeof longitude == "string") longitude = parseFloat(longitude);

        if (latitude == null || longitude == null || isNaN(latitude) || isNaN(longitude)) save[key] = null;
        else save[key] = { Latitude: latitude, Longitude: longitude };

        if (save[key] == null && !uned[key]) delete save[key];
        continue;
      }

      // convert back to direct array and ensure no nulls selected, should never occur but does on some browsers? and deduplicate data
      if (Object.prototype.hasOwnProperty.call(save[key], "results")) {
        save[key] =
          save[key].results
            ?.filter((i: string | number) => i)
            .filter((item: string | number, pos: number, arr: (string | number)[]) => arr.indexOf(item) == pos) ?? [];
        if (save[key].length == 0) save[key] = null;

        if (save[key] == null && !uned[key]) delete save[key];
        continue;
      }

      // convert JSON
      save[key] = JSON.stringify(save[key]);
      continue;
    }
  }

  private hasData(save: SharepointChoiceForm): boolean {
    for (var key in save) if (key != "Id") return true;

    return false;
  }

  private normalizePath(path: string): string {
    if (!path) return "";
    let s = path.indexOf("://");
    if (s >= 0) path = path.substring(path.indexOf("/", s + 4));
    return decodeURIComponent(path).replace(/\/$/, "");
  }

  // patch save list item data and parse any data types appropriate for use in <sharepoint-choice ngModel=""> attributes
  public async save(
    formDataIncIdToUpdate: SharepointChoiceForm,
    uneditedDataToBuildPatch: SharepointChoiceForm,
    listTitle: string,
  ): Promise<number> {
    var save = JSON.parse(JSON.stringify(formDataIncIdToUpdate));
    let id = save.Id;

    try {
      this.cleanSaveKeys(save, uneditedDataToBuildPatch);

      // save/update the item
      if (!id) {
        var saving = await this.sp.web.lists.getByTitle(listTitle).items.add(save);
        id = saving.Id;
      } else if (this.hasData(save)) {
        await this.sp.web.lists.getByTitle(listTitle).items.getById(id).update(save);
      }
    } catch (e) {
      alert(MapErrors(e, "Error saving item:"));
      throw e;
    }

    var errors: Array<string> = [];
    // process attachments as deletes then uploads
    if (
      id &&
      formDataIncIdToUpdate.Attachments &&
      formDataIncIdToUpdate.Attachments.results &&
      formDataIncIdToUpdate.Attachments.results.length > 0
    ) {
      var deletes = formDataIncIdToUpdate.Attachments.results
        .filter((a: SharepointChoiceAttachment) => {
          return a.Deleted;
        })
        .map((a: SharepointChoiceAttachment) => {
          return a.FileName;
        });

      for (let i = 0; i < deletes.length; i++)
        try {
          await this.sp.web.lists
            .getByTitle(listTitle)
            .items.getById(id)
            .attachmentFiles.getByName(deletes[i]!)
            .delete();
        } catch (e) {
          errors.push(
            `Error deleting attachment ${deletes[i]} for item ${id} in list ${listTitle} with error ${MapErrors(e)}`,
          );
        }

      var adds = formDataIncIdToUpdate.Attachments.results.filter((a: SharepointChoiceAttachment) => {
        return !a.Deleted && !a.ServerRelativeUrl && !a.Uploaded;
      });

      for (let a = 0; a < adds.length; a++) {
        let attachment = adds[a]!;
        if (attachment.Data == undefined) continue;
        try {
          let added = await this.sp.web.lists
            .getByTitle(listTitle)
            .items.getById(id)
            .attachmentFiles.add(attachment.FileName, attachment.Data);
          attachment.Uploaded = true;
          delete attachment.Data;
          attachment.ServerRelativeUrl = (await added.file()).ServerRelativeUrl;
        } catch (e) {
          errors.push(
            `Error adding attachment ${attachment.FileName} for item ${id} in list ${listTitle} with error ${MapErrors(e)}`,
          );
        }
      }
    }

    if (errors.length > 0) {
      let messages = errors.join("\n");
      alert(`Error saving attachments:\n\n${messages}`);
      throw messages;
    }

    return id;
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
      let save = JSON.parse(JSON.stringify(formDataIncIdToUpdate));

      try {
        this.cleanSaveKeys(save, undefined);

        // return the data in the format required for folder/child items
        let simpleSave = Object.keys(save)
          .map((key) => {
            let value = save[key];
            // exclude complex objects for 2nd update to catch
            if (typeof value == "object") return null;
            else if (typeof value == "string" && /^[1920]{2}[0-9]{2}\-[01][0-9]\-[0-3][0-9]/.test(value)) return null;
            else if (key.endsWith("Id")) return null;
            // define values as string required for the api
            else if (value === null) value = "";
            else if (typeof value == "boolean") value = value ? "1" : "0";
            else if (typeof value == "number") value = value.toString();
            else if (typeof value == "string") value = value.trim();
            return {
              FieldName: key,
              FieldValue: value,
            };
          })
          .filter((item) => item !== undefined && item !== null);

        let results = await this.sp.web.lists
          .getByTitle(listTitle)
          .addValidateUpdateItemUsingPath(simpleSave, folderPath, false, undefined, additionalProps);

        let updateResults = Array.isArray(results) ? results : (results as { value?: any[] })?.value || [];
        let itemId = updateResults.find((result) => result.ItemId)?.ItemId;

        if (!itemId) {
          let idField = updateResults.find((result) => result.FieldName == "Id")?.FieldValue;
          if (idField !== undefined && idField !== null && idField !== "") itemId = parseInt(idField, 10);
        }

        formDataIncIdToUpdate["Id"] = itemId;
      } catch (e: any) {
        if (!(e instanceof Error ? e.message : String(e)).includes("already exists") || !additionalProps?.leafName) {
          alert(MapErrors(e, "Error saving folder:"));
          throw e;
        }

        let existingPath = `${folderPath}/${additionalProps.leafName}`;
        let existing = await this.sp.web.lists
          .getByTitle(listTitle)
          .items.filter(`FileRef eq '${existingPath.replace(/'/g, "''")}'`)
          .select("Id")
          .top(1)();

        formDataIncIdToUpdate["Id"] = existing[0]?.Id;
      }

      // purposly blank the unedited here so all fields are rewritten later
      uneditedDataToBuildPatch = {};
    }

    if (!formDataIncIdToUpdate["Id"]) {
      let error = new Error(`Unable to determine folder item id for list ${listTitle}`);
      console.error("SharepointChoiceUtils.saveFolderListItem missing item id", error);
      throw error;
    }
    let id = await this.save(formDataIncIdToUpdate, uneditedDataToBuildPatch, listTitle);
    let item = await this.sp.web.lists.getByTitle(listTitle).items.getById(id).select("FileRef")();

    return {
      id,
      folderPath: this.normalizePath(item?.FileRef),
    };
  }

  // get query parameters, not strictly sharepoint but reused a lot
  public param(parameterToReturn: string): string | undefined {
    var rx = new RegExp(`[?&]${parameterToReturn}=([^&]+).*$`);
    var returnVal = document.location.search.match(rx);
    return !returnVal ? undefined : decodeURIComponent(returnVal[1]!).replace(/\+/g, " ");
  }

  public async ensurePath(path: string, start: number): Promise<void> {
    path = this.normalizePath(path);

    var p = path
      .split("/")
      .slice(0, start + 1)
      .join("/");
    var folder = this.sp.web.getFolderByServerRelativePath(p);
    try {
      var f = await folder();
      if (!f.Exists)
        await this.sp.web
          .getFolderByServerRelativePath(path.split("/").slice(0, start).join("/"))
          .addSubFolderUsingPath(path.split("/").slice(start)[0]!);
    } catch (e) {
      console.warn("SharepointChoiceUtils.ensurePath folder lookup failed; attempting to create it", e);
      await this.sp.web
        .getFolderByServerRelativePath(path.split("/").slice(0, start).join("/"))
        .addSubFolderUsingPath(path.split("/").slice(start)[0]!);
    }
    if (p != path) await this.ensurePath(path, start + 1);
  }

  public async getRoot(list: string): Promise<string> {
    let root = await this.sp.web.lists.getByTitle(list).rootFolder();
    return root.ServerRelativeUrl;
  }

  public async getFiles(path: string, additional: string | undefined): Promise<SharepointChoiceAttachment[]> {
    path = this.normalizePath(path);

    var files = await this.sp.web
      .getFolderByServerRelativePath(path + (additional ? "/" + additional : ""))
      .files.orderBy("TimeCreated")
      .expand("ListItemAllFields")();

    return await Promise.all(
      files.map(async (file) => {
        await this.cleanLoadKeys(file["ListItemAllFields"]);

        return {
          FileName: file.Name,
          TimeCreated: new Date(file.TimeCreated),
          ServerRelativeUrl: file.ServerRelativeUrl,
          ListItemAllFields: file["ListItemAllFields"],
          OldListItemAllFields: JSON.parse(JSON.stringify(file["ListItemAllFields"])),
        };
      }),
    );
  }

  public async relocateFolder(source: string, destination: string): Promise<string | null> {
    // ensure these are server relative paths
    var dst = this.normalizePath(destination);
    var src = this.normalizePath(source);

    // if the destination folder is the same as the current then return null
    if (src.toLowerCase().replace(/\/$/, "") == dst.toLowerCase().replace(/\/$/, "")) return null;

    // move the files to the new folder by making the parent then moving the folder directly
    var parent = dst.substring(0, dst.lastIndexOf("/"));
    await this.ensurePath(parent, !this.context || this.context.length < 2 ? 2 : 4);

    // move files, keep both may result in a renamed folder if the destination already exists
    var folder = await this.sp.web.getFolderByServerRelativePath(src).moveByPath(dst, {
      KeepBoth: true,
      RetainEditorAndModifiedOnMove: true,
      ShouldBypassSharedLocks: true,
    });

    // return where its been moved
    return decodeURIComponent((await folder()).ServerRelativeUrl || destination);
  }

  public async saveFiles(
    path: string,
    additional: string | undefined,
    url: { Url: string; Description: string } | undefined,
    files: { results: SharepointChoiceAttachment[] },
    metadata: SharepointChoiceForm | undefined,
  ): Promise<void> {
    path = this.normalizePath(path);

    // Common metadata is intentionally applied after per-file values and overrides matching fields.
    var commonmeta = metadata ? JSON.parse(JSON.stringify(metadata)) : {};
    if (url && url.Url) {
      if (!url.Description) url.Description = url.Url;
      commonmeta["Request"] = url;
    }

    var errors: Array<string> = [];
    try {
      var folder;
      if (metadata || url) {
        folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folder.update(commonmeta);
      }

      // subfolders for these
      if (additional) {
        path += "/" + additional;
        folder = await this.sp.web.getFolderByServerRelativePath(path).getItem();
        await folder.update(commonmeta);
      }
    } catch (e) {
      errors.push("Error saving folder metadata " + MapErrors(e));
    }

    // process saves and deletes
    for (let i = 0; i < files.results.length; i++) {
      let file = files.results[i]!;
      try {
        if (!file.ListItemAllFields) file.ListItemAllFields = {};

        // clone common metadata for files
        for (let m of Object.keys(commonmeta)) {
          file.ListItemAllFields[m] = commonmeta[m];
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
          var uploaded = await this.sp.web
            .getFolderByServerRelativePath(path)
            .files.addUsingPath(file.FileName, file.Data, {
              EnsureUniqueFileName: true,
            });
          // Record upload success before metadata updates so a retry never uploads the file again.
          file.TimeCreated = new Date();
          file.ServerRelativeUrl = uploaded.ServerRelativeUrl;
          delete file.Data;
          if (this.hasData(file.ListItemAllFields)) {
            let i = await this.sp.web.getFileByServerRelativePath(uploaded.ServerRelativeUrl).getItem();
            await i.update(file.ListItemAllFields);
          }
        } else if (this.hasData(file.ListItemAllFields)) {
          // get current item and check for changes
          let i = await this.sp.web
            .getFileByServerRelativePath(file.ServerRelativeUrl || path + "/" + file.FileName)
            .getItem();
          await i.update(file.ListItemAllFields);
        }
      } catch (e) {
        errors.push(`Error saving file ${file.FileName} in folder ${path} with error ${MapErrors(e)}`);
      }
    }

    if (errors.length > 0) {
      let messages = errors.join("\n");
      alert(`Error saving file:\n\n${messages}`);
      throw messages;
    }
  }
}
