import { ChangeDetectorRef, ElementRef, ErrorHandler } from "@angular/core";
import { Workbook } from "devextreme-exceljs-fork";
import JSZip from "@turbowarp/jszip";
import { afterEach, beforeAll, beforeEach, describe, expect, it, vi } from "vitest";
import {
  SharepointChoiceColumn,
  SharepointChoiceForm,
  SharepointChoiceRow,
  SharepointChoiceTabs,
} from "./sharepoint-choice.models";
import type { SharepointChoiceComponent as SharepointChoiceComponentType } from "./sharepoint-choice.component";
import type { SharepointChoiceTable as SharepointChoiceTableType } from "./sharepoint-choice.table";
import type { SharepointChoiceUtils as SharepointChoiceUtilsType } from "./sharepoint-choice.utils";
import { SharepointChoiceRefresh } from "./sharepoint-choice.refresh";

class TestChangeDetectorRef extends ChangeDetectorRef {
  override detach(): void {}
  override detectChanges(): void {}
  override markForCheck(): void {}
  override reattach(): void {}
}

const tables: SharepointChoiceTableType[] = [];
const components: SharepointChoiceComponentType[] = [];
const storage = new Map<string, string>();
let SharepointChoiceTableCtor: typeof SharepointChoiceTableType;
let SharepointChoiceComponentCtor: typeof SharepointChoiceComponentType;
let SharepointChoiceUtilsCtor: typeof SharepointChoiceUtilsType;
let MapErrors: typeof import("./sharepoint-choice.utils").MapErrors;

function createTable(): SharepointChoiceTableType {
  const table = new SharepointChoiceTableCtor(new TestChangeDetectorRef(), new ErrorHandler());
  tables.push(table);
  return table;
}

function createComponent(): SharepointChoiceComponentType {
  const component = new SharepointChoiceComponentCtor(
    new ElementRef(document.createElement("app-choice")),
    new TestChangeDetectorRef(),
    new SharepointChoiceRefresh(),
    new ErrorHandler(),
  );
  components.push(component);
  return component;
}

beforeAll(async () => {
  storage.clear();
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => ({
      ok: true,
      json: async () => ({ environmentMappings: [] }),
    })),
  );
  vi.stubGlobal("localStorage", {
    getItem: vi.fn((key: string) => storage.get(key) ?? null),
    setItem: vi.fn((key: string, value: string) => storage.set(key, value)),
  });
  SharepointChoiceTableCtor = (await import("./sharepoint-choice.table")).SharepointChoiceTable;
  SharepointChoiceComponentCtor = (await import("./sharepoint-choice.component")).SharepointChoiceComponent;
  const utils = await import("./sharepoint-choice.utils");
  SharepointChoiceUtilsCtor = utils.SharepointChoiceUtils;
  MapErrors = utils.MapErrors;
}, 60_000);

beforeEach(() => {
  storage.clear();
  vi.stubGlobal(
    "fetch",
    vi.fn(async () => ({
      ok: true,
      json: async () => ({ environmentMappings: [] }),
    })),
  );
  vi.stubGlobal("localStorage", {
    getItem: vi.fn((key: string) => storage.get(key) ?? null),
    setItem: vi.fn((key: string, value: string) => storage.set(key, value)),
  });
});

afterEach(() => {
  components.splice(0).forEach((component) => component.ngOnDestroy());
  tables.splice(0).forEach((table) => table.ngOnDestroy());
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("SharepointChoiceLogging", () => {
  it("uses the original Error.stack in Grafana logs", async () => {
    vi.useFakeTimers();
    vi.spyOn(console, "error").mockImplementation(() => {});

    const error = new Error("Native stack trace test");
    const { SharepointChoiceLogging } = await import("./sharepoint-choice.logging");
    const logger = new SharepointChoiceLogging();
    logger.handleError(error);

    const state = (
      window as typeof window & {
        GrafanaLogging?: { values: Array<[string, string]>; flushTimer: number | undefined };
      }
    ).GrafanaLogging;
    const entry = state?.values
      .map(([, line]) => JSON.parse(line) as { Message?: string; StackTrace?: string })
      .find((value) => value.Message === error.message);

    expect(entry?.StackTrace).toBe(error.stack);
  });
});

describe("SharepointChoiceTable", () => {
  it("excludes null, undefined, and empty values from upper-bound filters", () => {
    const table = createTable();
    table.allData = {
      Records: [{ Value: null }, { Value: undefined }, { Value: "" }, { Value: 2 }, { Value: 3 }],
    };
    table.filter = { Records: { Value: { less: 3 } } };

    expect(table.rows("Records").map((row) => row["Value"])).toEqual([2, 3]);
  });

  it("reuses cached empty row results", () => {
    const table = createTable();
    table.allData = { Records: [{ Value: "one" }, { Value: "two" }] };
    table.filter = { Records: { Value: { equals: "missing" } } };
    const fieldValue = vi.spyOn(table, "fieldValue");

    expect(table.rows("Records")).toEqual([]);
    expect(table.rows("Records")).toEqual([]);

    expect(fieldValue).toHaveBeenCalledTimes(2);
  });

  it("caches distinct filter values until the tab data changes", () => {
    const table = createTable();
    table.allData = { Records: [{ Value: "two" }, { Value: "one" }] };
    table.selectedTab = "Records";
    const fieldValue = vi.spyOn(table, "fieldValue");

    expect(table.filterDistinct("Value")).toEqual(["one", "two"]);
    expect(table.filterDistinct("Value")).toEqual(["one", "two"]);
    expect(fieldValue).toHaveBeenCalledTimes(2);

    table.allData = { Records: [{ Value: "three" }] };
    expect(table.filterDistinct("Value")).toEqual(["three"]);
    expect(fieldValue).toHaveBeenCalledTimes(3);
  });

  it("uses row tracking keys rather than page positions", () => {
    const table = createTable();
    table.allData = { Records: [{ Id: 1 }, { Id: 2 }] };
    const [first, second] = table.allData["Records"]!;

    expect(table.trackRow(1, first!)).toBe("id-1");
    expect(table.trackRow(0, second!)).toBe("id-2");
  });

  it("does not reuse rendered content across tabs with matching row IDs", () => {
    const table = createTable();
    const firstRow: SharepointChoiceRow = { Id: 1, Value: "first" };
    const secondRow: SharepointChoiceRow = { Id: 1, Value: "second" };
    const data: SharepointChoiceTabs = { First: [firstRow], Second: [secondRow] };
    const column = { field: "Value", cellRenderer: (value: unknown) => String(value) };
    table.allData = data;
    table.allCols = [column];

    table.selectedTab = "First";
    table.rows("First");
    expect(table.renderNode(column, firstRow["Value"], firstRow, 0, 0)).toBe("first");

    table.selectedTab = "Second";
    table.rows("Second");
    expect(table.renderNode(column, secondRow["Value"], secondRow, 0, 0)).toBe("second");

    table.selectedTab = "First";
    table.rows("First");
    expect(table.renderNode(column, firstRow["Value"], firstRow, 0, 0)).toBe("first");
  });

  it("invalidates rendered content when column renderers change", () => {
    const table = createTable();
    const row: SharepointChoiceRow = { Id: 1, Value: "value" };
    const oldColumn = { field: "Value", cellRenderer: () => "old" };
    table.allData = { Records: [row] };
    table.allCols = [oldColumn];
    table.selectedTab = "Records";
    table.rows("Records");
    expect(table.renderNode(oldColumn, row["Value"], row, 0, 0)).toBe("old");

    const newColumn = { field: "Value", cellRenderer: () => "new" };
    table.allCols = [newColumn];
    expect(table.renderNode(newColumn, row["Value"], row, 0, 0)).toBe("new");
  });

  it("normalizes invalid page sizes and accepts numeric string values", () => {
    const table = createTable();
    vi.spyOn(console, "error").mockImplementation(() => {});

    table.pageSize = -10;
    expect(table.pageSize).toBe(100);

    table.pageSize = 12.8;
    expect(table.pageSize).toBe(12);

    table.pageSize = "25";
    expect(table.pageSize).toBe(25);
  });

  it("skips malformed rows and columns while reporting their input paths", () => {
    const table = createTable();
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});

    table.allData = {
      Records: [{ Value: "valid" }, null],
      Invalid: 42,
    } as unknown as SharepointChoiceTabs;
    table.allCols = [{ field: 12 }, { field: "Value" }] as unknown as SharepointChoiceColumn[];

    expect(table.allData["Records"]).toHaveLength(1);
    expect(table.allData["Invalid"]).toBeUndefined();
    expect(table.allCols).toEqual([{ field: "Value" }]);
    expect(consoleError).toHaveBeenCalledTimes(3);
  });

  it("normalizes malformed persisted preferences before table rendering", () => {
    const table = createTable();
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
    table.prefix = "malformed-preferences";
    storage.set(
      "SharepointTable-malformed-preferences",
      JSON.stringify({
        Sort: { Records: [{ field: 12, direction: "ascending" }] },
        Filter: { Records: { Value: { contains: 7 } } },
        Order: { Records: 7 },
        Hide: { Records: 7 },
        Tab: "Records",
        Size: 25,
      }),
    );

    table.ngOnInit();
    table.allCols = [{ field: "Value" }];

    expect(() => table.fields("Records")).not.toThrow();
    expect(() => table.rows("Records")).not.toThrow();
    expect(consoleError).toHaveBeenCalled();
  });

  it("keeps table actions working when local storage is unavailable", () => {
    const table = createTable();
    vi.spyOn(console, "error").mockImplementation(() => {});
    vi.stubGlobal("localStorage", {
      getItem: vi.fn(() => {
        throw new Error("storage blocked");
      }),
      setItem: vi.fn(() => {
        throw new Error("storage blocked");
      }),
    });

    expect(() => (table.selectedTab = "Records")).not.toThrow();
    expect(() => (table.pageSize = 25)).not.toThrow();
  });

  it("removes active resize listeners when the table is destroyed", () => {
    const table = createTable();
    const target = document.createElement("span");
    const parent = document.createElement("th");
    parent.append(target);
    vi.spyOn(parent, "offsetWidth", "get").mockReturnValue(120);
    const removeDocumentListener = vi.spyOn(document, "removeEventListener");
    const removeWindowListener = vi.spyOn(window, "removeEventListener");
    const event = {
      preventDefault: vi.fn(),
      stopPropagation: vi.fn(),
      pageX: 100,
      target,
    } as unknown as MouseEvent;

    table.startResize(event, { field: "Name" });
    table.ngOnDestroy();

    expect(removeDocumentListener).toHaveBeenCalledWith("mousemove", expect.any(Function));
    expect(removeDocumentListener).toHaveBeenCalledWith("mouseup", expect.any(Function));
    expect(removeWindowListener).toHaveBeenCalledWith("blur", expect.any(Function));
  });

  it("puts the configured export title above headers and delays object URL cleanup", async () => {
    const table = createTable();
    const anchor = document.createElement("a");
    const anchorClick = vi.spyOn(anchor, "click");
    const nativeCreateElement = document.createElement.bind(document);
    vi.spyOn(document, "createElement").mockImplementation((tagName: string) =>
      tagName === "a" ? anchor : nativeCreateElement(tagName),
    );
    let exportBlob: Blob | undefined;
    vi.spyOn(URL, "createObjectURL").mockImplementation((blob: Blob | MediaSource) => {
      if (!(blob instanceof Blob)) throw new Error("Export did not create a blob");
      exportBlob = blob;
      return "blob:test";
    });
    const revokeObjectURL = vi.spyOn(URL, "revokeObjectURL").mockImplementation(() => {});
    table.allData = { Records: [{ Name: "Row" }] };
    table.allCols = [{ field: "Name" }];
    table.selectedTab = "Records";
    table.export = { title: "Report title" };

    await table.exportTable([{ Name: "Row" }]);

    expect(revokeObjectURL).not.toHaveBeenCalled();
    expect(anchorClick).toHaveBeenCalledOnce();
    if (!exportBlob) throw new Error("Export did not create a workbook blob");

    const exportedWorkbook = new Workbook();
    await exportedWorkbook.xlsx.load(await exportBlob.arrayBuffer());
    const worksheet = exportedWorkbook.worksheets[0]!;
    expect(worksheet.getCell(1, 1).value).toBe("Report title");
    expect(worksheet.getCell(2, 1).value).toBe("Name");
    expect(worksheet.getCell(3, 1).value).toBe("Row");

    await new Promise((resolve) => setTimeout(resolve, 1100));
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:test");
  });

  it("reports rejected export callbacks to the error handler", async () => {
    const error = new Error("Export source failed");
    const errorHandler = new ErrorHandler();
    const handleError = vi.spyOn(errorHandler, "handleError").mockImplementation(() => {});
    const table = new SharepointChoiceTableCtor(new TestChangeDetectorRef(), errorHandler);
    tables.push(table);
    table.allData = { Records: [{ Value: "Row" }] };
    table.allCols = [{ field: "Value" }];
    table.selectedTab = "Records";
    table.export = {
      sourceRows: async () => {
        throw error;
      },
    };

    await expect(table.exportTable(table.rows("Records"))).resolves.toBeUndefined();
    expect(handleError).toHaveBeenCalledWith(error);
  });
});

describe("SharepointChoiceComponent", () => {
  it("extracts EML attachments without filenames under generated names", async () => {
    const component = createComponent();
    vi.spyOn(component, "changed").mockImplementation(() => {});
    const results: Parameters<typeof component.appendFile>[2] = [];
    const eml = [
      "From: sender@example.com",
      "To: recipient@example.com",
      "Subject: Missing attachment name",
      "MIME-Version: 1.0",
      'Content-Type: multipart/mixed; boundary="test-boundary"',
      "",
      "--test-boundary",
      "Content-Type: application/octet-stream",
      "Content-Disposition: attachment",
      "Content-Transfer-Encoding: base64",
      "",
      "SGVsbG8=",
      "--test-boundary--",
      "",
    ].join("\r\n");

    await component.emls(new TextEncoder().encode(eml).buffer as ArrayBuffer, results, "message.eml");

    expect(results).toHaveLength(1);
    expect(results[0]?.FileName).toMatch(/^Attachment-[0-9a-f-]{36}\.bin$/i);
    expect(results[0]?.Length).toBe(5);
  });

  it("extracts ZIP entries without usable names under generated names", async () => {
    const component = createComponent();
    vi.spyOn(component, "changed").mockImplementation(() => {});
    const results: Parameters<typeof component.appendFile>[2] = [];
    const zip = new JSZip();
    zip.file("   ", "payload");
    const data = (await zip.generateAsync({ type: "arraybuffer" })) as ArrayBuffer;

    await component.zips(data, results, "archive.zip");

    expect(results).toHaveLength(1);
    expect(results[0]?.FileName).toMatch(/^Attachment-[0-9a-f-]{36}\.bin$/i);
  });

  it("assigns a generated filename when an extracted attachment has no name", async () => {
    const component = createComponent();
    vi.spyOn(component, "changed").mockImplementation(() => {});
    const results: Parameters<typeof component.appendFile>[2] = [];

    await component.appendFile(undefined, new Uint8Array([1]).buffer, results);

    expect(results).toHaveLength(1);
    expect(results[0]?.FileName).toMatch(/^Attachment-[0-9a-f-]{36}\.bin$/i);
  });

  it("reports the configured file size limit when rejecting an oversized attachment", async () => {
    const component = createComponent();
    const constructor = SharepointChoiceComponentCtor;
    const originalLimit = Object.getOwnPropertyDescriptor(constructor, "maxFileSizeBytes");
    Object.defineProperty(constructor, "maxFileSizeBytes", { value: 1.5 * 1024 * 1024 });

    try {
      await expect(component.appendFile("large.bin", new ArrayBuffer(2 * 1024 * 1024), [])).rejects.toThrow(
        "large.bin exceeds the 1.5 MB file size limit",
      );
    } finally {
      if (originalLimit) Object.defineProperty(constructor, "maxFileSizeBytes", originalLimit);
    }
  });

  it("applies the custom filter with form context and searches choices case-insensitively", () => {
    const component = createComponent();
    component.field = "Status";
    component.form = { Status: "Alpha", Enabled: true };
    component.spec = {
      Status: { TypeAsString: "Choice", Choices: ["Alpha", "Beta", "Gamma"] },
    };
    component.select = {
      filter: (choice: string, _index: number, _choices: string[], context: { form: SharepointChoiceForm }) =>
        context.form["Enabled"] && choice !== "Beta",
    };
    component.otherText = "a";

    expect(component.choices()).toEqual(["Alpha", "Gamma"]);
  });

  it("normalizes multi-choice values and limits disabled fields to selected choices", () => {
    const component = createComponent();
    component.field = "Tags";
    component.form = { Tags: ["Blue"] };
    component.spec = {
      Tags: { TypeAsString: "MultiChoice", Choices: ["Blue", "Green", "Red"] },
    };
    component.disabled = true;

    expect(component.form["Tags"]).toEqual({
      __metadata: { type: "Collection(Edm.String)" },
      results: ["Blue"],
    });
    expect(component.choices()).toEqual(["Blue"]);
  });

  it("maps Outlook errors before aggregating them for display", async () => {
    const component = createComponent();
    const contextWindow = window as typeof window & {
      _spPageContextInfo?: { webAbsoluteUrl: string; userLoginName?: string };
    };
    const previousContext = contextWindow._spPageContextInfo;
    contextWindow._spPageContextInfo = {
      webAbsoluteUrl: "https://example.test/sites/test/",
      userLoginName: "test-user",
    };
    const error = {
      title: "One or more validation errors occurred.",
      response: {
        data: {
          errors: {
            Code: ["A broker with this code already exists. Choose a different code."],
          },
        },
      },
    };
    vi.spyOn(SharepointChoiceUtilsCtor.prototype, "callApi").mockRejectedValue(error);
    const showAlert = vi.fn();
    vi.stubGlobal("alert", showAlert);
    const transfer = {
      getData: (type: string) =>
        type === "multimaillistmessagerows"
          ? JSON.stringify({
              mailboxInfos: [{ mailboxSmtpAddress: "test@example.com" }],
              subjects: ["Test message"],
              latestItemIds: ["message-id"],
            })
          : "",
    } as DataTransfer;
    const message =
      "Email append error: Test message.eml with error A broker with this code already exists. Choose a different code.";

    try {
      await expect(component.outlook(transfer)).rejects.toBe(message);
      expect(showAlert).toHaveBeenCalledWith(`Error retreiving item(s):\n\n${message}`);
    } finally {
      contextWindow._spPageContextInfo = previousContext;
    }
  });
});

describe("SharepointChoiceUtils", () => {
  it("extracts a readable message from escaped JSON embedded in an error", () => {
    const message = MapErrors(
      String.raw`Exception getting API data with status 400 response Exception and body {\"title\":\"Can not create a new broker under company 8 with code 1635 because it already exists as broker 1635. (Parameter \\u0027Code\\u0027)\"}`,
    );

    expect(message).toBe(
      "Can not create a new broker under company 8 with code 1635 because it already exists as broker 1635. (Parameter 'Code')",
    );
    expect(message).not.toContain("Exception getting API data");
  });

  it("prefers a specific nested validation error over a generic response title", () => {
    const message = MapErrors({
      title: "One or more validation errors occurred.",
      response: {
        data: {
          errors: {
            Code: ["A broker with this code already exists. Choose a different code."],
          },
        },
      },
    });

    expect(message).toBe("A broker with this code already exists. Choose a different code.");
  });

  it("maps SharePoint file locks to actionable guidance and retains debug details", () => {
    const details =
      'Error saving file Synthetic test document.docx in folder /sites/test/Submission with error Error: Error making HttpClient request in queryable [423] ::> {"odata.error":{"code":"-2147018894, Microsoft.SharePoint.SPException","message":{"value":"The file is locked for shared use by user@example.com."}}}';

    const message = MapErrors(details);

    expect(message).toContain('File "Synthetic test document.docx" is already open for shared use');
    expect(message).toContain("Please close any open tabs or clients using this file");
    expect(message).toContain("Debug details:");
    expect(message).toContain("-2147018894, Microsoft.SharePoint.SPException");
  });

  it("maps common MSAL interaction errors while retaining the original error", () => {
    const error = Object.assign(new Error("Popup interaction_in_progress"), {
      errorCode: "interaction_in_progress",
    });
    error.stack = "Error: Popup interaction_in_progress\n    at acquireTokenPopup (auth.ts:42)";

    const message = MapErrors(error, "Error logging in:");

    expect(message).toContain("A sign-in is already in progress");
    expect(message).toContain("Error details:\nPopup interaction_in_progress");
    expect(message).toContain("Debug details:");
    expect(message).toContain("Popup interaction_in_progress");
    expect(message).not.toContain("at acquireTokenPopup");
  });

  it("does not rewrite unrecognized error messages", () => {
    expect(MapErrors("A custom failure")).toBe("A custom failure");
  });

  it("normalizes SharePoint list item values for the form", async () => {
    const context = "https://example.test/sites/test/";
    const contextWindow = window as typeof window & {
      _spPageContextInfo?: { webAbsoluteUrl: string; userLoginName?: string };
    };
    const previousContext = contextWindow._spPageContextInfo;
    const watermarkWasPresent = document.getElementById("sp-environment-watermark") !== null;
    contextWindow._spPageContextInfo = { webAbsoluteUrl: context, userLoginName: "test-user" };

    try {
      const utils = new SharepointChoiceUtilsCtor();
      const item = {
        Title: "",
        Tags: ["blue", "skip\nmultiline", null, "green"],
        Owner: { LookupValue: "Ada" },
        Metadata: '{"created":"2024-01-02"}',
        VersionId: 7,
        "odata.type": "SP.Data.RequestsListItem",
      };
      const loadItem = vi.fn(async () => item);
      const getById = vi.fn(() => loadItem);
      const getByTitle = vi.fn(() => ({ items: { getById } }));
      utils.sp = { web: { lists: { getByTitle } } } as unknown as typeof utils.sp;

      const form = await utils.data(42, "Requests");

      expect(utils.context).toBe("https://example.test/sites/test");
      expect(getByTitle).toHaveBeenCalledWith("Requests");
      expect(getById).toHaveBeenCalledWith(42);
      expect(form["Title"]).toBeNull();
      expect(form["Tags"]).toEqual({
        results: ["blue", "green"],
        __metadata: { type: "Collection(Edm.String)" },
      });
      expect(form["Owner"]).toBe("Ada");
      expect(form["Metadata"]["created"]).toBeInstanceOf(Date);
      expect(form["VersionId"]).toBeUndefined();
      expect(form["odata.type"]).toBeUndefined();
    } finally {
      if (previousContext) contextWindow._spPageContextInfo = previousContext;
      else delete contextWindow._spPageContextInfo;
      if (!watermarkWasPresent) document.getElementById("sp-environment-watermark")?.remove();
    }
  });

  it("skips unchanged file metadata updates but still saves changed metadata", async () => {
    const contextWindow = window as typeof window & {
      _spPageContextInfo?: { webAbsoluteUrl: string; userLoginName?: string };
    };
    const previousContext = contextWindow._spPageContextInfo;
    const watermarkWasPresent = document.getElementById("sp-environment-watermark") !== null;
    contextWindow._spPageContextInfo = {
      webAbsoluteUrl: "https://example.test/sites/test",
      userLoginName: "test-user",
    };

    try {
      const utils = new SharepointChoiceUtilsCtor();
      const update = vi.fn(async () => undefined);
      const getItem = vi.fn(async () => ({ update }));
      const getFileByServerRelativePath = vi.fn(() => ({ getItem }));
      utils.sp = { web: { getFileByServerRelativePath } } as unknown as typeof utils.sp;

      await utils.saveFiles(
        "/sites/test/Shared Documents",
        undefined,
        undefined,
        {
          results: [
            {
              FileName: "unchanged.docx",
              ServerRelativeUrl: "/sites/test/Shared Documents/unchanged.docx",
              ListItemAllFields: { Title: "Existing title" },
              OldListItemAllFields: { Title: "Existing title" },
            },
            {
              FileName: "changed.docx",
              ServerRelativeUrl: "/sites/test/Shared Documents/changed.docx",
              ListItemAllFields: { Title: "Updated title" },
              OldListItemAllFields: { Title: "Existing title" },
            },
          ],
        },
        undefined,
      );

      expect(getFileByServerRelativePath).toHaveBeenCalledOnce();
      expect(getFileByServerRelativePath).toHaveBeenCalledWith("/sites/test/Shared Documents/changed.docx");
      expect(update).toHaveBeenCalledOnce();
      expect(update).toHaveBeenCalledWith({ Title: "Updated title" });
    } finally {
      if (previousContext) contextWindow._spPageContextInfo = previousContext;
      else delete contextWindow._spPageContextInfo;
      if (!watermarkWasPresent) document.getElementById("sp-environment-watermark")?.remove();
    }
  });
});
