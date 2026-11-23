import { ChangeDetectorRef, ElementRef, ErrorHandler, provideZonelessChangeDetection, type Type } from "@angular/core";
import { TestBed } from "@angular/core/testing";
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

const fixtures: Array<{ destroy(): void }> = [];

async function createFixture<T>(component: Type<T>) {
  await TestBed.configureTestingModule({
    imports: [component],
    providers: [provideZonelessChangeDetection()],
  }).compileComponents();
  const fixture = TestBed.createComponent(component);
  fixtures.push(fixture);
  return fixture;
}

const { mockWeb } = vi.hoisted(() => ({ mockWeb: vi.fn() }));

vi.mock("@pnp/sp/webs", async (importOriginal) => {
  const actual = await importOriginal<typeof import("@pnp/sp/webs")>();
  return { ...actual, Web: mockWeb };
});

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

async function createComponentFixture() {
  await TestBed.configureTestingModule({
    imports: [SharepointChoiceComponentCtor],
  }).compileComponents();
  return TestBed.createComponent(SharepointChoiceComponentCtor);
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
  mockWeb.mockReset();
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
  fixtures.splice(0).forEach((fixture) => fixture.destroy());
  components.splice(0).forEach((component) => component.ngOnDestroy());
  tables.splice(0).forEach((table) => table.ngOnDestroy());
  vi.useRealTimers();
  vi.restoreAllMocks();
  vi.unstubAllGlobals();
});

describe("rich text clipboard images", () => {
  it("embeds pasted raster image files as data URLs", async () => {
    const component = createComponent();
    component.field = "Body";
    component.form = { Body: "" };
    component.spec = { Body: { TypeAsString: "Note", RichText: true } };

    const editor = component.editor;
    expect(editor).toBeDefined();
    if (!editor) throw new Error("Rich text editor was not initialized.");

    const dataUrl = "data:image/png;base64,aGVsbG8=";
    vi.spyOn(FileReader.prototype, "readAsDataURL").mockImplementation(function (this: FileReader) {
      Object.defineProperty(this, "result", { configurable: true, value: dataUrl });
      this.dispatchEvent(new ProgressEvent("load"));
    });

    const item = {
      kind: "file",
      type: "image/png",
      getAsFile: () => new File(["image"], "pasted.png", { type: "image/png" }),
    } as DataTransferItem;
    const preventDefault = vi.fn();
    const event = { clipboardData: { items: [item] }, preventDefault } as unknown as ClipboardEvent;
    const handled = editor.view.someProp("handlePaste", (handlePaste) =>
      handlePaste(editor.view, event, editor.view.state.selection.content()),
    );

    expect(handled).toBe(true);
    expect(preventDefault).toHaveBeenCalledOnce();
    await vi.waitFor(() => {
      let pastedImageSource: string | undefined;
      editor.view.state.doc.descendants((node) => {
        if (node.type.name === "image") pastedImageSource = node.attrs["src"];
      });
      expect(pastedImageSource).toBe(dataUrl);
    });
  });
});

describe("zoneless OnPush rendering", () => {
  it("renders async choice suggestions after markForCheck without an extra detectChanges call", async () => {
    const fixture = await createFixture(SharepointChoiceComponentCtor);
    const component = fixture.componentInstance;
    const suggestion: SharepointChoiceForm = { Contact: "Ada Lovelace" };
    let resolveSearch!: (matches: Array<object>) => void;
    const search = vi.fn(
      () =>
        new Promise<Array<object>>((resolve) => {
          resolveSearch = resolve;
        }),
    );

    component.field = "Contact";
    component.spec = { Contact: { TypeAsString: "Text", Title: "Contact" } };
    component.form = { Contact: "Ada" };
    component.text = { search };
    fixture.detectChanges();

    const pendingSearch = component.onUpTextSearch("Ada");
    resolveSearch([suggestion]);
    await pendingSearch;
    await fixture.whenStable();

    expect(fixture.nativeElement.querySelector(".AutoComplete li")?.textContent?.trim()).toBe("Ada Lovelace");
  });

  it("clears stale choice suggestions when an async search rejects", async () => {
    const fixture = await createFixture(SharepointChoiceComponentCtor);
    const component = fixture.componentInstance;
    const error = new Error("Search failed");
    component.field = "Contact";
    component.spec = { Contact: { TypeAsString: "Text", Title: "Contact" } };
    component.form = { Contact: "Ada" };
    component.text = { search: vi.fn().mockRejectedValue(error) };
    fixture.detectChanges();
    component.results = [{ Contact: "Stale result" }];
    fixture.detectChanges();
    const handleError = vi.spyOn(component["errorHandler"] as ErrorHandler, "handleError").mockImplementation(() => {});

    await component.onUpTextSearch("Ada");
    await fixture.whenStable();

    expect(fixture.nativeElement.querySelector(".AutoComplete")).toBeNull();
    expect(handleError).toHaveBeenCalledWith(error);
  });

  it("keeps only the latest result when overlapping searches finish out of order", async () => {
    const fixture = await createFixture(SharepointChoiceComponentCtor);
    const component = fixture.componentInstance;
    const pending: Array<(matches: Array<object>) => void> = [];
    component.field = "Contact";
    component.spec = { Contact: { TypeAsString: "Text", Title: "Contact" } };
    component.form = { Contact: "Ada" };
    component.text = {
      search: vi.fn(
        () =>
          new Promise<Array<object>>((resolve) => {
            pending.push(resolve);
          }),
      ),
    };
    fixture.detectChanges();

    const firstSearch = component.onUpTextSearch("A");
    const secondSearch = component.onUpTextSearch("Ad");
    pending[1]!([{ Contact: "Latest result" }]);
    await secondSearch;
    pending[0]!([{ Contact: "Outdated result" }]);
    await firstSearch;
    await fixture.whenStable();

    expect(fixture.nativeElement.querySelector(".AutoComplete li")?.textContent?.trim()).toBe("Latest result");
  });

  it("renders input changes made after the initial render without manually checking the view", async () => {
    const fixture = await createFixture(SharepointChoiceComponentCtor);
    const component = fixture.componentInstance;
    component.field = "Title";
    component.spec = { Title: { TypeAsString: "Text", Title: "Title" } };
    component.form = { Title: "Initial" };
    fixture.detectChanges();

    component.form = { Title: "Updated by input setter" };
    await fixture.whenStable();

    expect((fixture.nativeElement.querySelector(".Text input") as HTMLInputElement).value).toBe(
      "Updated by input setter",
    );
  });

  it("updates a table cell after an awaited row callback and ignores malformed rows", async () => {
    const fixture = await createFixture(SharepointChoiceTableCtor);
    const table = fixture.componentInstance;
    const row = { Id: 1, Name: "Before" };
    const consoleError = vi.spyOn(console, "error").mockImplementation(() => {});
    let releaseCallback!: () => void;
    const callbackGate = new Promise<void>((resolve) => {
      releaseCallback = resolve;
    });
    const rowClicked = vi.fn(async (clickedRow: typeof row) => {
      await callbackGate;
      clickedRow.Name = "After async callback";
      return true;
    });

    table.prefix = "zoneless-async-table";
    table.rowClicked = rowClicked;
    table.allData = { Records: [null, row] } as unknown as SharepointChoiceTabs;
    table.allCols = [{ field: "Name" }];
    table.selectedTab = "Records";
    fixture.detectChanges();

    expect(consoleError).toHaveBeenCalled();
    expect(fixture.nativeElement.querySelector('tbody td[role="cell"]')?.textContent?.trim()).toBe("Before");

    const cell = fixture.nativeElement.querySelector('tbody td[role="cell"]') as HTMLTableCellElement;
    cell.dispatchEvent(new MouseEvent("click", { bubbles: true }));
    expect(rowClicked).toHaveBeenCalledOnce();

    releaseCallback();
    await rowClicked.mock.results[0]!.value;
    await fixture.whenStable();

    expect(fixture.nativeElement.querySelector('tbody td[role="cell"]')?.textContent?.trim()).toBe(
      "After async callback",
    );
  });

  it("renders in-place row mutations when refreshKey changes and invalidates cached cell content", async () => {
    const fixture = await createFixture(SharepointChoiceTableCtor);
    const table = fixture.componentInstance;
    const row = { Id: 2, Name: "Before refresh" };
    table.prefix = "zoneless-refresh-key";
    table.allData = { Records: [row] };
    table.allCols = [{ field: "Name" }];
    table.selectedTab = "Records";
    fixture.detectChanges();
    expect(fixture.nativeElement.querySelector('tbody td[role="cell"]')?.textContent?.trim()).toBe("Before refresh");

    row.Name = "After refresh";
    table.refreshKey = "row-updated";
    await fixture.whenStable();

    expect(fixture.nativeElement.querySelector('tbody td[role="cell"]')?.textContent?.trim()).toBe("After refresh");
  });

  it("reports rejected row callbacks and leaves the view stable", async () => {
    const fixture = await createFixture(SharepointChoiceTableCtor);
    const table = fixture.componentInstance;
    const row = { Id: 3, Name: "Stable row" };
    const error = new Error("Row callback failed");
    table.prefix = "zoneless-rejected-row-callback";
    table.rowClicked = vi.fn().mockRejectedValue(error);
    table.allData = { Records: [row] };
    table.allCols = [{ field: "Name" }];
    table.selectedTab = "Records";
    fixture.detectChanges();
    const handleError = vi.spyOn(table["errorHandler"] as ErrorHandler, "handleError").mockImplementation(() => {});

    (fixture.nativeElement.querySelector('tbody td[role="cell"]') as HTMLTableCellElement).dispatchEvent(
      new MouseEvent("click", { bubbles: true }),
    );
    await fixture.whenStable();

    expect(handleError).toHaveBeenCalledWith(error);
    expect(fixture.nativeElement.querySelector('tbody td[role="cell"]')?.textContent?.trim()).toBe("Stable row");
  });

  it("renders changed columns and loading state from OnPush input setters", async () => {
    const fixture = await createFixture(SharepointChoiceTableCtor);
    const table = fixture.componentInstance;
    table.prefix = "zoneless-table-inputs";
    table.allData = { Records: [{ Id: 4, Name: "Visible" }] };
    table.allCols = [{ field: "Name" }];
    table.selectedTab = "Records";
    fixture.detectChanges();

    table.allCols = [{ field: "Name", headerName: "Updated name" }];
    table.loading = true;
    await fixture.whenStable();

    expect(fixture.nativeElement.textContent).toContain("Updated name");
    expect(fixture.nativeElement.querySelector('tbody tr[role="row"]')?.classList.contains("loading")).toBe(true);
  });
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
  it("provides columns and current-page rows for the grid to display", () => {
    const table = createTable();
    table.allData = {
      Records: [
        { Id: 1, Name: "Alpha", Score: 2 },
        { Id: 2, Name: "Beta", Score: 1 },
      ],
    };
    table.allCols = [{ field: "Name" }, { field: "Score" }];
    table.selectedTab = "Records";

    expect(table.fields("Records").map((column) => column.field)).toEqual(["Name", "Score"]);
    expect(table.currentPageRows(table.rows("Records")).map((row) => [row["Name"], row["Score"]])).toEqual([
      ["Alpha", 2],
      ["Beta", 1],
    ]);
  });

  it("filters displayed rows through the column filter handler", () => {
    const table = createTable();
    table.allData = {
      Records: [
        { Id: 1, Name: "Alpha", Score: 1 },
        { Id: 2, Name: "Beta", Score: 2 },
        { Id: 3, Name: "beta-block", Score: 4 },
      ],
    };
    table.selectedTab = "Records";

    table.filterChange({ field: "Name", filter: "text" }, "contains", "BETA");

    expect(table.rows("Records").map((row) => row["Name"])).toEqual(["Beta", "beta-block"]);

    table.filterChange({ field: "Score", filter: "number" }, "greater", "1");
    table.filterChange({ field: "Score", filter: "number" }, "less", "3");

    expect(table.currentPageRows(table.rows("Records")).map((row) => row["Name"])).toEqual(["Beta"]);
  });

  it("sorts displayed rows in both directions from header clicks without mutating source data", () => {
    const table = createTable();
    table.allData = {
      Records: [
        { Id: 1, Score: 3 },
        { Id: 2, Score: 1 },
        { Id: 3, Score: 2 },
      ],
    };
    table.selectedTab = "Records";
    const scoreColumn = { field: "Score" };
    const originalScores = table.allData["Records"]!.map((row) => row["Score"]);

    table.headerClick(scoreColumn);
    expect(table.rows("Records").map((row) => row["Score"])).toEqual([1, 2, 3]);

    table.headerClick(scoreColumn);
    expect(table.rows("Records").map((row) => row["Score"])).toEqual([3, 2, 1]);

    table.headerClick(scoreColumn);
    expect(table.rows("Records").map((row) => row["Score"])).toEqual(originalScores);
    expect(table.allData["Records"]!.map((row) => row["Score"])).toEqual(originalScores);
  });

  it("shows the requested page and returns to the first page when page size changes", () => {
    const table = createTable();
    table.allData = {
      Records: [{ Value: "one" }, { Value: "two" }, { Value: "three" }, { Value: "four" }, { Value: "five" }],
    };
    table.selectedTab = "Records";
    table.pageSize = 2;

    expect(table.currentPageRows(table.rows("Records")).map((row) => row["Value"])).toEqual(["one", "two"]);

    table.pageNumber = 2;
    expect(table.currentPageRows(table.rows("Records")).map((row) => row["Value"])).toEqual(["three", "four"]);

    table.pageNumber = 3;
    expect(table.currentPageRows(table.rows("Records")).map((row) => row["Value"])).toEqual(["five"]);

    table.pageSize = 3;
    expect(table.pageNumber).toBe(1);
    expect(table.currentPageRows(table.rows("Records")).map((row) => row["Value"])).toEqual(["one", "two", "three"]);
  });

  it("keeps column order and visibility scoped to each tab", () => {
    const table = createTable();
    table.allData = { Records: [], Archive: [] };
    table.allCols = [{ field: "Name" }, { field: "Status" }];
    table.columnOrder = { Records: ["Status", "Name"] };

    expect(table.fields("Records").map((column) => column.field)).toEqual(["Status", "Name"]);
    expect(table.fields("Archive").map((column) => column.field)).toEqual(["Name", "Status"]);

    table.columnToggle({ field: "Status" }, "Records");

    expect(table.fields("Records").map((column) => column.field)).toEqual(["Name"]);
    expect(table.fields("Archive").map((column) => column.field)).toEqual(["Name", "Status"]);
  });

  it("clears sorting and filters while retaining search and column layout", () => {
    const table = createTable();
    table.search = { Name: "match" };
    table.allData = {
      Records: [
        { Name: "match row", Status: "open" },
        { Name: "other row", Status: "closed" },
      ],
    };
    table.allCols = [{ field: "Name" }, { field: "Status" }];
    table.selectedTab = "Records";
    table.sort = { Records: [{ field: "Name", direction: "desc" }] };
    table.filter = { Records: { Status: { equals: "open" } } };
    table.columnOrder = { Records: ["Status", "Name"] };
    table.hiddenColumns = { Records: ["Status"] };
    const cleared = vi.fn();
    table.cleared.subscribe(cleared);

    table.clearTable();

    expect(table.sort).toEqual({});
    expect(table.filter).toEqual({});
    expect(table.rows("Records").map((row) => row["Name"])).toEqual(["match row"]);
    expect(table.search).toEqual({ Name: "match" });
    expect(table.columnOrder).toEqual({ Records: ["Status", "Name"] });
    expect(table.hiddenColumns).toEqual({ Records: ["Status"] });
    expect(table.fields("Records").map((column) => column.field)).toEqual(["Name"]);
    expect(cleared).toHaveBeenCalledOnce();
  });

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

  it("renders a field when its form data is loaded after the initial render", async () => {
    const fixture = await createComponentFixture();
    try {
      const component = fixture.componentInstance;
      component.field = "Title";
      component.spec = { Title: { TypeAsString: "Text", Title: "Title" } };
      component.form = { Title: "" };
      fixture.detectChanges();

      const input = fixture.nativeElement.querySelector(".Text input") as HTMLInputElement | null;
      expect(input?.value).toBe("");

      component.form = { Title: "Loaded from SharePoint" };
      await fixture.whenStable();
      fixture.detectChanges();

      expect(input?.value).toBe("Loaded from SharePoint");
    } finally {
      fixture.destroy();
    }
  });

  it("debounces text search as the user types, renders matches, and applies a selection", async () => {
    const fixture = await createComponentFixture();
    try {
      const component = fixture.componentInstance;
      const form: SharepointChoiceForm = { Contact: "" };
      const match: SharepointChoiceForm = { Contact: "Ada Lovelace" };
      const parent = { source: "directory" };
      const search = vi.fn(async () => [match]);
      const select = vi.fn();

      component.field = "Contact";
      component.spec = { Contact: { TypeAsString: "Text", Title: "Contact" } };
      component.form = form;
      component.text = { search, select, parent };
      fixture.detectChanges();
      vi.useFakeTimers();

      const input = fixture.nativeElement.querySelector(".Text input") as HTMLInputElement;
      input.value = "Ada";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { key: "a", bubbles: true }));
      fixture.detectChanges();

      expect(search).not.toHaveBeenCalled();
      await vi.advanceTimersByTimeAsync(250);
      fixture.detectChanges();

      expect(search).toHaveBeenCalledOnce();
      expect(search).toHaveBeenCalledWith("Ada", parent);
      const suggestion = fixture.nativeElement.querySelector(".AutoComplete li") as HTMLLIElement | null;
      expect(suggestion?.textContent?.trim()).toBe("Ada Lovelace");

      await component.selectedText(match);

      expect(form["Contact"]).toBe("Ada Lovelace");
      expect(select).toHaveBeenCalledWith(match, parent);
    } finally {
      fixture.destroy();
    }
  });

  it("starts user searches only after three characters and renders returned suggestions", async () => {
    const fixture = await createComponentFixture();
    try {
      const component = fixture.componentInstance;
      const user = { Id: 7, Title: "Ada Lovelace", LoginName: "ada@example.test" };
      component.field = "Owner";
      component.spec = { Owner: { TypeAsString: "User", Title: "Owner" } };
      component.form = { OwnerId: null };
      fixture.detectChanges();
      const search = vi.spyOn(component, "onUpUserSearch").mockImplementation(async () => {
        component.users = [user];
        component["chRef"].markForCheck();
      });
      const input = fixture.nativeElement.querySelector(".User > input:not([style])") as HTMLInputElement;
      vi.useFakeTimers();

      input.value = "ad";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { key: "d", bubbles: true }));
      fixture.detectChanges();
      await vi.advanceTimersByTimeAsync(300);
      expect(search).not.toHaveBeenCalled();

      input.value = "ada";
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { key: "a", bubbles: true }));
      fixture.detectChanges();
      await vi.advanceTimersByTimeAsync(249);
      expect(search).not.toHaveBeenCalled();
      await vi.advanceTimersByTimeAsync(1);
      fixture.detectChanges();

      expect(search).toHaveBeenCalledOnce();
      expect(search).toHaveBeenCalledWith("ada");
      expect(component.users).toEqual([user]);
      const suggestion = fixture.nativeElement.querySelector(".AutoComplete li") as HTMLLIElement | null;
      expect(suggestion?.textContent?.trim()).toBe("Ada Lovelace");
    } finally {
      fixture.destroy();
    }
  });

  it("stores ensured single and multi-user selections without adding duplicates", async () => {
    const component = createComponent();
    const contextWindow = window as typeof window & {
      _spPageContextInfo?: { webAbsoluteUrl: string; userLoginName?: string };
    };
    const previousContext = contextWindow._spPageContextInfo;
    contextWindow._spPageContextInfo = {
      webAbsoluteUrl: "https://example.test/sites/test",
      userLoginName: "test-user",
    };
    component.field = "Owner";
    component.form = { OwnerId: null, ReviewersId: [] };
    component.spec = {
      Owner: { TypeAsString: "User", Scope: "/sites/people" },
      Reviewers: { TypeAsString: "UserMulti", Scope: "/sites/people" },
    };
    const ensureUser = vi.fn(async (loginName: string) => ({
      Id: loginName === "ada@example.test" ? 7 : 8,
    }));
    mockWeb.mockReturnValue({ ensureUser });

    try {
      await component.selectedUser({
        Id: 0,
        Title: "Ada Lovelace",
        LoginName: "ada@example.test",
      });

      expect(component.form["OwnerId"]).toBe(7);

      component.field = "Reviewers";
      await component.selectedUser({
        Id: 0,
        Title: "Grace Hopper",
        LoginName: "grace@example.test",
      });
      await component.selectedUser({
        Id: 0,
        Title: "Grace Hopper",
        LoginName: "grace@example.test",
      });

      expect(component.form["ReviewersId"].results).toEqual([8]);
      expect(ensureUser).toHaveBeenCalledTimes(2);
      expect(mockWeb).toHaveBeenCalledWith([expect.anything(), "/sites/people"]);
    } finally {
      if (previousContext) contextWindow._spPageContextInfo = previousContext;
      else delete contextWindow._spPageContextInfo;
    }
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
