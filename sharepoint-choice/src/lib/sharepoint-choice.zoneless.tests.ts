import { ChangeDetectionStrategy, ErrorHandler, provideZonelessChangeDetection } from "@angular/core";
import { TestBed } from "@angular/core/testing";
import { afterEach, describe, expect, it, vi } from "vitest";
import { SharepointChoiceForm, SharepointChoiceTabs } from "./sharepoint-choice.models";
import { SharepointChoiceComponent } from "./sharepoint-choice.component";
import { SharepointChoiceTable } from "./sharepoint-choice.table";

const fixtures: Array<{ destroy(): void }> = [];

async function createFixture(component: typeof SharepointChoiceComponent | typeof SharepointChoiceTable) {
  await TestBed.configureTestingModule({
    imports: [component],
    providers: [provideZonelessChangeDetection()],
  }).compileComponents();
  const fixture = TestBed.createComponent(component);
  fixtures.push(fixture);
  return fixture;
}

afterEach(() => {
  fixtures.splice(0).forEach((fixture) => fixture.destroy());
  vi.restoreAllMocks();
});

describe("zoneless OnPush rendering", () => {
  it("renders async choice suggestions after markForCheck without an extra detectChanges call", async () => {
    const fixture = await createFixture(SharepointChoiceComponent);
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
    const fixture = await createFixture(SharepointChoiceComponent);
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
    const fixture = await createFixture(SharepointChoiceComponent);
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
    const fixture = await createFixture(SharepointChoiceComponent);
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
    const fixture = await createFixture(SharepointChoiceTable);
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
    expect(table.changeDetection).not.toBe(ChangeDetectionStrategy.Default);
  });

  it("renders in-place row mutations when refreshKey changes and invalidates cached cell content", async () => {
    const fixture = await createFixture(SharepointChoiceTable);
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
    const fixture = await createFixture(SharepointChoiceTable);
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
    const fixture = await createFixture(SharepointChoiceTable);
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
    expect(fixture.nativeElement.querySelector("tbody")?.classList.contains("loading")).toBe(true);
  });
});
