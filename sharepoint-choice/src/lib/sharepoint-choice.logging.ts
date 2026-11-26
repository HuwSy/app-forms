import { ErrorHandler } from "@angular/core";
import { App } from "./sharepoint-choice.settings";

type GrafanaLogEntry = [timestamp: string, serializedLogLine: string];

interface SharePointPageContext {
  userLoginName?: string;
  webTitle?: string;
  webAbsoluteUrl?: string;
}

interface GrafanaLoggerState {
  stream: Record<string, string>;
  values: GrafanaLogEntry[];
  pendingMessages: Set<string>;
  sending: boolean;
  retryCount: number;
  flushTimer: number | undefined;
  lastTimestampNs: bigint;
  configurationWarningShown: boolean;
}

type LoggerWindow = Window & {
  GrafanaLogging?: unknown;
  _spPageContextInfo?: SharePointPageContext;
};

const FLUSH_DELAY_MS = 50;
const REQUEST_TIMEOUT_MS = 15_000;
const RETRY_BASE_DELAY_MS = 1_000;
const MAX_RETRY_DELAY_MS = 30_000;
const MAX_RETRY_COUNT = 7;
const STACK_TRACE_LIMIT = 2_048;

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null;
}

function isGrafanaLogEntry(value: unknown): value is GrafanaLogEntry {
  return Array.isArray(value) && typeof value[0] === "string" && typeof value[1] === "string";
}

export class SharepointChoiceLogging implements ErrorHandler {
  private readonly _grafana: GrafanaLoggerState;

  public constructor() {
    this._grafana = this.getSharedLoggerState();
  }

  private getSharedLoggerState(): GrafanaLoggerState {
    const browserWindow = window as LoggerWindow;
    const existing = browserWindow.GrafanaLogging;
    const existingObject = isRecord(existing) ? existing : undefined;
    if (
      existingObject &&
      isRecord(existingObject["stream"]) &&
      Object.values(existingObject["stream"]).every((label) => typeof label === "string") &&
      Array.isArray(existingObject["values"]) &&
      existingObject["values"].every(isGrafanaLogEntry) &&
      existingObject["pendingMessages"] instanceof Set &&
      typeof existingObject["sending"] === "boolean" &&
      typeof existingObject["retryCount"] === "number" &&
      (existingObject["flushTimer"] === undefined || typeof existingObject["flushTimer"] === "number") &&
      typeof existingObject["lastTimestampNs"] === "bigint" &&
      typeof existingObject["configurationWarningShown"] === "boolean"
    ) {
      return existing as GrafanaLoggerState;
    }

    const values = Array.isArray(existingObject?.["values"]) ? existingObject["values"].filter(isGrafanaLogEntry) : [];
    const stream: Record<string, string> = {};
    const configuredStream = existingObject?.["stream"] ?? App.Grafana;
    if (isRecord(configuredStream)) {
      for (const [label, rawValue] of Object.entries(configuredStream)) {
        stream[label] = rawValue == null ? "" : String(rawValue);
      }
    }

    if (Object.keys(stream).length === 0) {
      Object.assign(stream, {
        Environment: App.Release,
        System: "SharePoint-Choice",
        Hostname: window.location.origin,
        Username: "",
        UserAgent: navigator.userAgent,
        Language: navigator.language,
      });
    }

    const lastTimestampNs = values.reduce((latestTimestampNs, [timestamp]) => {
      if (!/^\d+$/.test(timestamp)) return latestTimestampNs;

      const parsedTimestampNs = BigInt(timestamp);
      return parsedTimestampNs > latestTimestampNs ? parsedTimestampNs : latestTimestampNs;
    }, 0n);
    const state: GrafanaLoggerState = {
      stream,
      values,
      pendingMessages: new Set(values.map(([, serializedLogLine]) => serializedLogLine)),
      sending: false,
      retryCount: 0,
      flushTimer: undefined,
      lastTimestampNs,
      configurationWarningShown: false,
    };

    browserWindow.GrafanaLogging = state;
    return state;
  }

  public handleError(error: unknown): void {
    void this.processError(error).catch((loggingError: unknown) => {
      console.error("SharepointChoiceLogging failed to process an error.", loggingError);
    });
  }

  private async processError(error: unknown): Promise<void> {
    const source = isRecord(error) ? (error["error"] ?? error) : error;
    const sourceObject = isRecord(source) ? source : undefined;
    const outerObject = isRecord(error) ? error : undefined;
    const levelValue = sourceObject?.["level"] ?? outerObject?.["level"];
    const level = typeof levelValue === "string" ? levelValue : "Unknown";
    const messageValue = outerObject?.["message"] ?? sourceObject?.["message"] ?? source;
    let message: string;
    if (messageValue instanceof Error) message = messageValue.message;
    else if (typeof messageValue === "string") message = messageValue;
    else {
      try {
        message = JSON.stringify(messageValue) ?? String(messageValue);
      } catch (serializationError) {
        console.warn("Unable to serialize a SharePoint Choice error message.", serializationError);
        message = String(messageValue);
      }
    }

    const stackValue = sourceObject?.["stack"] ?? outerObject?.["stack"];
    const rawStackTrace = typeof stackValue === "string" ? stackValue : "";
    const stackTrace =
      rawStackTrace.length >= STACK_TRACE_LIMIT
        ? `${rawStackTrace.substring(0, STACK_TRACE_LIMIT - 3)}...`
        : rawStackTrace;

    this.updateSharePointContext();

    const pathSegments = window.location.pathname.split("/");
    const pathMatch = window.location.pathname.match(/\/[^\/]*\/[^\/]*\/([^?]*)/);
    // Accepted privacy and payload-size risk: page query/hash values, user identifiers, errors, and stack traces
    // are sent without redaction or a total payload-size limit; callers should not put secrets in URLs or errors.
    const pageParameters = {
      prefix: pathSegments[1] ?? "",
      site: pathSegments[2] ?? "",
      path: pathMatch?.[1] ?? "",
      search: window.location.search,
      hash: window.location.hash,
    };

    switch (level.substring(0, 1).toUpperCase()) {
      case "E":
        console.error(error);
        break;
      case "W":
        console.warn(error);
        break;
      case "I":
        console.info(error);
        break;
      case "D":
        console.log(error);
        break;
      default:
        console.trace(error);
        break;
    }

    const serializedLogLine = JSON.stringify({
      Params: pageParameters,
      Level: level,
      Message: message,
      StackTrace: stackTrace,
    });

    if (this._grafana.pendingMessages.has(serializedLogLine)) return;

    const wallClockTimestamp = BigInt(Date.now()) * 1_000_000n;
    const timestamp =
      wallClockTimestamp > this._grafana.lastTimestampNs ? wallClockTimestamp : this._grafana.lastTimestampNs + 1n;
    this._grafana.lastTimestampNs = timestamp;

    this._grafana.values.push([timestamp.toString(), serializedLogLine]);
    this._grafana.pendingMessages.add(serializedLogLine);
    this.scheduleFlush(FLUSH_DELAY_MS);
  }

  private updateSharePointContext(): void {
    const pageContext = (window as LoggerWindow)._spPageContextInfo;
    if (!pageContext) return;

    const replacements: Record<string, string | undefined> = {
      Username: pageContext.userLoginName,
      System: pageContext.webTitle,
      Hostname: pageContext.webAbsoluteUrl,
    };

    for (const label of Object.keys(this._grafana.stream)) {
      for (const [name, value] of Object.entries(replacements)) {
        if (label.includes(name) && value) this._grafana.stream[label] = value;
      }
    }
  }

  private scheduleFlush(delay: number): void {
    if (this._grafana.sending || this._grafana.flushTimer !== undefined || this._grafana.values.length === 0) {
      return;
    }

    this._grafana.flushTimer = window.setTimeout(() => {
      this._grafana.flushTimer = undefined;
      void this.sendErrors();
    }, delay);
  }

  private async sendErrors(): Promise<void> {
    if (this._grafana.sending || this._grafana.values.length === 0) return;

    // Remote logging is skipped on localhost. Without a release endpoint, pending records are warned about and
    // discarded; there is no durable offline store.
    if (window.location.hostname === "localhost") {
      this.discardPendingLogs();
      return;
    }

    const configuredEndpoints: unknown = App.ApiServers?.["grafana"];
    const endpoint = isRecord(configuredEndpoints) ? configuredEndpoints[App.Release] : undefined;
    if (typeof endpoint !== "string" || endpoint.length === 0) {
      if (!this._grafana.configurationWarningShown) {
        console.warn(`Grafana logging is not configured for release ${App.Release}; pending logs will be discarded.`);
        this._grafana.configurationWarningShown = true;
      }
      this.discardPendingLogs();
      return;
    }

    const batch = this._grafana.values.slice();
    const controller = new AbortController();
    const timeout = window.setTimeout(() => controller.abort(), REQUEST_TIMEOUT_MS);
    this._grafana.sending = true;

    let retryDelayMs: number | undefined;
    try {
      const body = JSON.stringify({
        streams: [{ stream: this._grafana.stream, values: batch }],
      });
      const response = await fetch(endpoint, {
        method: "POST",
        body,
        headers: {
          "Content-Type": "application/json",
        },
        signal: controller.signal,
      });

      if (!response.ok) {
        throw new Error(`Grafana returned HTTP ${response.status} ${response.statusText}.`);
      }

      this._grafana.values.splice(0, batch.length);
      for (const [, serializedLogLine] of batch) {
        this._grafana.pendingMessages.delete(serializedLogLine);
      }
      this._grafana.retryCount = 0;
      console.info(`Submitted ${batch.length} logs to Grafana.`);
    } catch (error) {
      this._grafana.retryCount = Math.min(this._grafana.retryCount + 1, MAX_RETRY_COUNT);
      retryDelayMs = Math.min(
        RETRY_BASE_DELAY_MS * 2 ** Math.min(this._grafana.retryCount - 1, MAX_RETRY_COUNT),
        MAX_RETRY_DELAY_MS,
      );
      console.error(`Error logging ${batch.length} records to Grafana; retrying in ${retryDelayMs}ms.`, error);
    } finally {
      window.clearTimeout(timeout);
      this._grafana.sending = false;
    }

    if (retryDelayMs !== undefined) this.scheduleFlush(retryDelayMs);
    else if (this._grafana.values.length > 0) this.scheduleFlush(FLUSH_DELAY_MS);
  }

  private discardPendingLogs(): void {
    this._grafana.values.length = 0;
    this._grafana.pendingMessages.clear();
  }
}
