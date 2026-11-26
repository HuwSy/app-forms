// Load configuration from JSON file
interface EnvironmentMapping {
  patterns: string[];
  environment: string;
}

interface ApiConfiguration {
  server?: string;
  name?: string;
  port?: number;
  [environment: string]: string | number | undefined;
}

interface AppConfiguration {
  Release: string;
  Tenancy: string;
  GraphClient: string;
  Grafana: Record<string, string | null> | null;
  ApiServers: Record<string, Record<string, string>>;
  ApiToken: Record<string, Record<string, string>>;
  ApiMap: Record<string, ApiConfiguration>;
}

interface AppConfigurationData {
  environmentMappings: EnvironmentMapping[];
  grafana?: Record<string, string | null> | null;
  tenancy?: string;
  graphClient?: string;
  apiServers?: AppConfiguration["ApiServers"];
  apiToken?: AppConfiguration["ApiToken"];
  apiMap?: AppConfiguration["ApiMap"];
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isStringRecord(value: unknown): value is Record<string, string> {
  return isRecord(value) && Object.values(value).every((item) => typeof item === "string");
}

function isStringOrNullRecord(value: unknown): value is Record<string, string | null> {
  return isRecord(value) && Object.values(value).every((item) => item === null || typeof item === "string");
}

function isStringRecordMap(value: unknown): value is Record<string, Record<string, string>> {
  return isRecord(value) && Object.values(value).every(isStringRecord);
}

function isApiConfiguration(value: unknown): value is ApiConfiguration {
  return (
    isRecord(value) &&
    Object.entries(value).every(([, item]) => typeof item === "string" || typeof item === "number") &&
    (value["server"] === undefined || typeof value["server"] === "string") &&
    (value["name"] === undefined || typeof value["name"] === "string") &&
    (value["port"] === undefined || typeof value["port"] === "number")
  );
}

function isAppConfigurationData(value: unknown): value is AppConfigurationData {
  if (!isRecord(value) || !Array.isArray(value["environmentMappings"])) return false;
  if (
    !value["environmentMappings"].every(
      (mapping) =>
        isRecord(mapping) &&
        Array.isArray(mapping["patterns"]) &&
        mapping["patterns"].every((pattern) => typeof pattern === "string") &&
        typeof mapping["environment"] === "string",
    )
  )
    return false;

  return (
    (value["grafana"] === undefined || value["grafana"] === null || isStringOrNullRecord(value["grafana"])) &&
    (value["tenancy"] === undefined || typeof value["tenancy"] === "string") &&
    (value["graphClient"] === undefined || typeof value["graphClient"] === "string") &&
    (value["apiServers"] === undefined || isStringRecordMap(value["apiServers"])) &&
    (value["apiToken"] === undefined || isStringRecordMap(value["apiToken"])) &&
    (value["apiMap"] === undefined ||
      (isRecord(value["apiMap"]) && Object.values(value["apiMap"]).every(isApiConfiguration)))
  );
}

let appConfigPromise: Promise<AppConfiguration> | undefined;

export function loadAppConfig(): Promise<AppConfiguration> {
  return (appConfigPromise ??= (async () => {
    const configUrl = "/sites/AppCatalog/SiteAssets/sharepoint-choice.json";

    try {
      const response = await fetch(configUrl);
      if (!response.ok) {
        throw new Error(`Failed to load config: ${response.statusText}`);
      }
      const responseData: unknown = await response.json();
      if (!isAppConfigurationData(responseData)) throw new TypeError("The SharePoint Choice configuration is invalid.");
      const appConfig = responseData;

      const currentUrl = window.location.href.toLowerCase();
      let releaseEnvironment = "DEV";
      for (const mapping of appConfig.environmentMappings) {
        let matched = false;
        for (const pattern of mapping.patterns) {
          if (currentUrl.includes(pattern)) {
            releaseEnvironment = mapping.environment;
            matched = true;
            break;
          }
        }
        if (matched) break;
      }

      const grafanaConfig: Record<string, string | null> | null = appConfig.grafana || null;
      if (grafanaConfig) {
        for (const labelName in grafanaConfig) {
          const labelValue = grafanaConfig[labelName];
          if (typeof labelValue === "string") {
            grafanaConfig[labelName] = labelValue
              .replace("{origin}", window.location.origin)
              .replace("{release}", releaseEnvironment)
              .replace("{userAgent}", navigator?.userAgent || "")
              .replace("{language}", navigator?.language || "");
          }
        }
      }

      return {
        Release: releaseEnvironment,
        Tenancy: appConfig.tenancy || "",
        GraphClient: appConfig.graphClient || "",
        Grafana: grafanaConfig,
        ApiServers: appConfig.apiServers || {},
        ApiToken: appConfig.apiToken || {},
        ApiMap: appConfig.apiMap || {},
      };
    } catch (error) {
      console.error("Error loading app configuration:", error);
      return {
        Release: "DEV",
        Tenancy: "",
        GraphClient: "",
        Grafana: null,
        ApiServers: {},
        ApiToken: {},
        ApiMap: {},
      };
    }
  })());
}

export const App = await loadAppConfig();
