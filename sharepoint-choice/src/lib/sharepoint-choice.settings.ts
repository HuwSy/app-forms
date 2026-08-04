// Load configuration from JSON file
async function loadAppConfig() {
  const configUrl = "/sites/AppCatalog/SiteAssets/sharepoint-choice.json";

  try {
    const response = await fetch(configUrl);
    if (!response.ok) {
      throw new Error(`Failed to load config: ${response.statusText}`);
    }
    const config = await response.json();

    const url = window.location.href.toLowerCase();
    const release = determineRelease(url, config.environmentMappings);

    let grafana = config.grafana || null;
    if (grafana) {
      for (const key in grafana) {
        if (typeof grafana[key] === "string") {
          grafana[key] = grafana[key]
            .replace("{origin}", window.location.origin)
            .replace("{release}", release)
            .replace("{userAgent}", navigator?.userAgent || "")
            .replace("{language}", navigator?.language || "");
        }
      }
    }

    return {
      Release: release,
      Tenancy: config.tenancy || "",
      GraphClient: config.graphClient || "",
      Grafana: grafana,
      ApiServers: config.apiServers || {},
      ApiToken: config.apiToken || {},
      ApiMap: config.apiMap || {},
    };
  } catch (error) {
    console.error("Error loading app configuration:", error);
    return getDefaultConfig();
  }
}

function determineRelease(url: string, environmentMappings: any[]): string {
  for (const mapping of environmentMappings) {
    for (const pattern of mapping.patterns) {
      if (url.includes(pattern)) {
        return mapping.environment;
      }
    }
  }
  return "DEV";
}

function getDefaultConfig() {
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

export const App = await loadAppConfig();
