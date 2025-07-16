var release = ~document.location.href.toLowerCase().indexOf('/app-') || ~document.location.href.toLowerCase().indexOf('prd/') ? 'LIVE'
    : ~document.location.href.toLowerCase().indexOf('/pre-') || ~document.location.href.toLowerCase().indexOf('pre/') ? 'PRE'
      : ~document.location.href.toLowerCase().indexOf('/tst-') || ~document.location.href.toLowerCase().indexOf('tst/') ? 'TST'
        : ~document.location.href.toLowerCase().indexOf('/sit-') || ~document.location.href.toLowerCase().indexOf('sit/') ? 'SIT'
          : 'DEV';

export const App = {
  Release: release,

  Tenancy: '',
  GraphClient: "",

  Grafana: {
    "Hostname": document.location.origin,
    "Category": 'SPO',
    "Environment": release,
    "Username": undefined,
    "UserAgent": navigator.userAgent,
    "Language": navigator.language,
  },

  ApiServers: {
    'grafana': {
      LIVE: 'https://grafanaprd/loki/api/v1/push',
      PRE: 'https://grafanapre/loki/api/v1/push',
      TST: 'https://grafanatst/loki/api/v1/push',
      SIT: 'https://grafanasit/loki/api/v1/push',
      DEV: 'https://grafanadev/loki/api/v1/push',
    },
    'common': {
      LIVE: 'https://apiprd',
      PRE: 'https://apipre',
      TST: 'https://apitst',
      SIT: 'https://apisit',
      DEV: 'https://apidev',
    }
  },
  ApiToken: {
    'common': {
      LIVE: 'api://API.',
      PRE: 'api://API-Pre.',
      TST: 'api://API-Tst.',
      SIT: 'api://API-Sit.',
      DEV: 'api://API-Dev.',
    }
  },
  ApiMap: {
    'test': {
      LIVE: '',
      PRE: '',
      TST: '',
      SIT: '',
      DEV: '',
      name: 'Test',
      port: 44301,
      server: 'common'
    }
  }
}
