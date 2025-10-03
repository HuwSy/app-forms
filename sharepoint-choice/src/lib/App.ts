const url = window.location.href.toLowerCase();
const release = (url.includes('/app-') || url.includes('prd/')) ? 'LIVE'
  : (url.includes('/pre-') || url.includes('pre/')) ? 'PRE'
    : (url.includes('/tst-') || url.includes('tst/')) ? 'TST'
      : (url.includes('/sit-') || url.includes('sit/')) ? 'SIT'
        : 'DEV';

export const App = {
  Release: release,

  Tenancy: '',
  GraphClient: "",

  Grafana: null,

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
