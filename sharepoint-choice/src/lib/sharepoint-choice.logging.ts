import { ErrorHandler } from '@angular/core';
import { fromError } from 'stacktrace-js';
import { App } from './App'

export class SharepointChoiceLogging implements ErrorHandler {
  private _grafana: {
    "stream": { [key: string]: string },
    "values": Array<Array<string>>
  };

  constructor() {
    let w: any = window;

    // base object for this page, only 1 logger per page should be present to avoid time conflicts
    this._grafana = w[`GrafanaLogging`] || {
      "stream": App.Grafana || {
        "Environment": App.Release,
        "System": 'SharePoint-Choice',
        "Hostname": window.location.origin,
        "Username": undefined,
        "UserAgent": navigator?.userAgent,
        "Language": navigator?.language,
      },
      "values": []
    };
  }

  public async handleError(error: any): Promise<void> {
    let w: any = window;

    // get the stack trace and reduce for grafana posting
    let stackTrace = '';
    try {
      stackTrace = JSON.stringify(await fromError(error, { offline: true }));
      if (!stackTrace || stackTrace == 'null')
        stackTrace = '';
      else if (stackTrace.length >= 2048)
        stackTrace = stackTrace.substring(0, 2045) + '...';
    } catch (e) {
      // fail stack trace silently and log the remainder of the error as is
    }

    // attempt fixup of user name, site title etc on logging, as spo doesnt have user in context immediately
    if (w._spPageContextInfo) {
      for (let t in this._grafana.stream) {
        if (t.includes('Username') && w._spPageContextInfo.userLoginName)
          this._grafana.stream[t] = w._spPageContextInfo.userLoginName;
        if (t.includes('System') && w._spPageContextInfo.webTitle)
          this._grafana.stream[t] = w._spPageContextInfo.webTitle;
        if (t.includes('Hostname') && w._spPageContextInfo.webAbsoluteUrl)
          this._grafana.stream[t] = w._spPageContextInfo.webAbsoluteUrl;
      }
    }

    // define current page path params
    var split = window.location.pathname.split('/');
    var path = window.location.pathname.match(/\/[^\/]*\/[^\/]*\/([^?]*)/);
    let Params = {
      prefix: split[1] ?? '',
      site: split[2] ?? '',
      path: path && path[1] ? path[1] : '',
      search: window.location.search,
      hash: window.location.hash
    };

    // console here to keep click into source working
    switch (error.level?.substring(0, 1).toUpperCase()) {
      case 'E':
        console.error(error);
        break;
      case 'W':
        console.warn(error);
        break;
      case 'I':
        console.info(error);
        break;
      case 'D':
        console.log(error);
        break;
      default:
        console.trace(error);
        break;
    }

    let body = [((new Date()).getTime() * 1_000_000).toString(), JSON.stringify({ Params, Level: error.level ?? 'Unknown', Message: error.message ?? error, StackTrace: stackTrace })];

    // prevent flooding from multiple loggers on screen as timestamps must be sequential or grafana will not accept
    if (this._grafana.values.filter(x => x[1] == body[1] && parseInt(x[0]) + 1_000_000 < parseInt(body[0])).length == 0) {
      this._grafana.values.push(body);
      this.sendErrors();
    }
  }

  private sendErrors() {
    let w: any = window;

    // any logs sending on current page then delay and retry
    if (w._grafana) {
      setTimeout(this.sendErrors, 25);
      return;
    }

    // get how many records right now
    var saving = this._grafana.values.length;

    // no logs then end
    if (saving == 0)
      return;

    // build post message sending
    let body = JSON.stringify({
      "streams": [
        this._grafana
      ]
    });

    // console log what will send
    for (var e = 0; e < saving; e++)
      console.error(this._grafana.values[e][1]);

    // truncate whats being logged so new additions are going in here during posting
    this._grafana.values = this._grafana.values.splice(saving);

    // dont log under localhost or if no grafana api map
    if (window.location.host.includes('localhost') || !App.ApiServers?.['grafana']?.[App.Release])
      return;

    // start posting
    w._grafana = true;
    fetch(App.ApiServers['grafana'][App.Release], {
      method: 'POST',
      body: body,
      headers: {
        'Content-Type': 'application/json'
      }
    }).then(b => {
      w._grafana = false;
      console.log(`Submitted ${saving} to grafana ${b}.`);
    }).catch(e => {
      w._grafana = false;
      console.error(`Error logging ${saving} to grafana ${e}.`);
    })
  }
}
