export const App = {
    AppName: 'SharePoint-Choice',
    Category: 'SPO',
    Entity: 'All',
    Release: ~document.location.href.toLowerCase().indexOf('/app-') ? 'LIVE' 
        : ~document.location.href.toLowerCase().indexOf('/pre-') ? 'PRE' 
        : ~document.location.href.toLowerCase().indexOf('/tst-') ? 'TST' 
        : ~document.location.href.toLowerCase().indexOf('/sit-') ? 'SIT' 
        : 'DEV',
    Token: '',
    Tenancy: '',
    GraphClient: ''
}
