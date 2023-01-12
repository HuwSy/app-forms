export const App = {
    AppName: 'SharePoint-Choice',
    Release: ~document.location.href.toLowerCase().indexOf('/prd-') ? 'LIVE' 
        : ~document.location.href.toLowerCase().indexOf('/pre-') ? 'PRE' 
        : ~document.location.href.toLowerCase().indexOf('/tst-') ? 'TST' 
        : ~document.location.href.toLowerCase().indexOf('/sit-') ? 'SIT' 
        : 'DEV',
    Token: 'https://api',
    TinyMCEKey: '',
    Tenancy: ''
}