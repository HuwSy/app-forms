// wrapper service for common code to angular structures

"use strict";
(function () {
    var appMod = angular.module('appforms', []);
    
    appMod.factory("listSvc", [function () {
        return {
            get: Webs.GetLibs,
            search: Webs.GetLibsSearch,
            searchChildren: Webs.GetFilesSearch,
            removeItem: Lists.RemoveListItem,
            postItem: Lists.SetListItem,
            getItems: Lists.GetListItem,
            getAttachments: Lists.GetListAttachments,
            delAttachments: Lists.DelListAttachments,
            addAttachments: Lists.AddListAttachments,
            upload: Lists.UploadAttachments,
			getByCaml: Lists.GetCaml,
            getList: Lists.GetListObject,
            contentType: Lists.GetSiteContentType,
            field: Lists.GetField,
            history: Lists.GetItemHistory,
            versions: Lists.GetItemHistory,
            exists: Lists.UrlExists,
            content: Lists.UrlContent
        }
    }]);

    appMod.factory("fileSvc", [function () {
        return {
            get: Webs.GetLibs,
            search: Webs.GetLibsSearch,
            searchChildren: Webs.GetFilesSearch,
            searchFiles: Webs.GetFilesUnderPath,
            removeItem: Lists.RemoveListItem,
            currentLocation: Lists.GetCurrentFolder,
            getList: Lists.GetListObject,
            contentType: Lists.GetSiteContentType,
			getByCaml: Lists.GetCaml,
            filesInFolder: Lists.GetInFolder,
            filesUnderPath: Lists.GetRecursiveFiles,
            foldersUnderPath: Lists.GetSubFolders,
            createFolder: Lists.CreateFolder,
            createFolders: Lists.CreateFolders,
            moveCopyFolders: Lists.IntoFolders,
            moveCopyFiles: Lists.MoveCopyDoc,
            addFile: Lists.AddFile,
            getFile: Lists.GetFile,
            field: Lists.GetField
        }
    }]);

    appMod.factory("modalSvc", [function () {
        return {
            url: Modal.URL,
            html: Modal.Modal,
            prompt: Modal.Prompt,
            confirm: Modal.YESNO,
            error: Modal.Error,
            title: Modal.Title,
            fields: Modal.Fields,
            redirect: Modal.Redirect,
            close: Modal.Close,
            open: Modal.SPModal
        }
    }]);
	
    appMod.factory("webSvc", [function () {
        return {
            email: Webs.Email,
            editing: Webs.IsEditing,
            get: Webs.GetWebs,
            search: Webs.GetWebsSearch,
            current: Webs.CurrentWeb,
			create: Webs.CreateWeb,
            setup: Webs.SetupWeb,
            user: Webs.GetUser,
            ensureUser: Webs.EnsureUser,
            getGroups: Webs.GetGroups,
            ensureGroup: Webs.EnsureGroup,
            addUsers: Webs.AddUsers
        }
    }]);
	
    appMod.factory("querySvc", [function () {
        return {
            get: Override.ParameterByName
        }
    }]);
	
    appMod.factory("termSvc", [function () {
        return {
            getTermset: Terms.GetTermSet,
            flattenToTree: Terms.GetTermSetAsTree,
            findInTree: Terms.Find,
            storeTree: Terms.UpdateSessionStorage,
            currentTerm: Terms.GetCurrentTerm,
            findInSet: Terms.GetTermObjectByNameId,
            createTerm: Terms.CreateTermWithinContext,
            getFieldSet: Terms.GetFieldTermIds,
            toObject: Terms.GetFieldValue,
            toString: Terms.ToString
        };
    }]);
})();