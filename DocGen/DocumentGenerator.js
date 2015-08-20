/// <reference path="References/jquery.d.ts"/>
/// <reference path="References/SharePoint.d.ts"/>
/// <reference path="References/lodash.d.ts"/>
/// <reference path="References/DocumentGeneratorConfig.d.ts"/>
var DocumentGeneratorConfig;
(function (DocumentGeneratorConfig) {
    DocumentGeneratorConfig.config = {
        listName: "QuickLinks",
        excludedLists: [
            'AppPackagesList',
            'OData__x005f_catalogs_x002f_appdata',
            'DraftAppsList',
            'OData__x005f_catalogs_x002f_design',
            'ContentTypeSyncLogList',
            'IWConvertedForms',
            'Shared_x0020_Documents',
            'FormServerTemplates',
            'GettingStartedList',
            'OData__x005f_catalogs_x002f_lt',
            'OData__x005f_catalogs_x002f_MaintenanceLogs',
            'OData__x005f_catalogs_x002f_masterpage',
            'PublishedFeedList',
            'ProjectPolicyItemList',
            'SiteAssets',
            'SitePages',
            'OData__x005f_catalogs_x002f_solutions',
            'Style_x0020_Library',
            'TaxonomyHiddenListList',
            'OData__x005f_catalogs_x002f_theme',
            'UserInfo',
            'OData__x005f_catalogs_x002f_wp',
            'OData__x005f_catalogs_x002f_wfpub',
            'Documents',
            'PublishingImages',
            'Pages',
            'WorkflowTasks'
        ],
        excludedFieldTypes: [
            'Lookup'
        ],
        excludedContentTypesForFields: [
            'Folder',
            'Item'
        ]
    };
})(DocumentGeneratorConfig || (DocumentGeneratorConfig = {}));
var DocumentionService = (function () {
    function DocumentionService(pageContext) {
        this.pageContext = pageContext;
    }
    DocumentionService.prototype.getItems = function (url, filterCallback) {
        var restUrl = this.pageContext.webServerRelativeUrl + url;
        console.log(restUrl);
        var d = jQuery.Deferred();
        var request = jQuery.ajax({
            url: restUrl,
            method: "GET",
            headers: { "Accept": "application/json; odata=verbose" },
            success: (function (data) {
                var parsedData = data.d.results;
                if (parsedData.length > 0) {
                    var result = filterCallback(parsedData);
                    d.resolve(result);
                }
                else {
                    d.resolve([]);
                }
            }),
            error: (function (data, errorCode, errorMessage) {
                if (window.console)
                    console.error("Request failed:" + data + ", " + errorCode + ", " + errorMessage);
                d.reject();
            })
        });
        return d.promise();
    };
    DocumentionService.prototype.getLists = function (itemsToExclude) {
        return this.getItems("/_api/web/lists?$expand=DefaultView'", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.EntityTypeName) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    DocumentionService.prototype.getSiteContentTypes = function (listId, itemsToExclude) {
        return this.getItems("/_api/web/contenttypes'", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.Name) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    DocumentionService.prototype.getContentTypes = function (listId, itemsToExclude) {
        return this.getItems("/_api/web/lists(guid'" + listId + "')/contenttypes'", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.Name) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    DocumentionService.prototype.getFields = function (listId, itemsToExclude) {
        return this.getItems("/_api/web/lists(guid'" + listId + "')/fields", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.InternalName) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    DocumentionService.prototype.getContentTypeFields = function (listId, contentTypeId, itemsToExclude) {
        return this.getItems("/_api/web/lists(guid'" + listId + "')/contenttypes('" + contentTypeId + "')/fields", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.InternalName) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    DocumentionService.prototype.getViewFields = function (listId, viewId, itemsToExclude) {
        return this.getItems("/_api/web/lists(guid'" + listId + "')/views('" + viewId + "')/ViewFields", function (parsedData) {
            if (itemsToExclude && itemsToExclude.length > 0) {
                return _.filter(parsedData, function (u) {
                    return itemsToExclude.indexOf(u.InternalName) === -1;
                });
            }
            else {
                return parsedData;
            }
        });
    };
    return DocumentionService;
})();
var DocumentationRenderer = (function () {
    function DocumentationRenderer(pageContext) {
        this.pageContext = pageContext;
    }
    DocumentationRenderer.prototype.buildMarkup = function (groupName, listContainer) {
        //var arrowImgSrc = this.siteUrl + '/Style Library/ProcFit/img/link-arrow.png';
        var d = jQuery.Deferred();
        var service = new DocumentionService(this.pageContext);
        var that = this;
        var excludedLists = DocumentGeneratorConfig.config.excludedLists;
        /* [
         'AppPackagesList',
         'OData__x005f_catalogs_x002f_appdata',
         'DraftAppsList',
         'OData__x005f_catalogs_x002f_design',
         'ContentTypeSyncLogList',
         'IWConvertedForms',
         'Shared_x0020_Documents',
         'FormServerTemplates',
         'GettingStartedList',
         'OData__x005f_catalogs_x002f_lt',
         'OData__x005f_catalogs_x002f_MaintenanceLogs',
         'OData__x005f_catalogs_x002f_masterpage',
         'PublishedFeedList',
         'ProjectPolicyItemList',
         'SiteAssets',
         'SitePages',
         'OData__x005f_catalogs_x002f_solutions',
         'Style_x0020_Library',
         'TaxonomyHiddenListList',
         'OData__x005f_catalogs_x002f_theme',
         'UserInfo',
         'OData__x005f_catalogs_x002f_wp',
         'OData__x005f_catalogs_x002f_wfpub',
         'Documents',
         'PublishingImages',
         'Pages',
         'WorkflowTasks'
     ];*/
        service.getLists(excludedLists).done(function (lists) {
            //var htmlBuilder = [];
            //that.htmlBuilder.push("<h2>Lists: </h2>");
            lists.forEach(function (item) {
                listContainer.append("<div class=\"link-clickable\"><a href=\"" + item.Url + "\">" + item.EntityTypeName + "</a><span>, ID: " + item.Id + "</span><span>, DefaultViewID: " + item.DefaultView.Id + "</span></div>"); //Title
            });
            //listContainer.append("<div id=\"FieldContainer\"></div>");
            lists.forEach(function (list) {
                var listName = list.Title.split(' ').join('');
                //Content types
                var tableContentTypeBuilder = [];
                tableContentTypeBuilder.push("<div id=\"" + listName + "ContentTypeContainer\"><h3>" + list.Title + " Content Types: </h3>");
                tableContentTypeBuilder.push("<table id=\"" + listName + "ContentTypeTable\" class=\"altrowstable\" id=\"alternatecolor\">");
                tableContentTypeBuilder.push("<tr>");
                tableContentTypeBuilder.push("<th>Name</th><th>ID</th><th>Group</th>");
                tableContentTypeBuilder.push("</tr>");
                tableContentTypeBuilder.push("</table></div>");
                listContainer.append(tableContentTypeBuilder.join(""));
                service.getContentTypes(list.Id).done(function (contentTypes) {
                    var contentTypeContainer = jQuery('#' + listName + 'ContentTypeTable');
                    //var listNameref = listName;
                    contentTypes.forEach(function (contentType) {
                        console.log(that.htmlBuilder);
                        contentTypeContainer.append("<tr><td>" + contentType.Name + "</td><td>" + contentType.Id.StringValue + "</td><td>" + contentType.Group + "</td></tr>");
                        //
                        if (DocumentGeneratorConfig.config.excludedContentTypesForFields.indexOf(contentType.Name) === -1) {
                            var contentTypeName = contentType.Name.split(' ').join('');
                            var tableContentTypeFieldsBuilder = [];
                            tableContentTypeFieldsBuilder.push("<div id=\"" + listName + contentTypeName + "ContentTypeFieldContainer\"><h3>" + list.Title + " " + contentType.Name + " Fields: </h3>");
                            tableContentTypeFieldsBuilder.push("<table id=\"" + listName + contentTypeName + "ContentTypeFieldTable\" class=\"altrowstable\" id=\"alternatecolor\">");
                            tableContentTypeFieldsBuilder.push("<tr>");
                            tableContentTypeFieldsBuilder.push("<th>Name</th><th>ID</th><th>Group</th><th>Content Name</th><th>Content type ID</th>");
                            tableContentTypeFieldsBuilder.push("</tr>");
                            tableContentTypeFieldsBuilder.push("</table></div>");
                            $('#' + listName + 'ContentTypeContainer').append(tableContentTypeFieldsBuilder.join(""));
                            service.getContentTypeFields(list.Id, contentType.Id.StringValue).done(function (contentTypeFields) {
                                var contentTypeFieldsContainer = jQuery('#' + listName + contentTypeName + 'ContentTypeFieldTable');
                                contentTypeFields.forEach(function (contentTypeField) {
                                    console.log(that.htmlBuilder);
                                    contentTypeFieldsContainer.append("<tr><td>" + contentTypeField.Title + "</td><td>" + contentTypeField.Id + "</td><td>" + contentTypeField.Group + "</td><td>" + contentType.Name + "</td><td>" + contentType.Id.StringValue + "</td></tr>");
                                });
                            });
                        }
                        //
                    });
                });
                //Default View Fields
                var viewFieldsBuilder = [];
                viewFieldsBuilder.push("<div id=\"" + listName + "ViewFieldContainer\"><h3>" + list.Title + " Default View Fields: </h3>");
                viewFieldsBuilder.push("<table id=\"" + listName + "ViewFieldTable\" class=\"altrowstable\" id=\"alternatecolor\">");
                viewFieldsBuilder.push("<tr>");
                viewFieldsBuilder.push("<th>Title</th><th>InternalName</th> <th>TypeDisplayName</th>");
                viewFieldsBuilder.push("</tr>");
                viewFieldsBuilder.push("</table></div>");
                listContainer.append(viewFieldsBuilder.join(""));
                service.getViewFields(list.Id, list.DefaultView.Id).done(function (fields) {
                    var viewFieldContainer = jQuery('#' + listName + 'ViewFieldTable');
                    fields.forEach(function (field) {
                        console.log(that.htmlBuilder);
                        viewFieldContainer.append("<tr><td>" + field.Title + "</td><td>" + field.InternalName + "</td><td>" + field.TypeDisplayName + "</td></tr>");
                    });
                });
                //All List Columns
                var tableFliedBuilder = [];
                tableFliedBuilder.push("<div id=\"" + listName + "FieldContainer\"><h3>" + list.Title + " Fields: </h3>");
                tableFliedBuilder.push("<table id=\"" + listName + "FieldTable\" class=\"altrowstable\" id=\"alternatecolor\">");
                tableFliedBuilder.push("<tr>");
                tableFliedBuilder.push("<th>Title</th><th>InternalName</th> <th>TypeDisplayName</th>");
                tableFliedBuilder.push("</tr>");
                tableFliedBuilder.push("</table></div>");
                listContainer.append(tableFliedBuilder.join(""));
                service.getFields(list.Id).done(function (fields) {
                    var fieldContainer = jQuery('#' + listName + 'FieldTable');
                    fields.forEach(function (field) {
                        console.log(that.htmlBuilder);
                        fieldContainer.append("<tr><td>" + field.Title + "</td><td>" + field.InternalName + "</td><td>" + field.TypeDisplayName + "</td></tr>");
                    });
                });
            });
            d.resolve();
            //d.resolve(that.htmlBuilder.join(""));
        });
        return d.promise();
    };
    DocumentationRenderer.prototype.renderDocumentation = function (parentElement) {
        this.buildMarkup(this.groupName, parentElement).done(function () {
            console.log('rendering finished');
            //parentElement.append(htmlLinks);
        });
    };
    return DocumentationRenderer;
})();
//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js"); 
//# sourceMappingURL=DocumentGenerator.js.map