'use strict';

//ExecuteOrDelayUntilScriptLoaded(initializePage, "sp.js");

var Paging = window.Paging || {};
Paging.Utilities = Paging.Utilities || {};

//#region Common Code
Paging.Utilities.readAjaxCall = function (restUrl) {
    var dfd = $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json;odata=verbose",
        headers: {
            "accept": "application/json;odata=verbose"
        }
    });
    return dfd.promise();
}
Paging.Utilities.displayError = function (error) {
    $('#message').append("<br/> There was an error getting the data..");
    $('#message').append(error);
}
//#endregion


//#region List Items
Paging.Utilities.displayListItems = function (data) {
    var msg = "<br/><br/><b>List Items are:</b><ul>";
    if (data.d.results.length) {
        if (data.d.results.length > 1) {
            msg += data.d.results.length + " items:<ul>";
            $.each(data.d.results, function (index, value) {
                msg += "<li>" + value.Title + "</li>";
            });

        }
        else {
            msg += "1 item:<ul>";
            msg += "<li>" + data.d.results[0].Title + "</li>";
        }
        msg += "</ul>";
    }
    else {
        msg += "no items";
    }

    if (data.d.__next) {
        msg += "<br />";
        msg += "Results are paged - next set of results available here: <br />" + data.d.__next;
    }
    $('#message').append(msg);
}
Paging.Utilities.displayFilteredListItems = function (data, parentName) {
    var msg = "<br/><br/><b>The following items have a SampleData value of '" + parentName + "'</b>:<ul>";
    $.each(data.d.results, function (index, value) {
        msg += "<li>" + value.Title + "</li>";
    });
    msg += "</ul>";
    $('#message').append(msg);
}
Paging.Utilities.displayFieldInfo = function (data) {
    var msg = "<b>There are " + data.d.results.length + " lists on this site.<br />";
    msg += "Field information:</b><br />";

    //iterate collection of lists
    $.each(data.d.results, function (index, value) {
        msg += "<li><u>" + value.Title + " has these fields:</u></li><ul>";

        //iterate fields on each list
        $.each(data.d.results[index].Fields.results, function (index, value) {
            msg += "<li>" + value.Title + "</li>";
        });

        msg += "</ul>";
    });
    msg += "</ul>";
    $('#message').append(msg);
}
//#endregion

$(document).ready(function () {
    var pager = new Paging.Pager();
    $("#readItemsPaged").click(pager.readListItemsPaged);
    $("#readIncorrectlyPaged").click(pager.readListItemsIncorrectlyPaged);
    $("#readFilteredItems").click(pager.readFilteredItems);
    $("#readListFields").click(pager.readFields);
});
Paging.Pager = function () {
    var pageNum = 0;
    var pageSize = 3;

    //#region Paging List Items
    function _readListItemsPaged() {
        var msg = "Getting next page lists items...";
        var listName = "SampleData";

        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/GetByTitle('" + listName + "')/items?$top=" + pageSize);
        //need to add skiptoken unencoded or else the % in %26 and %3d gets encoded to %25 which breaks things
        restUrl = encodeURI(restUrl) + "&$skiptoken=Paged%3dTRUE%26p_ID%3d" + pageNum * pageSize;
        pageNum += 1;
        var readPromise = Paging.Utilities.readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                Paging.Utilities.displayListItems(data);
            },
            function (jqXHR, status, error) {
                Paging.Utilities.displayError(error);
            }
        );
    }

    function _readListItemsIncorrctPaged() {
        var msg = "Getting next page lists items...";
        var listName = "SampleData";

        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/GetByTitle('" + listName + "')/items?$skip=" + pageNum * pageSize + "&$top=" + pageSize);

        pageNum += 1;
        var readPromise = Paging.Utilities.readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                Paging.Utilities.displayListItems(data);
            },
            function (jqXHR, status, error) {
                Paging.Utilities.displayError(error);
            }
        );
    }

    function _readFilteredItems() {
        var msg = "Getting child items...";
        var parentName = prompt("Get Children for which Parent?", "Item 1");
        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/GetByTitle('ChildList')/items?$select=Title,SampleData/Title&$expand=SampleData/Title&$filter=(SampleData/Title eq '" + parentName + "')");

        var readPromise = Paging.Utilities.readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                Paging.Utilities.displayFilteredListItems(data, parentName);
            },
            function (jqXHR, status, error) {
                Paging.Utilities.displayError(error);
            }
        );
    }

    function _readFields() {
        var msg = "Getting List Fields...";
        $('#message').text(msg);
        var restUrl = SP.Utilities.UrlBuilder.urlCombine(_spPageContextInfo.webServerRelativeUrl,
            "_api/web/lists/?$expand=fields/title");

        var readPromise = Paging.Utilities.readAjaxCall(restUrl);
        readPromise.then(
            function (data, status, jqXHR) {
                Paging.Utilities.displayFieldInfo(data);
            },
            function (jqXHR, status, error) {
                Paging.Utilities.displayError(error);
            }
        );
    }
    //#endregion
    return {
        readListItemsPaged: _readListItemsPaged,
        readListItemsIncorrectlyPaged: _readListItemsIncorrctPaged,
        readFilteredItems: _readFilteredItems,
        readFields: _readFields
    }


}
