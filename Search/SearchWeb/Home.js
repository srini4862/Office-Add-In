(function () {
    "use strict";

    var messageBanner;

    Office.initialize = function (reason) {

        var tokenObtained;

        $(document).ready(function () {

            initializeFabric();

            //Search box
            $(".ms-SearchBox").SearchBox();

            var authContext = getAuthContext();

            //For updating serach box with selected text
            Office.context.document.addHandlerAsync("documentSelectionChanged", selectionChange, function (result) { });

            //For updating link to the office document
            $('#insertText').click(updateDocument);

            //sign in and out
            $("#signInLink").click(function () {
                authContext.login();
            });
            $("#signOutLink").click(function () {
                authContext.logOut();
            });

            //save tokens if this is a return from AAD
            authContext.handleWindowCallback();

            //Set's access token to global variable
            getAccessToken(authContext);

            //Search button
            $('#search').click(searchText);

            //Hide content
            $("#content").hide();

        });

        function initializeFabric() {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
        }

        function selectionChange() {
            try {
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
                    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                        write('Action failed. Error: ' + asyncResult.error.message);
                    }
                    else {
                        $(".ms-SearchBox-field").val(asyncResult.value);
                    }
                });

                $('.ms-SearchBox-label').hide();
                // Show cancel button by adding is-active class
                $('.ms-SearchBox').addClass('is-active');
            }
            catch (err) {
                errorHandler(err);
            }
        }

        function updateDocument() {
            try {
                var url = $("div.is-selected > span.ms-ListItem-secondaryText").text();

                Office.context.document.setSelectedDataAsync(url, function (asyncResult) {
                    var error = asyncResult.error;
                    if (asyncResult.status === "failed") {
                        $('#display-data').text("Failure" + error.message);
                    }
                });
            }
            catch (err) {
                errorHandler(err);
            }

        }

        function getAuthContext() {
            try {
                //OAuth
                //authorization context
                var tenantUrl = 'jilledi1.onmicrosoft.com';
                var siteCollectionUrl = 'https://jilledi1.sharepoint.com/sites/poc';
                var clientIdValue = '721d5fa8-8345-43e5-85d0-ccc2c743ae03';

                var authContext = new AuthenticationContext({
                    instance: 'https://login.microsoftonline.com/',
                    tenant: tenantUrl,
                    clientId: clientIdValue,
                    postLogoutRedirectUri: window.location.origin,
                    cacheLocation: 'localStorage'
                });

                return authContext;
            }
            catch (err) {
                errorHandler(err);
            }
        }

        function getAccessToken(authContext) {
            try {
                var resource = 'https://jilledi1.sharepoint.com';
                var user = authContext.getCachedUser();
                if (user) {  //successfully logged in

                    //welcome user
                    $("#loginMessage").text("Welcome, " + user.userName);
                    $("#signInLink").hide();
                    $("#signOutLink").show();

                    //call rest endpoint
                    authContext.acquireToken(resource, function (error, token) {

                        if (error || !token) {
                            $("#loginMessage").text('ADAL Error Occurred: ' + error);
                            return;
                        }
                        else {
                            tokenObtained = token;
                        }
                    });

                }
                else if (authContext.getLoginError()) { //error logging in
                    $("#signInLink").show();
                    $("#signOutLink").hide();
                    $("#loginMessage").text(authContext.getLoginError());
                }
                else { //not logged in
                    $("#signInLink").show();
                    $("#signOutLink").hide();
                    $("#loginMessage").text("You are not logged in.");
                }
            }
            catch (err) {
                errorHandler(err);
            }

        }

        function searchText() {
            try {
                var searchText = $('.ms-SearchBox-field').val();

                if (searchText == '' || tokenObtained == undefined) {
                    return false;
                }

                var endpoint = "https://jilledi1.sharepoint.com/sites/poc/_api/search/query?querytext='" + searchText + "'&contenttype:'STS_ListItem_DocumentLibrary'&top=5";

                $.ajax({
                    type: 'GET',
                    url: endpoint,
                    headers: {
                        'Accept': 'application/json',
                        'Authorization': 'Bearer ' + tokenObtained,
                    },
                }).done(function (data) {
                    getResults(data);
                    $(".ms-ListItem").ListItem();
                }).fail(function (err) {
                    jQuery("#loginMessage").text('Error calling REST endpoint: ' +
                        err.statusText);
                }).always(function () {
                });
            }
            catch (err) {
                errorHandler(err);
            }
        }

        function getResults(data) {
            try {
                var results = data.PrimaryQueryResult.RelevantResults.Table.Rows;
                //data.PrimaryQueryResult.RelevantResults.Table.Rows[0].Cells[4].Value
                if (results.length > 0) {
                    $("#content").show();
                }
                else {
                    $("#content").hide();
                }
                $('#resultList').empty();

                var title, link, author, modified;
                var listItems = '';

                $.each(results, function () {

                    $.each(this.Cells, function () {

                        if (this.Key == "Title") {
                            title = this.Value;
                        }
                        else if (this.Key == 'Path') {
                            link = this.Value;
                        }
                        else if (this.Key == 'Author') {
                            author = this.Value;
                        }
                        else if (this.Key == 'LastModifiedTime') {
                            modified = (new Date(this.Value)).toString('mm/dd/yyyy');
                        }
                    });

                    listItems += getResult(title, link, author, modified);

                });
                $('#resultList').append(listItems);
            }
            catch (err) {
                errorHandler(err);
            }
        }

        function getResult(title, link, author, modified) {

            var listItem = '<div class="ms-ListItem is-selectable">' +
            '<span class="ms-ListItem-primaryText">' + title + '</span>' +
            '<span class="ms-ListItem-secondaryText">' + link + '</span>' +
            '<span class="ms-ListItem-tertiaryText">Author: ' + author + '   Modified On : ' + modified + '</span>' +
            '<div class="ms-ListItem-selectionTarget js-toggleSelection"></div>' +
            '<div class="ms-ListItem-actions">' +
            '</div>' +
            '</div>';

            return listItem;

        }

    };

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }

})();