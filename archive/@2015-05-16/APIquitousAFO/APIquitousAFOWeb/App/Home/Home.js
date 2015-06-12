/// <reference path="../App.js" />



(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#publish').click(publishData);
            $('#test-button').click(testSubmit);

            var source = $("#entry-template").html();
            var template = Handlebars.compile(source);
            var context = { title: "My New Post", body: "This is my first post!" };
            var html = template(context);

            $('#bartosz').append(html);

        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    //app.showNotification('The selected text is:', '"' + result.value + '"');
                    /*
                    var data = {'aq_name': document.getElementsByName("aq_name").value,
                        'aq_description': document.getElementsByName("aq_description").value,
                        'aq_organization':document.getElementsByName("aq_organization").value,
                        'aq_created_by':document.getElementsByName("aq_created_by").value
                    };
                    */
                    var data = {
                        "aq_name": $("input#aq_name").val(),
                        "aq_description": $("textarea#aq_description").val(),
                        "aq_organization": $("input#aq_organization").val(),
                        "aq_created_by": $("input#aq_created_by").val()
                               };
                    data = JSON.stringify(data);

                    var xhr = new XMLHttpRequest();
                    xhr.open('POST', "http://apiquitous.ngrok.com/publish", true);
                    xhr.setRequestHeader('Content-Type', 'application/json');
                    xhr.send(data);

                    /*
                    var responsePublish;
                    $.support.cors = true;
                    $.ajax({
                        contentType: 'application/json; charset=utf-8',
                        url: "http://apiquitous.ngrok.com/publish",
                        dataType: "jsonp",
                        //accepts: { json: 'application/json' },
                        //jsonp: 'showData',
                        data: data,
                        //jsonpCallBack: 'jan', 
                        //data: JSON.stringify({ "Janusz": "Korwin-Mikke" }),
                        //async: false,
                        cache: false,
                        success: function (data) {
                            //responsePublish = response;
                            app.showNotification(data);
                        },
                        fail: function (xhr, status, error) {
                            app.showNotification(status);
                        },
                        complete: function (response, jqXHR) {
                            app.showNotification('Tadek.', response);
                        },
                    });
                    //app.showNotification('Your data has been published.', responsePublish);
                    */                
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        )};

        // Publishes selected range to persistent storage
    function publishData() {
        /*
        var data = {
            'aq_name': document.getElementsByName("aq_name").value,
            'aq_description': document.getElementsByName("aq_description").value,
            'aq_organization': document.getElementsByName("aq_organization").value,
            'aq_created_by': document.getElementsByName("aq_created_by").value
        };
        */

        //var data = document.getElementsByName("aq_name");
        //data = JSON.stringify(data);
        
        /*
        $.support.cors = true;

        $("#aq_publish_form").ajaxSubmit({
            contentType: "application/json; charset=utf-8",
            url: "http://apiquitous.ngrok.com/publish",
            type: "POST",
            dataType: "json",
            data: JSON.stringify(document.getElementsByName("aq_name")),
            async: false,
            cache: false
        })
        */
        /*
        $.ajax({
            contentType: "application/json; charset=utf-8",
            url: "http://apiquitous.ngrok.com/publish",
            type: "POST",
            dataType: "json",
            data: data,
            async: false,
            cache: false
        });
        */
    };

    
    function testSubmit() {

    $("#aq_publish_form").submit(function () {

        $.ajax({
            contentType: "application/json; charset=utf-8",
            url: "http://apiquitous.ngrok.com/publish",
            type: "POST",
            dataType: "json",
            data: JSON.stringify({"Janusz":"Korwin-Mikke"}),
            //data: JSON.stringify($("#aq_publish_form").serialize()), // serializes the form's elements.
            success: function (data) {
                alert("Udalo sie");
            }
        });

        return false; // avoid to execute the actual submit of the form.
    });
    
    };
    
})();