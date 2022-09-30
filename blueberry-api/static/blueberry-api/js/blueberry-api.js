        $( document ).ready(function() {

            var auth_cookie = getCookie("auth");
            if (auth_cookie) {
                document.getElementById("login-logout").innerHTML = "Log out";
                document.getElementById("login-logout").href = "/";
                document.getElementById("index_login_info").style.visibility = "hidden";;

            }
            else {
                document.getElementById("login-logout").innerHTML = "Log in";
                document.getElementById("login-logout").href = "/login";
            }
            $('.checkbox').checkbox();
        });


        function getCookie(cname) {
            var name = cname + "=";
            var ca = document.cookie.split(';');
            for(var i=0; i<ca.length; i++) {
                var c = ca[i];
                while (c.charAt(0)==' ') c = c.substring(1);
                if (c.indexOf(name) == 0) return c.substring(name.length,c.length);
            }
            return "";
        }

        function logout () {
            var auth_cookie = getCookie("auth");
            if (auth_cookie) {
                $.get('/logout');
            }
            if (auth_cookie) {
                document.cookie = "auth=; expires=Thu, 01 Jan 1970 00:00:00 UTC";
            }
        }

        function addFavorite () {
            $.ajax({
                  type: "POST",
                  url: '/add-in/add_favorite',
                  data: {bapi_id: $('#inputID').val()},
                  async: false,
                  success: function (jqXHR) {$('#inputID').val(jqXHR);
                                             if (jqXHR == "ID has been added to your favorites.") {
                                                $('#inputID').css('color', 'green');
                                                setTimeout(function(){ location.reload(); }, 1000);

                                             }
                                             else {
                                                $('#inputID').css('color', 'red');
                                             }
                  }
            })
        }

        function removeFavorite (bapi_id) {
            $.ajax({
                  type: "POST",
                  url: '/add-in/remove_favorite',
                  data: {bapi_id: bapi_id},
                  async: false,
                  success: function (jqXHR) {$('#inputID').val(jqXHR);
                                             if (jqXHR == "ID has been removed.") {
                                                $('#inputID').css('color', 'green');

                                             }
                                             else {
                                                $('#inputID').css('color', 'red');
                                             }
                  }
            })
            setTimeout(function(){ location.reload(); }, 1000);

        }

        function addPipeline () {
            $.ajax({
                  type: "POST",
                  //headers: {
                  //      'Accept': 'application/json',
                  //      'Content-Type': 'application/json'
                  //},
                  url: '/pipelines/add_pipeline',
                  //data: $.toJSON({"name": $('#pipeline_name').val()}),
                  data: {name: $('#inputPipelineName').val()},
                  success: function (jqXHR) {
                             $('#inputPipelineName').val(jqXHR);
                             if (jqXHR == "Pipeline has been added.") {
                                $('#inputPipelineName').css('color', 'green');
                                setTimeout(function(){ location.reload(); }, 1000);

                             }
                             else {
                                $('#inputPipelineName').css('color', 'red');
                             }
                  }
              })
        }


        function deletePipeline (name) {
            $.ajax({
                  type: "POST",
                  url: '/pipelines/delete_pipeline',
                  data: {name: name},
                  async: false,
                  success: function (jqXHR) {$('#inputPipelineName').val(jqXHR);
                                             if (jqXHR == "Pipeline has been deleted.") {
                                                $('#inputPipelineName').css('color', 'red');

                                             }
                                             else {
                                                $('#inputPipelineName').css('color', 'red');
                                             }
                  }
            })
            setTimeout(function(){ location.reload(); }, 1000);

        }

        function editPipeline (name) {
            $("#editPipelineTextDiv").text(name + " - Tasks");
            $("#editPipelineDiv").show();
        }

        function showConfDiv () {
            var bapi_id = $('#inputBlueberryID').val();
            var data_type = bapi_id.split(".")[2];
            var url = "/" + data_type + ".fetch";
            $.ajax({
                  type: "POST",
                  headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                  },
                  url: url,
                  data: $.toJSON({"bapi_id": bapi_id,
                                  "workbook_path": "",
                                  "worksheet": "",
                                  "workbook": "",
                                  "destination_cell": "",
                                  "user": ""}),
                  //data: {name: $('#inputPipelineName').val()},
                  success: function (jqXHR) {
                             var values = JSON.parse(jqXHR.data[0]);
                             showBlueberryData(values);
                             $('#inputPipelineName').val(jqXHR);
                             if (jqXHR == "Pipeline has been added.") {
                                $('#inputPipelineName').css('color', 'green');
                                setTimeout(function(){ location.reload(); }, 1000);

                             }
                             else {
                                $('#inputPipelineName').css('color', 'red');
                             }
                  }
              })
            var selection = document.getElementById("taskType");
            var selectionValue = selection.options[selection.selectedIndex].value;
            var selectionValueID = "#" + selectionValue + "Conf";
            $(selectionValueID).show();

        }


        function showBlueberryData (values) {
            $('.bapi_data').TidyTable({
			        columnTitles: generateColumnHeaders(values),
			        columnValues: transpose(values)
			        /*
			        columnValues: [
                                    ['1', '1A', '1B', '1C', '1D', '1E'],
                                    ['2', '2A', '2B', '2C', '2D', '2E'],
                                    ['3', '3A', '3B', '3C', '3D', '3E'],
                                    ['4', '4A', '4B', '4C', '4D', '4E'],
                                    ['5', '5A', '5B', '5C', '5D', '5E']
			                        ]
			        */
		    });
        }

        function transpose(a)
        {
          return a[0].map(function (_, c) { return a.map(function (r) { return r[c]; }); });
        }

        Element.prototype.remove = function() {
            this.parentElement.removeChild(this);
        }

        NodeList.prototype.remove = HTMLCollection.prototype.remove = function() {
            for(var i = this.length - 1; i >= 0; i--) {
                if(this[i] && this[i].parentElement) {
                    this[i].parentElement.removeChild(this[i]);
                }
            }
        }

        function generateColumnHeaders(values) {
            var columnHeaders = [];
            for (i = 1; i <= values.length; i++) {
                columnHeaders.push("Column " + i.toString());
            }
            return columnHeaders;
        }


