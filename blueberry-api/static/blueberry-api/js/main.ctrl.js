angular.module('blueberry-angular-app').controller("MainController", function(){
    var vm = this;
    vm.values = [
                    [1, 2, 3, 4, 5, 6],
                    [7, 8, 9, 10, 11, 12],
                    [13, 14, 15, 16, 17, 19],
                    [7, 8, 9, 10, 11, 12],
                    [13, 14, 15, 16, 17, 19]
                ];
    vm.fetchedData = [[1,2],[3,4]];
    vm.bapi_id = "";
    vm.data = [[1],[1]];
    vm.hotTable;
    vm.hotTableHeaders;

    vm.fetchData = function() {
            var data_type = vm.bapi_id.split(".")[2];
            var url = "/" + data_type + ".fetch";
            $.ajax({
                  type: "POST",
                  headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                  },
                  async: false,
                  url: url,
                  data: $.toJSON({"bapi_id": vm.bapi_id,
                                  "workbook_path": "",
                                  "worksheet": "",
                                  "workbook": "",
                                  "destination_cell": "",
                                  "user": ""}),
                  success: function (jqXHR) {
                             vm.data = transpose(JSON.parse(jqXHR.data[0]));
                             vm.hotTableHeaders = jqXHR.headers_list;
                             if (document.getElementById("example").hasAttribute("firstOne")) {
                                var bapiDisplay = document.getElementById("example");
                                bapiDisplay.remove();
                                var newDiv = document.createElement('div');
                                newDiv.id = "example";
                                var exampleContainer = document.getElementById("example_container");
                                exampleContainer.appendChild(newDiv);
                                //element.removeAttribute('firstOne');
                             }
                             var bapiDisplay = document.getElementById("example");
                             bapiDisplay.setAttribute('firstOne', true);
                             vm.excelGrid();
                             //var fetchedData = JSON.parse(jqXHR.data[0]);
                             //showBlueberryData(fetchedData);
                             //$('#inputPipelineName').val(jqXHR);
                  }
              })
    }

    vm.showConfDiv = function () {
            var selection = document.getElementById("taskType");
            var selectionValue = selection.options[selection.selectedIndex].value;
            var selectionValueID = "#" + selectionValue + "Conf";
            $(selectionValueID).show();
        }

    vm.excelGrid = function () {
        var container = document.getElementById('example');
        var hot = new Handsontable(container, {
          data: vm.data,
          minSpareRows: 0,
          rowHeaders: false,
          colHeaders: false,
          contextMenu: false
        });
        vm.hotTable = hot;
    }

    vm.showOutput = function()  {

    }

    vm.selectColumns = function () {
      var url = "/proxy/select_headers";
      $.ajax({
                  type: "POST",
                  headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                  },
                  async: false,
                  url: url,
                  data: $.toJSON({"data": JSON.stringify(vm.data),
                                  "headers_list_all": ["Jan", "Maria"],
                                  //"headers_list_all": JSON.stringify(vm.hotTableHeaders),
                                  "headers_list_selected": ["Jan"]}),
                  success: function (jqXHR) {
                             vm.data = transpose(JSON.parse(jqXHR.data[0]));
                             vm.hotTableHeaders = jqXHR.headers_list;
                             if (document.getElementById("example").hasAttribute("firstOne")) {
                                var bapiDisplay = document.getElementById("example");
                                bapiDisplay.remove();
                                var newDiv = document.createElement('div');
                                newDiv.id = "example";
                                var exampleContainer = document.getElementById("example_container");
                                exampleContainer.appendChild(newDiv);
                                //element.removeAttribute('firstOne');
                             }
                             var bapiDisplay = document.getElementById("example");
                             bapiDisplay.setAttribute('firstOne', true);
                             vm.excelGrid();
                             //var fetchedData = JSON.parse(jqXHR.data[0]);
                             //showBlueberryData(fetchedData);
                             //$('#inputPipelineName').val(jqXHR);
                  }
              })
    }

    vm.createCORSRequest = function() {
        var method = "GET";
        var url = "http://riskcontrol.pythonanywhere.com/api";
        var xhr = new XMLHttpRequest();
        xhr.withCredentials = true;
        if ("withCredentials" in xhr) {

          // Check if the XMLHttpRequest object has a "withCredentials" property.
          // "withCredentials" only exists on XMLHTTPRequest2 objects.
          xhr.open(method, url, true);

        } else if (typeof XDomainRequest != "undefined") {

          // Otherwise, check if XDomainRequest.
          // XDomainRequest only exists in IE, and is IE's way of making CORS requests.
          xhr = new XDomainRequest();
          xhr.open(method, url);

        } else {

          // Otherwise, CORS is not supported by the browser.
          xhr = null;

        }
        xhr.onload = function() {
                     var responseText = xhr.responseText;
                     console.log(responseText);
                     // process the response.
                    };

        xhr.onerror = function() {
          console.log('There was an error!');
        };
        xhr.send();  
        } 

    vm.testPost2 = function () {
      var url = "http://riskcontrol.pythonanywhere.com/api";
      $.ajax({
                  type: "POST",
                  headers: {
                        'Accept': 'application/json',
                        'Content-Type': 'application/json'
                  },
                  async: false,
                  url: url,
                  data: $.toJSON({"jsonrpc": "2.0",
                                  "method": "App.index",
                                  "params": {},
                                  "id": "1"}),
                  success: function (jqXHR) {
                             alert(jqXHR);
                             //var fetchedData = JSON.parse(jqXHR.data[0]);
                             //showBlueberryData(fetchedData);
                             //$('#inputPipelineName').val(jqXHR);
                  }
              })
    }



    //vm.getHeaders = function (hotTable) {
    //    return hotTable.getDataAtRow(0);
    //}



});