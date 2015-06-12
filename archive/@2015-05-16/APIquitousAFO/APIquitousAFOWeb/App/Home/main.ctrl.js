angular.module('fetch').controller("MainController", function () {
    
    var vm = this;
    vm.debug = 1;
    vm.debug1 = '';
    vm.debug2 = '';
    vm.debug3 = '';
    vm.debug4 = '';
    vm.debug5 = '';
    vm.debug6 = '';


    vm.user = 'bartosz.piechnik@ch.abb.com'

    vm.menusList = [
        ['publish',
            ['new', 'update']
        ],
        ['fetch',
            ['new', 'update']
        ]
    ];

    vm.fetchNewAQID = "";

    vm.fetchNew = function () {
        var data = {
            'aq_id': vm.fetchNewAQID,
            'user': vm.user
        };
        //data = JSON.stringify(data);
        vm.debug2 = 'Function works';
        $.support.cors = true;


                /*
        $.ajax({
    
            contentType: 'application/json; charset=utf-8',
            url: "http://apiquitous.ngrok.com/fetch_new",
            dataType: "jsonp",
            data: data,
            cache: false,
            success: function (data) {
                vm.debug1 = data;
            },
            fail: function (xhr, status, error) {
                vm.debug2 = xhr;
            },
            complete: function (response, jqXHR) {
                vm.debug3 = response;
            },
        });
        */
    };

    function createCORSRequest(method, url) {
        var xhr = new XDomainRequest();
        xhr.open(method, url);
        //xhr.setRequestHeader('Content-Type', 'application/json');
        //xhr.setRequestHeader('X-Custom-Header', 'value');
        //xhr.withCredentials = true;
        return xhr;
    }

    vm.makeCorsRequest = function () {
        // All HTML5 Rocks properties support CORS.
        var data = {
            'aq_id': vm.fetchNewAQID,
            'user': vm.user
        };
        data = JSON.stringify(data);
        var url = 'http://apiquitous.ngrok.com/fetch_new';

        var xhr = createCORSRequest('POST', url);
        
        if (!xhr) {
            alert('CORS not supported');
            return;
        }

        // Response handlers.
        xhr.onload = function () {
            var text = xhr.responseText;
            vm.debug1 = text;
        };

        xhr.onerror = function () {
            vm.debug1 = "It didn't work";
        };
        vm.debug2 = data;
        xhr.send(data);
    }

    // THIRD TRY - XMLHttpRequest

    function createCORSRequest2(method, url) {
        var xhr = new XMLHttpRequest();
        xhr.open(method, url, true);
        if (method == 'POST') {
            xhr.setRequestHeader('Content-Type', 'application/json');
        }
        vm.debug3 = xhr;
        return xhr;
        
    }


    // Make the actual CORS request.
    vm.makeCorsRequest2 = function () {
        // All HTML5 Rocks properties support CORS.
        var data = {
            'aq_id': vm.fetchNewAQID,
            'user': vm.user
        };
        data = JSON.stringify(data);
        var url = 'http://apiquitous.ngrok.com/fetch_new';

        var xhr = createCORSRequest2('POST', url);
        if (!xhr) {
            alert('CORS not supported');
            return;
        }

        // Response handlers.
        xhr.onload = function () {
            var text = xhr.responseText;
            vm.debug1 = text;
        };

        xhr.onesuccess = function () {
            vm.debug1 = "Well done";
        };

        vm.debug2 = data;
        xhr.send(data);
        //debugger;
    }

    // Make the actual CORS request.
    vm.getFetched = function () {
        // All HTML5 Rocks properties support CORS.
        var url = 'http://apiquitous.ngrok.com/get_fetched';

        var xhr = createCORSRequest2('POST', url);

        xhr.send(null);

        xhr.onsuccess = function () {
            vm.debug5 = "Success";
        };

        // Response handlers.
        xhr.onload = function () {
            var text = xhr.responseText;
            text = JSON.parse(text);
            //debugger;
            vm.debug4 = text;
            vm.names = text['names'];
            vm.users = text['users'];
            vm.workbook_paths = text['workbook_paths'];
            vm.workbooks = text['workbooks'];
            vm.worksheets = text['worksheets'];
            vm.destination_cells = text['destination_cells'];
        };


        xhr.onfail = function () {
            vm.debug5 = "Failure";
        };
    }

    vm.getPublished = function () {
        // All HTML5 Rocks properties support CORS.
        var url = 'http://apiquitous.ngrok.com/get_published';

        var xhr = createCORSRequest2('POST', url);

        xhr.send(null);

        xhr.onsuccess = function () {
            vm.debug5 = "Success";
        };

        // Response handlers.
        xhr.onload = function () {
            var published = xhr.responseText;
            published = JSON.parse(published);
            //debugger;
            vm.debug4 = published;
            vm.published_ids = published['ids'];
            vm.published_users = published['users'];
            vm.published_workbook_paths = published['workbook_paths'];
            vm.published_workbooks = published['workbooks'];
            vm.published_worksheets = published['worksheets'];
            vm.published_destination_cells = published['destination_cells'];
            vm.published_data_types = published['data_types'];
        };


        xhr.onfail = function () {
            vm.debug5 = "Failure";
        };
    }


    vm.fetch_many = function () {
        // All HTML5 Rocks properties support CORS.
        var url = 'http://apiquitous.ngrok.com/fetch_many';

        var xhr = createCORSRequest2('POST', url);

        var data = {
            'ids': vm.names,
            'users': vm.users,
            'workbook_paths': vm.workbook_paths,
            'workbooks': vm.workbooks,
            'worksheets': vm.worksheets,
            'destination_cells': vm.destination_cells
        };
        data = JSON.stringify(data);

        xhr.send(data);

        xhr.onsuccess = function () {
            vm.debug5 = "Fetch all sent";
        };

        xhr.onfail = function () {
            vm.debug5 = "Failure";
        };
    }

    vm.publish_many = function () {
        // All HTML5 Rocks properties support CORS.
        var url = 'http://apiquitous.ngrok.com/publish_many';

        var xhr = createCORSRequest2('POST', url);

        var data = {
            'ids': vm.published_ids,
            'users': vm.published_users,
            'workbook_paths': vm.published_workbook_paths,
            'workbooks': vm.published_workbooks,
            'worksheets': vm.published_worksheets,
            'destination_cells': vm.published_destination_cells,
            'data_types': vm.published_data_types
        };
        data = JSON.stringify(data);

        xhr.send(data);

        xhr.onsuccess = function () {
            vm.debug5 = "Publish all sent";
        };

        xhr.onfail = function () {
            vm.debug5 = "Failure";
        };
    }



    vm.tableToggle = {
        'one': false,
        'two': false,
        'three': false
    };
    vm.tableToggle0 = false;
    vm.tableToggle1 = false;
    vm.tableToggle2 = false;

    vm.tablePublishedToggle0 = false;
    vm.tablePublishedToggle1 = false;
    vm.tablePublishedToggle2 = false;

    vm.menu1 = false;
    vm.menu2 = false;
    vm.menu3 = false;

    vm.toggleVerticalTable = function (table) {
        // All HTML5 Rocks properties support CORS.
        //debugger;
        /*
        if (vm.tableToogle[table]) {
            vm.tableToogle[table] = true;
        } else {
            vm.tableToggle[table] = false;
        }
        */
        
        if (table == 'one') {
            vm.tableToggle0 = true;
        } else if (table == 'two') {
            vm.tableToggle1 = true;
        } else {
            vm.tableToggle2 = true;
        }


    };

    vm.togglePublishedVerticalTable = function (published_table) {
        // All HTML5 Rocks properties support CORS.
        //debugger;
        /*
        if (vm.tableToogle[table]) {
            vm.tableToogle[table] = true;
        } else {
            vm.tableToggle[table] = false;
        }
        */

        if (published_table == 'one') {
            vm.tablePublishedToggle0 = true;
        } else if (published_table == 'two') {
            vm.tablePublishedToggle1 = true;
        } else {
            vm.tablePublishedToggle2 = true;
        }


    };


    vm.loadVideoDetails = function (videoIndex) {
        // Dynamically create a new HTML SCRIPT element in the webpage.
        vm.debug3 = 'New Script';
        var script = document.createElement("script");
        // Specify the URL to retrieve the indicated video from a feed of a current list of videos,
        // as the value of the src attribute of the SCRIPT element. 
        script.setAttribute("src", "https://gdata.youtube.com/feeds/api/videos/" +
            videos[videoIndex].Id + "?alt=json-in-script&callback=videoDetailsLoaded");
        // Insert the SCRIPT element at the end of the HEAD section.
        document.getElementsByTagName('head')[0].appendChild(script);
    };

    
    function existenceList(level) {
        var level = level - 1;
        var temp = [];
        for (var i in vm.menusList) {
            temp.push([vm.menusList[i][level], true]);
        };
        return temp;
    };

    vm.level1 = existenceList(1);

    function falseAllItems() {
        vm.level1Display.publish = false;
        vm.level1Display.fetch = false;
        vm.publishLevel2Display.new = false;
        vm.publishLevel2Display.update = false;
        vm.fetchLevel2Display.new = false;
        vm.fetchLevel2Display.update = false;
    };

    // Implement all lists manually as a hack. In the later stage learn about graph theory and implement it as graphs. https://github.com/chenglou/data-structures

    vm.level1 = ['publish', 'fetch']
    vm.level1Display = {
        'publish': false,
        'fetch': false
    }

    vm.publishLevel2 = ['new', 'update']
    vm.publishLevel2Display = {
        'new': false,
        'update': false
    }

    vm.fetchLevel2 = ['new', 'update']
    vm.fetchLevel2Display = {
        'new': false,
        'update': false
    }

    vm.updateLevelsLists = function (list) {
        var list = list;
        falseAllItems();
        switch (list[0]) {
            case 1:
                vm.level1Display.publish = true;
                break;
            case 2:
                vm.level1Display.fetch = true;
                break;
        };
        switch (list[1]) {
            case 3:
                vm.publishLevel2Display.new = true;
                break;
            case 4:
                vm.publishLevel2Display.update = true;
                break;
            case 5:
                vm.fetchLevel2Display.new = true;
                break;
            case 6:
                vm.fetchLevel2Display.update = true;
                break;
        };
    };

    vm.debug = JSON.stringify(vm.level1Display) + JSON.stringify(vm.publishLevel2Display) + JSON.stringify(vm.fetchLevel2Display);

    //vm.menusListBools = Object.map(function (item) { return [ key, vm.currentFooterMenu[key] ] }).slice(1);

    vm.menus = {
        'publish': 
            {
                'visible': true,
                'new': 
                    {
                        'visible': true,
                    },
                'update': false
            },
        'fetch':
            {
                'visible': false,
                'new':
                    {
                        'visible': true,
                    },
                'update': false
            }
        };

    vm.menusDisplay = [
        'publish',
        'new'
    ];

    vm.currentLevel = 1;

    vm.currentFooterMenu = vm.menus[vm.menusDisplay[vm.currentLevel - 1]];
    vm.currentFooterMenu = Object.keys(vm.currentFooterMenu).map(function (key) { return [ key, vm.currentFooterMenu[key] ] }).slice(1);

    //vm.debug = vm.currentFooterMenu;

    vm.display =
        {
            publishNew: false,
            publishUpdate: true,
            fetchNew: false,
            fetchUpdate: false,
            bartosz: false
        };
    
    vm.footerDisplay =
        {
            level2: false,
            level1: true
        };

    vm.adjustFooterHeight = function () {
        var x = 0;
        for (var item in vm.footerDisplay) {
            if (vm.footerDisplay[item] == true) {
                x++;
            }
        };
        vm.debug = x;
        $('#content-footer').height(35 * x);
    };

    vm.hideAll = function (one) {
        //vm.display['publishUpdate'] = false;
        for (var item in vm.display) {
            vm.display[item] = false;
        };
        vm.display[one] = true;
    };
    /*
        angular.forEach(display, function (value, key) {
            value = false;
        }, log);



    vm.addShow = function () {
        vm.shows.push(vm.new);
        vm.new = {};
    };
    */

    vm.janusz = true;
    vm.title = 'AngularJS Tutorial example';
    vm.searchInput = '';
    vm.shows = [
    {
        title: 'Game of Thrones',
        year: 2011,
        favorite: true
    },
    {
        title: 'Walking Dead',
        year: 2010,
        favorite: false
    },
    {
        title: 'Firefly',
        year: 2002,
        favorite: true
    },
    {
        title: 'Banshee',
        year: 2013,
        favorite: true
    },
    {
        title: 'Greys Anatomy',
        year: 2005,
        favorite: false
    }
    ];

    vm.orders = [
    {
        id: 1,
        title: 'Year Ascending',
        key: 'year',
        reverse: false
    },
    {
        id: 2,
        title: 'Year Descending',
        key: 'year',
        reverse: true
    },
    {
        id: 3,
        title: 'Title Ascending',
        key: 'title',
        reverse: false
    },
    {
        id: 4,
        title: 'Title Descending',
        key: 'title',
        reverse: true
    }
    ];

    vm.new = {};
    vm.addShow = function () {
        vm.shows.push(vm.new);
        vm.new = {};
    };

});