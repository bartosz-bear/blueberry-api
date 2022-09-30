       <script> 

    /*
        $("#btnSubmit2").click(function () {
            //collect userName and password entered by users
            //var username = $("#inputEmail").val();
            //var password = $("#inputPassword").val();

            auth(username, password);
        });
        */

        //authenticate function to make ajax call
        function httpRequest(destination) {
            var username = $("#inputEmail").val();
            var password = $("#inputPassword").val();
            var session_id = "none";

            if (destination == "/logging_testing") {
                session_id = $.cookie('session');
            }

            session_id = "none";

            data = {"username": username, "password": password};
            //data = JSON.stringify(data);

            /*
            $.ajax({
              type: "POST",
              url: destination,
              data: data,
              dataType : "json",
              complete: function(responseText) {
                var jsonResponse = JSON.parse(responseText);
                alert(jsonResponse);}
            });
            */

            var remote = $.ajax
            ({
                type: "POST",
                url: destination,
                dataType: 'json',
                async: false,
                data: data
            });

            alert(remote);

        }


        function showCookie() {
            alert($.cookie('session'));
        }

        function logOut() {
            $.get('/logout');
        }

        function testing() {
            $.get('/test');
        }

        // Hashing function from the previous example
        function make_base_auth(user, password) {
          var tok = user + ':' + pass;
          var hash = Base64.encode(tok);
          return "Basic " + hash;
        }

    </script>