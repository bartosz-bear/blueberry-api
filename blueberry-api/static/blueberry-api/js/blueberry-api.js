        $( document ).ready(function() {

            var auth_cookie = getCookie("auth");
            if (auth_cookie) {
                document.getElementById("login-logout").innerHTML = "Log out";
                document.getElementById("login-logout").href = "/display";
            }
            else {
                document.getElementById("login-logout").innerHTML = "Log in";
                document.getElementById("login-logout").href = "/login";
            }
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