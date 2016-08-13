<?php
Session_Start();
$finger_id = $_SESSION["finger_id"];
$active_user = $_SESSION["active_user"];
$active_user_zh = $_SESSION["active_user_zh"];
?>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
    <head>
        <META http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <script src="js/jquery-1.10.2.js"></script>


    </head>
    <body>
        <center>

            <h1><u>應用入口</u></h1>

            <div style="font-size: 48pt">

                <?php // echo "session id is $finger_id <br>";  ?>
                <?php echo "登入帳號︰ $active_user "; ?>
                <?php echo "<br>登入用戶︰ $active_user_zh"; ?>

            </div>

        </center>
    </body>
</html>
