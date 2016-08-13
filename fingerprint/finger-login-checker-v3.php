<?php

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
Session_Start();
$_SESSION["finger_id"] = "";
$_SESSION["device_id"] = "";
$_SESSION["activle_user"] = "";
$_SESSION["active_user_zh"] = "";
$_SESSION["note2"] = "";
$_SESSION["note3"] = "";
$_SESSION["note4"] = "";


$finger_id = $_POST['finger_id'];
$device_id = $_POST['device_id'];

$note2 = $_POST['note2'];
$note3 = $_POST['note3'];
$note4 = substr($_POST['note4'], 0, 32); // FIX, a invisible char at the end 

$reg_list=array(611,711,172,713,714,715);
if ( !in_array($finger_id,$reg_list)){
    echo "指紋編號 $finger_id 未註冊!";
    return;
}




if ($device_id != "abc987") {
    echo "wrong device";
    return;
}

/*
  jason <jason.hsu@skyrock-casting.com>,
  711 FC/孙永飞 <yf.sun@skyrockcasting.com>,
  712 wuwenqing <wenqing.wu@fulltech-metal.com>,
  713 wwy.wu <wwy.wu@fulltech-metal.com>,
  714 vicky <vicky.li@skyrock-casting.com>,
  715 富甲Harvey <harvey.zhu@skyrock-casting.com>,
 */


$nameList[611] = "mark.chen";
//$nameList[612] = "mark.chen";
//$nameList[613] = "mark.chen";
$nameZhList[611] = "陳炳陵";
//$nameZhList[612] = "陳炳陵-左手食指";
//$nameZhList[613] = "陳炳陵-左手無名指";
$step2md5[611] = "8f10d544832295d0d8751ea4e5fb1f8f";

//
$nameList[612] = "derrick.chen";
$nameZhList[612] = "陳宏宇";
$step2md5[612] = "1e43de448a714c36928334521b19179a"; //1e43de448a714c36928334521b19179a *612step1.txt





$nameList[711] = "yf.sun";
$nameList[712] = "wenqing.wu";
$nameList[713] = "wwy.wu";
$nameList[714] = "vicky.li";
$nameList[715] = "harvey.zhu";
$nameZhList[711] = "孙永飞";
$nameZhList[712] = "吳文清";
$nameZhList[713] = "吳文益";
$nameZhList[714] = "李兰英";
$nameZhList[715] = "朱中行";

//eb5a73fc4442263437d3687945403dd4 *711.jpg
//2112d8af0319abf39c10051e59eb52a0 *712.jpg
//1ced688cce2443602d19a5b6daf579fd *713.jpg
//4671cd6ec85b6df0a6d20392ab67f056 *714.jpg
//c6a60afa5562f2873d77a2eb0be1a891 *715.jpg
$step2md5[711] = "eb5a73fc4442263437d3687945403dd4";
$step2md5[712] = "2112d8af0319abf39c10051e59eb52a0";
$step2md5[713] = "1ced688cce2443602d19a5b6daf579fd";
$step2md5[714] = "4671cd6ec85b6df0a6d20392ab67f056";
$step2md5[715] = "c6a60afa5562f2873d77a2eb0be1a891";




$active_user = $nameList[$finger_id];
$active_user_zh = $nameZhList[$finger_id];
$active_step2md5 = $step2md5[$finger_id];

//if ($active_step2md5 == $note4) {
$check_result = strcmp($active_step2md5, $note4);
//echo strlen($active_step2md5)."<br>";
//echo strlen($note4)."<br>";
//$check_result=strcmp($note4, $note4);
//echo "xxxcheck_result=$check_result <br>";

if ($check_result !== 0) {
    //echo "wrong step2md5  <br>$active_step2md5<br>$note4 ";
    echo "Wrong step2md5!";

    return;
}

//    echo " step2md5 is checked! <br>$active_step2md5<br>$note4 ";
//else 
//    echo "wrong step2md5  <br>$active_step2md5<br>$note4 ";
//    return;
//}



if (strlen($active_user) > 0) {
    $_SESSION["finger_id"] = $finger_id;
    $_SESSION["active_user"] = $active_user;
    $_SESSION["active_user_zh"] = $active_user_zh;

//    echo "xxxok999 $note2 $note3 $note4";
//    echo "ok999|$note2|$note3";
    echo "ok999";
} else {
    echo "wrong finger";
}

//$pass=$_POST['pass'];
//echo "id=$id, name=$name, active_user=$active_user";

//echo "ok777";
