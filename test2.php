<?php

$thumb = new Imagick();
$thumb->readImage('images/banana_shake.jpg');    $thumb->resizeImage(320,240,Imagick::FILTER_LANCZOS,1);
$thumb->writeImage('images/banana_shake_thumb.jpg');
$thumb->clear();
$thumb->destroy(); 

?>