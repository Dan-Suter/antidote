<?php

/*
 * PHP GD
 * resize an image using GD library
 */

// File and new size
//the original image has 800x600
$filename = 'images/banana_shake.jpg';
//the resize will be a percent of the original size
$percent = 0.5;

// Content type
header('Content-Type: image/jpeg');

// Get new sizes
list($width, $height) = getimagesize($filename);
$newwidth = $width * $percent;
$newheight = $height * $percent;

// Load
$thumb = imagecreatetruecolor($newwidth, $newheight);
$source = imagecreatefromjpeg($filename);

// Resize
imagecopyresized($thumb, $source, 0, 0, 0, 0, $newwidth, $newheight, $width, $height);

// Output and free memory
//the resized image will be 400x300
imagejpeg($thumb);
imagedestroy($thumb);
?>