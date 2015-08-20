<?php
/**
 * upload.php
 *
 * Copyright 2013, Moxiecode Systems AB
 * Released under GPL License.
 *
 * License: http://www.plupload.com/license
 * Contributing: http://www.plupload.com/contributing
 */

#!! IMPORTANT: 
#!! this file is just an example, it doesn't incorporate any security checks and 
#!! is not recommended to be used in production environment as it is. Be sure to 
#!! revise it and customize to your needs.
define('PROJECT_ROOT',$_SERVER["DOCUMENT_ROOT"]);
require(PROJECT_ROOT.'\admin\connection.php');
require(PROJECT_ROOT.'\files\php-image-resize-master\src\ImageResize.php');

// Make sure file is not cached (as it happens for example on iOS devices)
header("Expires: Mon, 26 Jul 1997 05:00:00 GMT");
header("Last-Modified: " . gmdate("D, d M Y H:i:s") . " GMT");
header("Cache-Control: no-store, no-cache, must-revalidate");
header("Cache-Control: post-check=0, pre-check=0", false);
header("Pragma: no-cache");

/* 
// Support CORS
header("Access-Control-Allow-Origin: *");
// other CORS headers if any...
if ($_SERVER['REQUEST_METHOD'] == 'OPTIONS') {
	exit; // finish preflight CORS requests here
}
*/

// 5 minutes execution time
@set_time_limit(5 * 60);

// Uncomment this one to fake upload time
// usleep(5000);

// Settings
$foldername = $_REQUEST["folder_name"];
//echo($foldername);
$targetDir = ini_get("upload_tmp_dir") . DIRECTORY_SEPARATOR . "plupload";
$targetDir = 'C:\inetpub\wwwroot\antidote\images\\';
$targetDir .= $foldername;
$cleanupTargetDir = true; // Remove old files
$maxFileAge = 5 * 3600; // Temp file age in seconds


// Create target dir
if (!file_exists($targetDir)) {
	@mkdir($targetDir);
}

// Get a file name
//update to script dan 17/07/2015*********************************************************
//override file name with new filename
if (isset($_REQUEST["name"])) {
	$fileName = $_REQUEST["newfilename"];;
} elseif (!empty($_FILES)) {
	$fileName = $_FILES["file"]["name"];
} else {
	$fileName = uniqid("file_");
}
$filePath = $targetDir . '\original'. DIRECTORY_SEPARATOR . $fileName;

// Chunking might be enabled
$chunk = isset($_REQUEST["chunk"]) ? intval($_REQUEST["chunk"]) : 0;
$chunks = isset($_REQUEST["chunks"]) ? intval($_REQUEST["chunks"]) : 0;



// Remove old temp files	
if ($cleanupTargetDir) {
	if (!is_dir($targetDir) || !$dir = opendir($targetDir)) {
		die('{"jsonrpc" : "2.0", "error" : {"code": 100, "message": "Failed to open temp directory."}, "id" : "id"}');
	}

	while (($file = readdir($dir)) !== false) {
		$tmpfilePath = $targetDir . DIRECTORY_SEPARATOR . $file;

		// If temp file is current file proceed to the next
		if ($tmpfilePath == "{$filePath}.part") {
			continue;
		}

		// Remove temp file if it is older than the max age and is not the current file
		if (preg_match('/\.part$/', $file) && (filemtime($tmpfilePath) < time() - $maxFileAge)) {
			@unlink($tmpfilePath);
		}
	}
	closedir($dir);
}	


// Open temp file
if (!$out = @fopen("{$filePath}.part", $chunks ? "ab" : "wb")) {
	die('{"jsonrpc" : "2.0", "error" : {"code": 102, "message": "Failed to open output stream."}, "id" : "id"}');
}

if (!empty($_FILES)) {
	if ($_FILES["file"]["error"] || !is_uploaded_file($_FILES["file"]["tmp_name"])) {
		die('{"jsonrpc" : "2.0", "error" : {"code": 103, "message": "Failed to move uploaded file."}, "id" : "id"}');
	}

	// Read binary input stream and append it to temp file
	if (!$in = @fopen($_FILES["file"]["tmp_name"], "rb")) {
		die('{"jsonrpc" : "2.0", "error" : {"code": 101, "message": "Failed to open input stream."}, "id" : "id"}');
	}
} else {	
	if (!$in = @fopen("php://input", "rb")) {
		die('{"jsonrpc" : "2.0", "error" : {"code": 101, "message": "Failed to open input stream."}, "id" : "id"}');
	}
}

while ($buff = fread($in, 4096)) {
	fwrite($out, $buff);
}

@fclose($out);
@fclose($in);

// Check if file has been uploaded
if (!$chunks || $chunk == $chunks - 1) {
	// Strip the temp .part suffix off 
	rename("{$filePath}.part", $filePath);
}

if ($foldername == "food" || $foldername == "recipe") {
	$mini_img = new thumb;
	$mini_img->load($filePath); 
	$mini_img->resize(800,600); 
	$mini_img->save($targetDir."/xlarge/".$fileName);
	$mini_img->load($filePath); 
	$mini_img->resize(450,375); 
	$mini_img->save($targetDir."/large/".$fileName);
	$mini_img->load($filePath); 
	$mini_img->resize(225,188); 
	$mini_img->save($targetDir."/med/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(112,94);     
	$mini_img->save($targetDir."/small/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(62,46);   
	$mini_img->save($targetDir."/thumb/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(31,23);   
	$mini_img->save($targetDir."/xsthumb/".$fileName);
}
else
{//assume must be loading a person image
	$mini_img = new thumb;
	$mini_img->load($filePath); 
	$mini_img->resize(600,800); 
	$mini_img->save($targetDir."/xlarge/".$fileName);
	$mini_img->load($filePath); 
	$mini_img->resize(375,450); 
	$mini_img->save($targetDir."/large/".$fileName);
	$mini_img->load($filePath); 
	$mini_img->resize(188,225); 
	$mini_img->save($targetDir."/med/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(94,112);     
	$mini_img->save($targetDir."/small/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(46,62);   
	$mini_img->save($targetDir."/thumb/".$fileName);
	$mini_img->load($filePath);  
	$mini_img->resize(23,31);   
	$mini_img->save($targetDir."/xsthumb/".$fileName);	
}
 
//update database image path
	if ($foldername == "food") {
		$sql = "UPDATE food SET image_path = '/images/food/med/".$_REQUEST["newfilename"]."' WHERE uid_food ='".$_REQUEST["uid"]."';";
	}
	if ($foldername == "recipe") {
		$sql = "UPDATE recipes SET image = '/images/recipe/med/".$_REQUEST["newfilename"]."' WHERE uid_recipe ='".substr($_REQUEST["newfilename"], 0, 8)."';";
	}

if ($conn->query($sql) === TRUE) {
    echo "Record updated successfully";
} else {
    echo "Error updating record: " . $conn->error;
}

// Return Success JSON-RPC response
die('{"jsonrpc" : "2.0", "result" : null, "id" : "id"}');

class thumb{   
function load($img){   
$img_info = getimagesize($img);   
$img_type = $img_info[2];   
if($img_type == 1){   
$this->image = imagecreatefromgif($img);     
}  
elseif($img_type == 2){  
$this->image = imagecreatefromjpeg($img);    
}  
elseif($img_type == 3){  
$this->image = imagecreatefrompng($img);     
}  
}  
function get_height(){
return imagesy($this->image);   
}
function get_width(){  
return imagesx($this->image);   
}
function resize($width,$height){
$img_new = imagecreatetruecolor($width,$height);  
   imagecopyresampled($img_new,$this->image,0,0,0,0,$width,$height,$this->get_width(),$this->get_height());   
$this->image = $img_new;   
}
function save($img,$img_type = 'imagetype_jpeg'){
@$this->image_type = $img_info[2];   
if($img_type == 'imagetype_gif'){   
imagegif($this->image,$img);     
}  
elseif($img_type == 'imagetype_jpeg'){   
imagejpeg($this->image,$img);     
}  
elseif($img_type == 'imagetype_png'){   
imagepng($this->image,$img);     
}  
}
}