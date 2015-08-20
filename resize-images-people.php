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
$foldername="people";

$targetDir="C:\inetpub\wwwroot\antidote\images\\people\\";
//$fileName="Black beans.jpg";


$sql = "SELECT name,uid_people FROM people";
$result = $conn->query($sql);

if ($result->num_rows > 0) {
    // output data of each row
    while($row = $result->fetch_assoc()) {
    					$filePath="C:\inetpub\wwwroot\antidote\images\\people\original\\".$row["uid_people"].".jpg";
        			$fileName=$row["uid_people"].".jpg";
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
									$sql = "UPDATE food SET image_path = concat('/images/food/med/',name,'.jpg) WHERE uid_food ='".$fileName=$row["uid_food"]."';";
								}
								if ($foldername == "recipe") {
									$sql = "UPDATE recipes SET image = '/images/recipe/med/".$_REQUEST["newfilename"]."' WHERE uid_recipe ='".substr($_REQUEST["newfilename"], 0, 8)."';";
								}
							
							if ($conn->query($sql) === TRUE) {
							    echo "Record updated successfully";
							} else {
							    echo "Error updating record: " . $conn->error;
							}
    }
} else {
    echo "0 results";
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