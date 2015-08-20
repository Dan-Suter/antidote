<?php
        // Load the absolute server path to the directory the script is running in
        echo  $_SERVER["DOCUMENT_ROOT"]."\admin\connection.php</br>";
        define('PROJECT_ROOT',$_SERVER["DOCUMENT_ROOT"]."\admin");
        require(PROJECT_ROOT.'\connection.php');


        $fileDir = dirname(__FILE__);

        // Make sure we end with a slash
        if (substr($fileDir, -1) != '/') {
            $fileDir .= '/';
        }

        // Load the absolute server path to the document root
        $docRoot = $_SERVER['DOCUMENT_ROOT'];
        echo "Doc Root " . $docRoot . " </br>";
        // Make sure we end with a slash
        if (substr($docRoot, -1) != '/') {
            $docRoot .= '/';
        }

        // Remove docRoot string from fileDir string as subPath string
        $subPath = preg_replace('~' . $docRoot . '~i', '', $fileDir);

        // Add a slash to the beginning of subPath string
        $subPath = '/' . $subPath;          

        // Test subPath string to determine if we are in the web root or not
        if ($subPath == '/') {
            // if subPath = single slash, docRoot and fileDir strings were the same
            echo "We are running in the web foot folder of http://" . $_SERVER['SERVER_NAME'];
        } else {
            // Anyting else means the file is running in a subdirectory
            echo "We are running in the '" . $subPath . "' subdirectory of http://" . $_SERVER['SERVER_NAME'];
        }
?>