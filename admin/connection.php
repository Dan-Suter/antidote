<?php
$servername = "localhost";
$username = "antidote2";
$password = "antidote";
$dbname = "antidote";

$conn = new mysqli($servername, $username, $password, $dbname);
// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
} 
?>