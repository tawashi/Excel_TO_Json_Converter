<?php

if ($argc < 2) {
    echo "Usage: php" . basename(__FILE__) . " filePath\n";
    exit(1);
}

$data = require $argv[1];

echo json_encode($data);
