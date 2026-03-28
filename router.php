<?php

$uri = parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);

if (preg_match('#^/api/(\w+)\.php$#', $uri, $m)) {
    $apiFile = __DIR__ . '/api/' . $m[1] . '.php';
    if (file_exists($apiFile)) {
        require $apiFile;
        return true;
    }
}

$publicFile = __DIR__ . '/public' . $uri;
if ($uri !== '/' && file_exists($publicFile) && is_file($publicFile)) {
    return false;
}

$index = __DIR__ . '/public/index.html';
if (file_exists($index)) {
    readfile($index);
    return true;
}

http_response_code(404);
echo '404 — Не найдено';
