<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use DocxCard\DocxConverter;
use DocxCard\ApiException;
use DocxCard\Security;

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    echo json_encode(['success' => false, 'error' => 'Только POST'], JSON_UNESCAPED_UNICODE);
    exit;
}

try {
    Security::cleanupStorage();
    Security::rateLimit('convert', 20);
    Security::requireContentLength(Security::MAX_UPLOAD_BYTES + 1048576);

    $deps = DocxConverter::checkDependencies();
    if (!$deps['ok']) {
        throw new ApiException('Сервис временно не готов к обработке документов', 503);
    }

    $uploadedFile = Security::uploadedDocx('template');
    $storageDir = Security::storageDir('templates');

    $templateId = bin2hex(random_bytes(8));
    $storedName = $templateId . '.docx';
    $storedPath = $storageDir . '/' . $storedName;

    if (!move_uploaded_file($uploadedFile['tmp_name'], $storedPath)) {
        throw new ApiException('Не удалось сохранить загруженный файл', 500);
    }

    $converter = new DocxConverter();
    $fields = $converter->analyzeTemplate($storedPath);

    if (empty($fields)) {
        @unlink($storedPath);
        throw new ApiException(
            'В шаблоне не найдено плейсхолдеров вида {variable}. '
            . 'Используйте паттерны типа {fieldName_text:Label} или {fieldName:Label}'
        );
    }

    $result = $converter->convert($storedPath, $fields);

    file_put_contents(
        $storageDir . '/' . $templateId . '.json',
        json_encode([
            'original_name' => $uploadedFile['name'],
            'fields' => $fields,
            'created_at' => date('c'),
        ], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT)
    );

    echo json_encode([
        'success' => true,
        'template_id' => $templateId,
        'original_name' => $uploadedFile['name'],
        'fields' => $fields,
        'pages' => $result['pages'],
        'debug' => $result['debug'],
    ], JSON_UNESCAPED_UNICODE);

} catch (ApiException $e) {
    http_response_code($e->statusCode());
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage(),
    ], JSON_UNESCAPED_UNICODE);
} catch (Throwable $e) {
    Security::logError($e);
    http_response_code(500);
    echo json_encode([
        'success' => false,
        'error' => 'Не удалось обработать документ',
    ], JSON_UNESCAPED_UNICODE);
}
