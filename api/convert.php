<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use DocxCard\DocxConverter;

header('Content-Type: application/json; charset=utf-8');

// CORS нужен для локального фронта - на проде убрать или ограничить origin.
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
    $deps = DocxConverter::checkDependencies();
    if (!$deps['ok']) {
        throw new RuntimeException('Отсутствуют системные зависимости: ' . implode(', ', $deps['missing']));
    }

    if (empty($_FILES['template']) || $_FILES['template']['error'] !== UPLOAD_ERR_OK) {
        throw new RuntimeException('Файл шаблона не загружен или ошибка загрузки');
    }

    $uploadedFile = $_FILES['template'];
    $ext = strtolower(pathinfo($uploadedFile['name'], PATHINFO_EXTENSION));
    if ($ext !== 'docx') {
        throw new RuntimeException('Поддерживаются только .docx файлы');
    }

    $finfo = new finfo(FILEINFO_MIME_TYPE);
    $mime = $finfo->file($uploadedFile['tmp_name']);
    $allowedMimes = [
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'application/zip',
        'application/octet-stream',
    ];
    if (!in_array($mime, $allowedMimes, true)) {
        throw new RuntimeException('Неверный тип файла: ' . $mime);
    }

    $storageDir = __DIR__ . '/../storage/templates';
    if (!is_dir($storageDir)) {
        mkdir($storageDir, 0755, true);
    }

    $templateId = bin2hex(random_bytes(8));
    $storedName = $templateId . '.docx';
    $storedPath = $storageDir . '/' . $storedName;

    if (!move_uploaded_file($uploadedFile['tmp_name'], $storedPath)) {
        throw new RuntimeException('Не удалось сохранить загруженный файл');
    }

    $converter = new DocxConverter();
    $fields = $converter->analyzeTemplate($storedPath);

    if (empty($fields)) {
        @unlink($storedPath);
        throw new RuntimeException(
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

} catch (Throwable $e) {
    http_response_code(400);
    echo json_encode([
        'success' => false,
        'error' => $e->getMessage(),
    ], JSON_UNESCAPED_UNICODE);
}
