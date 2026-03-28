<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use DocxCard\DocxConverter;

header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(204);
    exit;
}

if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    header('Content-Type: application/json; charset=utf-8');
    http_response_code(405);
    echo json_encode(['success' => false, 'error' => 'Только POST'], JSON_UNESCAPED_UNICODE);
    exit;
}

try {
    $input = json_decode(file_get_contents('php://input'), true);
    if (!$input) {
        throw new RuntimeException('Неверный JSON в теле запроса');
    }

    $templateId = $input['template_id'] ?? '';
    $data = $input['data'] ?? [];

    if (!preg_match('/^[a-f0-9]{16}$/', $templateId)) {
        throw new RuntimeException('Неверный template_id');
    }

    $storageDir = __DIR__ . '/../storage/templates';
    $templatePath = $storageDir . '/' . $templateId . '.docx';
    $metaPath = $storageDir . '/' . $templateId . '.json';

    if (!file_exists($templatePath) || !file_exists($metaPath)) {
        throw new RuntimeException('Шаблон не найден');
    }

    $meta = json_decode(file_get_contents($metaPath), true);
    $fields = $meta['fields'] ?? [];

    if (empty($fields)) {
        throw new RuntimeException('Для этого шаблона не определены поля');
    }

    $outputDir = __DIR__ . '/../storage/output';
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0755, true);
    }

    $outputName = 'filled_' . $templateId . '_' . time() . '.docx';
    $outputPath = $outputDir . '/' . $outputName;

    $converter = new DocxConverter();
    $converter->fillTemplate($templatePath, $outputPath, $fields, $data);

    $downloadName = preg_replace('/\.docx$/i', '', $meta['original_name'] ?? 'документ') . '_filled.docx';
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . $downloadName . '"');
    header('Content-Length: ' . filesize($outputPath));
    readfile($outputPath);

    @unlink($outputPath);

} catch (Throwable $e) {
    header('Content-Type: application/json; charset=utf-8');
    http_response_code(400);
    echo json_encode(['success' => false, 'error' => $e->getMessage()], JSON_UNESCAPED_UNICODE);
}
