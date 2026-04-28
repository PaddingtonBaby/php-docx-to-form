<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use DocxCard\DocxConverter;
use DocxCard\ApiException;
use DocxCard\Security;

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
    Security::cleanupStorage();
    Security::rateLimit('fill', 60);

    $input = Security::readJsonBody();
    $templateId = $input['template_id'] ?? '';
    $data = Security::normalizeFormData($input['data'] ?? []);

    if (!preg_match('/^[a-f0-9]{16}$/', $templateId)) {
        throw new ApiException('Неверный template_id', 400);
    }

    $storageDir = Security::storageDir('templates');
    $templatePath = $storageDir . '/' . $templateId . '.docx';
    $metaPath = $storageDir . '/' . $templateId . '.json';

    if (!file_exists($templatePath) || !file_exists($metaPath)) {
        throw new ApiException('Шаблон не найден', 404);
    }

    $meta = json_decode(file_get_contents($metaPath), true);
    if (!is_array($meta)) {
        throw new ApiException('Шаблон поврежден', 400);
    }

    $fields = $meta['fields'] ?? [];

    if (empty($fields)) {
        throw new ApiException('Для этого шаблона не определены поля', 400);
    }

    $outputDir = Security::storageDir('output');

    $outputName = 'filled_' . $templateId . '_' . time() . '_' . bin2hex(random_bytes(4)) . '.docx';
    $outputPath = $outputDir . '/' . $outputName;

    $converter = new DocxConverter();
    $converter->fillTemplate($templatePath, $outputPath, $fields, $data);

    $downloadName = Security::downloadName((string)($meta['original_name'] ?? 'document'));
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment; filename="' . rawurlencode($downloadName) . '"; filename*=UTF-8\'\'' . rawurlencode($downloadName));
    header('Content-Length: ' . filesize($outputPath));
    readfile($outputPath);

    @unlink($outputPath);

} catch (ApiException $e) {
    header('Content-Type: application/json; charset=utf-8');
    http_response_code($e->statusCode());
    echo json_encode(['success' => false, 'error' => $e->getMessage()], JSON_UNESCAPED_UNICODE);
} catch (Throwable $e) {
    Security::logError($e);
    header('Content-Type: application/json; charset=utf-8');
    http_response_code(500);
    echo json_encode(['success' => false, 'error' => 'Не удалось заполнить документ'], JSON_UNESCAPED_UNICODE);
}
