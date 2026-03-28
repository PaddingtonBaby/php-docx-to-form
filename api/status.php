<?php

declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

use DocxCard\DocxConverter;

header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');

$deps = DocxConverter::checkDependencies();

$phpwordInstalled = class_exists(\PhpOffice\PhpWord\TemplateProcessor::class);

echo json_encode([
    'ok' => $deps['ok'] && $phpwordInstalled,
    'php_version' => PHP_VERSION,
    'phpword_installed' => $phpwordInstalled,
    'dependencies' => [
        'soffice' => !in_array('soffice (LibreOffice) — apt install libreoffice-core', $deps['missing']),
        'pdftotext' => !in_array('pdftotext — apt install poppler-utils', $deps['missing']),
        'pdftoppm' => !in_array('pdftoppm — apt install poppler-utils', $deps['missing']),
    ],
    'missing' => $deps['missing'],
], JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT);
