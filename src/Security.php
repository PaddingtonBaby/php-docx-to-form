<?php

declare(strict_types=1);

namespace DocxCard;

use ZipArchive;

class Security
{
    public const MAX_UPLOAD_BYTES = 10485760;
    public const MAX_JSON_BYTES = 524288;
    private const MAX_ZIP_FILES = 250;
    private const MAX_ZIP_UNCOMPRESSED_BYTES = 83886080;
    private const MAX_ZIP_RATIO = 100;
    private const TEMPLATE_TTL_SECONDS = 86400;
    private const RATE_WINDOW_SECONDS = 600;

    public static function storageDir(string $name): string
    {
        $baseDir = dirname(__DIR__, 2) . '/php-docx-to-form-storage';
        $dir = $baseDir . '/' . $name;
        self::ensureDir($dir);
        return $dir;
    }

    public static function ensureDir(string $dir): void
    {
        if (is_dir($dir)) return;
        if (!mkdir($dir, 0750, true) && !is_dir($dir)) {
            throw new ApiException('Не удалось подготовить хранилище', 500);
        }
    }

    public static function cleanupStorage(): void
    {
        foreach (['templates', 'output', 'rate'] as $dirName) {
            $dir = self::storageDir($dirName);
            foreach (glob($dir . '/*') ?: [] as $path) {
                if (!is_file($path)) continue;
                if (filemtime($path) !== false && filemtime($path) < time() - self::TEMPLATE_TTL_SECONDS) {
                    @unlink($path);
                }
            }
        }
    }

    public static function rateLimit(string $scope, int $limit): void
    {
        $ip = $_SERVER['REMOTE_ADDR'] ?? 'unknown';
        $safeScope = preg_replace('/[^a-z0-9_-]/i', '', $scope) ?: 'api';
        $key = hash('sha256', $safeScope . '|' . $ip);
        $path = self::storageDir('rate') . '/' . $key . '.json';
        $now = time();

        $state = ['start' => $now, 'count' => 0];
        if (is_file($path)) {
            $loaded = json_decode((string)file_get_contents($path), true);
            if (is_array($loaded)) {
                $state['start'] = (int)($loaded['start'] ?? $now);
                $state['count'] = (int)($loaded['count'] ?? 0);
            }
        }

        if ($state['start'] <= $now - self::RATE_WINDOW_SECONDS) {
            $state = ['start' => $now, 'count' => 0];
        }

        $state['count']++;
        file_put_contents($path, json_encode($state), LOCK_EX);

        if ($state['count'] > $limit) {
            throw new ApiException('Слишком много запросов, попробуйте позже', 429);
        }
    }

    public static function requireContentLength(int $maxBytes): void
    {
        $length = (int)($_SERVER['CONTENT_LENGTH'] ?? 0);
        if ($length > $maxBytes) {
            throw new ApiException('Слишком большой запрос', 413);
        }
    }

    public static function uploadedDocx(string $field): array
    {
        if (empty($_FILES[$field]) || !is_array($_FILES[$field])) {
            throw new ApiException('Файл шаблона не загружен', 400);
        }

        $file = $_FILES[$field];
        $error = (int)($file['error'] ?? UPLOAD_ERR_NO_FILE);
        if ($error !== UPLOAD_ERR_OK) {
            throw new ApiException(self::uploadErrorMessage($error), $error === UPLOAD_ERR_INI_SIZE ? 413 : 400);
        }

        $tmpName = (string)($file['tmp_name'] ?? '');
        if ($tmpName === '' || !is_uploaded_file($tmpName)) {
            throw new ApiException('Файл шаблона не загружен', 400);
        }

        $size = (int)($file['size'] ?? 0);
        if ($size <= 0) {
            throw new ApiException('Пустой файл шаблона', 400);
        }
        if ($size > self::MAX_UPLOAD_BYTES) {
            throw new ApiException('Файл слишком большой. Максимум 10 МБ', 413);
        }

        $ext = strtolower(pathinfo((string)($file['name'] ?? ''), PATHINFO_EXTENSION));
        if ($ext !== 'docx') {
            throw new ApiException('Поддерживаются только .docx файлы', 415);
        }

        $finfo = new \finfo(FILEINFO_MIME_TYPE);
        $mime = (string)$finfo->file($tmpName);
        $allowedMimes = [
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/zip',
            'application/octet-stream',
        ];
        if (!in_array($mime, $allowedMimes, true)) {
            throw new ApiException('Неверный тип файла', 415);
        }

        self::validateDocxZip($tmpName);
        return $file;
    }

    public static function validateDocxZip(string $path): void
    {
        $zip = new ZipArchive();
        if ($zip->open($path) !== true) {
            throw new ApiException('Не удалось открыть DOCX', 400);
        }

        $totalSize = 0;
        $hasDocument = false;

        try {
            if ($zip->numFiles <= 0 || $zip->numFiles > self::MAX_ZIP_FILES) {
                throw new ApiException('Некорректный DOCX архив', 400);
            }

            for ($i = 0; $i < $zip->numFiles; $i++) {
                $stat = $zip->statIndex($i);
                if (!$stat) {
                    throw new ApiException('Некорректный DOCX архив', 400);
                }

                $name = (string)$stat['name'];
                if ($name === 'word/document.xml') {
                    $hasDocument = true;
                }
                if ($name === '' || str_contains($name, '..') || str_starts_with($name, '/') || preg_match('/^[a-z]:/i', $name)) {
                    throw new ApiException('Некорректный DOCX архив', 400);
                }

                $size = (int)$stat['size'];
                $compressedSize = max(1, (int)$stat['comp_size']);
                $totalSize += $size;

                if ($totalSize > self::MAX_ZIP_UNCOMPRESSED_BYTES) {
                    throw new ApiException('DOCX слишком большой после распаковки', 413);
                }
                if ($size > 1048576 && $size / $compressedSize > self::MAX_ZIP_RATIO) {
                    throw new ApiException('DOCX выглядит как поврежденный или опасный архив', 400);
                }
            }

            if (!$hasDocument) {
                throw new ApiException('Неверный DOCX: word/document.xml не найден', 400);
            }
        } finally {
            $zip->close();
        }
    }

    public static function readJsonBody(): array
    {
        self::requireContentLength(self::MAX_JSON_BYTES);
        $raw = file_get_contents('php://input');
        if ($raw === false || $raw === '') {
            throw new ApiException('Неверный JSON в теле запроса', 400);
        }
        if (strlen($raw) > self::MAX_JSON_BYTES) {
            throw new ApiException('Слишком большой запрос', 413);
        }

        $input = json_decode($raw, true);
        if (!is_array($input)) {
            throw new ApiException('Неверный JSON в теле запроса', 400);
        }

        return $input;
    }

    public static function normalizeFormData(mixed $data): array
    {
        if (!is_array($data)) {
            throw new ApiException('Неверные данные формы', 400);
        }

        $clean = [];
        foreach ($data as $key => $value) {
            if (!is_string($key) || !preg_match('/^[a-zA-Z][a-zA-Z0-9_]*$/', $key)) continue;
            if (is_array($value) || is_object($value)) continue;
            $text = trim((string)$value);
            if (mb_strlen($text) > 5000) {
                throw new ApiException('Слишком длинное значение поля', 413);
            }
            $clean[$key] = $text;
        }
        return $clean;
    }

    public static function downloadName(string $name): string
    {
        $base = preg_replace('/\.docx$/i', '', $name) ?: 'document';
        $base = preg_replace('/[\r\n"\\\/]+/', '_', $base) ?: 'document';
        return $base . '_filled.docx';
    }

    public static function logError(\Throwable $e): void
    {
        error_log('[docx-card] ' . $e->getMessage() . ' in ' . $e->getFile() . ':' . $e->getLine());
    }

    private static function uploadErrorMessage(int $error): string
    {
        return match ($error) {
            UPLOAD_ERR_INI_SIZE, UPLOAD_ERR_FORM_SIZE => 'Файл слишком большой. Максимум 10 МБ',
            UPLOAD_ERR_NO_FILE => 'Файл шаблона не загружен',
            default => 'Ошибка загрузки файла',
        };
    }
}
