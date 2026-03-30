<?php

declare(strict_types=1);

namespace DocxCard;

use DOMDocument;
use DOMXPath;
use ZipArchive;
use RecursiveDirectoryIterator;
use RecursiveIteratorIterator;

class DocxConverter
{
    private string $libreOfficeBin;
    private int $jpegDpi;
    private int $timeoutSeconds;

    public function __construct(
        string $libreOfficeBin = 'soffice',
        int $jpegDpi = 200,
        int $timeoutSeconds = 45
    ) {
        $this->libreOfficeBin = $libreOfficeBin;
        $this->jpegDpi = $jpegDpi;
        $this->timeoutSeconds = $timeoutSeconds;
    }

    public function convert(string $templatePath, array $fields): array
    {
        if (!file_exists($templatePath)) {
            throw new \RuntimeException('Файл шаблона не найден: ' . $templatePath);
        }

        $tmpDir = sys_get_temp_dir() . '/docx_card_' . getmypid() . '_' . bin2hex(random_bytes(4));
        @mkdir($tmpDir, 0755, true);

        $loProfile = sys_get_temp_dir() . '/docx_card_lo_' . getmypid() . '_' . bin2hex(random_bytes(4));
        @mkdir($loProfile, 0755, true);

        $debug = [
            'template' => basename($templatePath),
            'fields_count' => count($fields),
            'steps' => [],
        ];

        try {
            $rawVariables = $this->extractRawVariables($templatePath);
            $debug['raw_variables_found'] = $rawVariables;
            $debug['steps'][] = 'Извлечено ' . count($rawVariables) . ' плейсхолдеров из DOCX (XML)';

            $cleanDocx = $this->hideVariablesInDocx($templatePath, $tmpDir);
            $debug['steps'][] = 'Создание DOCX файла';

            $pdfPath = $this->convertToPdf($cleanDocx, $tmpDir, $loProfile);
            $debug['steps'][] = 'Конвертация в PDF через LibreOffice';

            $extractor = new FieldExtractor();
            $markersByPage = $extractor->extractPositions($pdfPath, $fields);
            $totalMarkers = array_sum(array_map('count', $markersByPage));
            $debug['markers_found'] = $totalMarkers;
            $debug['steps'][] = 'Извлечено ' . $totalMarkers . ' позиций маркеров на ' . count($markersByPage) . ' страницах';

            $images = $this->generateJpegs($pdfPath, $tmpDir);
            $debug['steps'][] = 'Сгенерировано ' . count($images) . ' превью JPEG (' . $this->jpegDpi . ' DPI)';

            $mapped = [];
            foreach ($markersByPage as $pageMarkers) {
                foreach ($pageMarkers as $m) {
                    $mapped[$m['index']] = true;
                }
            }
            $unmapped = [];
            foreach ($fields as $i => $f) {
                if (!isset($mapped[$i + 1])) {
                    $unmapped[] = $f['name'] ?? ('field_' . ($i + 1));
                }
            }
            $debug['mapped_fields'] = count($mapped);
            $debug['unmapped_fields'] = $unmapped;

            $pages = [];
            foreach ($images as $pageIdx => $imageData) {
                $pages[] = [
                    'image' => $imageData,
                    'markers' => $markersByPage[$pageIdx] ?? [],
                ];
            }

            return [
                'pages' => $pages,
                'debug' => $debug,
            ];
        } finally {
            @exec('rm -rf ' . escapeshellarg($tmpDir));
            @exec('rm -rf ' . escapeshellarg($loProfile));
        }
    }

    public function extractRawVariables(string $templatePath): array
    {
        $zip = new ZipArchive();
        if ($zip->open($templatePath) !== true) {
            throw new \RuntimeException('Не удалось открыть DOCX');
        }

        $variables = [];
        $xmlFiles = ['word/document.xml'];

        for ($i = 0; $i < $zip->numFiles; $i++) {
            $name = $zip->getNameIndex($i);
            if (preg_match('/^word\/(header|footer)\d*\.xml$/', $name)) {
                $xmlFiles[] = $name;
            }
        }

        foreach ($xmlFiles as $xmlFile) {
            $content = $zip->getFromName($xmlFile);
            if ($content === false) continue;

            $plainText = strip_tags($content);
            if (preg_match_all('/\{(?![0-9A-F]{8}-)[^}]+\}/', $plainText, $matches)) {
                foreach ($matches[0] as $var) {
                    $variables[] = [
                        'raw' => $var,
                        'source' => $xmlFile,
                    ];
                }
            }

            $dom = new DOMDocument();
            @$dom->loadXML($content);
            $xpath = new DOMXPath($dom);
            $ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
            $xpath->registerNamespace('w', $ns);

            $paragraphs = $xpath->query('//w:p');
            foreach ($paragraphs as $para) {
                $tNodes = $xpath->query('.//w:t', $para);
                $fullText = '';
                foreach ($tNodes as $t) {
                    $fullText .= $t->textContent;
                }
                if (preg_match_all('/\{(?![0-9A-F]{8}-)[^}]+\}/', $fullText, $m)) {
                    foreach ($m[0] as $var) {
                        $found = false;
                        foreach ($variables as $v) {
                            if ($v['raw'] === $var && $v['source'] === $xmlFile) {
                                $found = true;
                                break;
                            }
                        }
                        if (!$found) {
                            $variables[] = [
                                'raw' => $var,
                                'source' => $xmlFile . ' (split across runs)',
                            ];
                        }
                    }
                }
            }
        }

        $zip->close();
        return $variables;
    }

    public function analyzeTemplate(string $templatePath): array
    {
        $rawVars = $this->extractRawVariables($templatePath);
        $fields = [];
        $seen = [];

        foreach ($rawVars as $rv) {
            $raw = $rv['raw'];
            $inner = trim($raw, '{}');

            $name = '';
            $type = 'text';
            $label = '';

            if (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)_([a-zA-Z]+):(.+)$/', $inner, $m)) {
                $name = $m[1];
                $type = $this->normalizeFieldType($m[2]);
                $label = trim($m[3]);
            }
            elseif (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*):(.+)$/', $inner, $m)) {
                $name = $m[1];
                $label = trim($m[2]);
            }
            elseif (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)_([a-zA-Z]+)$/', $inner, $m)) {
                $name = $m[1];
                $type = $this->normalizeFieldType($m[2]);
                $label = ucfirst($name);
            }
            elseif (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)$/', $inner, $m)) {
                $name = $m[1];
                $label = ucfirst($name);
            }

            if ($name === '' || isset($seen[$name])) continue;
            $seen[$name] = true;

            $fields[] = [
                'name' => $name,
                'type' => $type,
                'label' => $label,
                'required' => false,
                'raw_pattern' => $raw,
            ];
        }

        return $fields;
    }

    public function fillTemplate(string $templatePath, string $outputPath, array $fields, array $data): void
    {
        $processor = new \PhpOffice\PhpWord\TemplateProcessor($templatePath);
        $processor->setMacroOpeningChars('{');
        $processor->setMacroClosingChars('}');

        $phpwordVars = $processor->getVariables();

        $valueMap = [];
        foreach ($fields as $field) {
            $valueMap[$field['name']] = $data[$field['name']] ?? '';
        }

        foreach ($phpwordVars as $rawVar) {
            $plainText = preg_replace('/\s+/', ' ', trim(strip_tags($rawVar)));
            $fieldName = $this->resolveFieldName($plainText);
            if ($fieldName !== null && isset($valueMap[$fieldName])) {
                $processor->setValue($rawVar, $valueMap[$fieldName]);
            }
        }

        $phpwordTmp = $outputPath . '.phpword.docx';
        $processor->saveAs($phpwordTmp);

        // PhpWord пересобирает ZIP целиком и теряет часть метаданных оригинала.
        // Поэтому берется только вывод измененных XML и подставляются в исходник.
        // Хз почему так, надо понять, но иначе в итоговом DOCX пропадают картинки и стили, хотя в PHPWord они есть.
        copy($templatePath, $outputPath);

        $zSrc = new ZipArchive();
        $zDst = new ZipArchive();
        if ($zSrc->open($phpwordTmp) !== true || $zDst->open($outputPath) !== true) {
            @unlink($phpwordTmp);
            throw new \RuntimeException('Не удалось открыть DOCX для трансплантации');
        }

        for ($i = 0; $i < $zSrc->numFiles; $i++) {
            $name = $zSrc->getNameIndex($i);
            if (preg_match('/^word\/(document|header\d*|footer\d*)\.xml$/', $name)) {
                $content = $zSrc->getFromName($name);
                if ($content !== false) {
                    $zDst->addFromString($name, $content);
                }
            }
        }

        $zSrc->close();
        $zDst->close();
        @unlink($phpwordTmp);
    }

    private function resolveFieldName(string $inner): ?string
    {

        if (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)_([a-zA-Z]+):(.+)$/', $inner, $m)) {
            return $m[1];
        }

        if (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*):(.+)$/', $inner, $m)) {
            return $m[1];
        }

        if (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)_([a-zA-Z]+)$/', $inner, $m)) {
            return $m[1];
        }

        if (preg_match('/^([a-zA-Z][a-zA-Z0-9_]*)$/', $inner, $m)) {
            return $m[1];
        }

        return null;
    }

    private function normalizeFieldType(string $type): string
    {
        $map = [
            'text' => 'text',
            'textarea' => 'textarea',
            'date' => 'date',
            'number' => 'number',
            'phone' => 'phone',
            'tel' => 'phone',
            'num' => 'number',
            'str' => 'text',
            'string' => 'text',
        ];
        return $map[strtolower($type)] ?? 'text';
    }

    private function hideVariablesInDocx(string $templatePath, string $tmpDir): string
    {
        $extractDir = $tmpDir . '/docx_contents';
        @mkdir($extractDir, 0755, true);

        $zip = new ZipArchive();
        if ($zip->open($templatePath) !== true) {
            throw new \RuntimeException('Не удалось открыть DOCX шаблон');
        }
        $zip->extractTo($extractDir);
        $zip->close();

        $xmlPath = $extractDir . '/word/document.xml';
        if (!file_exists($xmlPath)) {
            throw new \RuntimeException('Неверный DOCX: word/document.xml не найден');
        }

        $xmlFiles = array_merge(
            [$xmlPath],
            glob($extractDir . '/word/header*.xml') ?: [],
            glob($extractDir . '/word/footer*.xml') ?: []
        );

        foreach ($xmlFiles as $xmlFile) {
            $this->setVariableFontColorWhite($xmlFile);
        }

        $outDocx = $tmpDir . '/clean.docx';
        $outZip = new ZipArchive();
        $outZip->open($outDocx, ZipArchive::CREATE | ZipArchive::OVERWRITE);
        $basePath = realpath($extractDir);

        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($basePath, RecursiveDirectoryIterator::SKIP_DOTS),
            RecursiveIteratorIterator::LEAVES_ONLY
        );
        foreach ($files as $file) {
            $filePath = $file->getRealPath();
            $relativePath = substr($filePath, strlen($basePath) + 1);
            $outZip->addFile($filePath, $relativePath);
        }
        $outZip->close();

        return $outDocx;
    }

    private function setVariableFontColorWhite(string $xmlFile): void
    {
        if (!file_exists($xmlFile)) return;

        $dom = new DOMDocument();
        $dom->loadXML(file_get_contents($xmlFile));
        $xpath = new DOMXPath($dom);
        $ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
        $xpath->registerNamespace('w', $ns);

        $changed = false;
        $paragraphs = $xpath->query('//w:p');

        foreach ($paragraphs as $para) {
            $runs = $xpath->query('.//w:r', $para);
            if ($runs->length === 0) continue;

            $segments = [];
            foreach ($runs as $run) {
                $tNodes = $xpath->query('w:t', $run);
                if ($tNodes->length === 0) continue;
                $segments[] = ['run' => $run, 'text' => $tNodes->item(0)->textContent];
            }

            $offset = 0;
            foreach ($segments as &$seg) {
                $seg['start'] = $offset;
                $seg['end'] = $offset + mb_strlen($seg['text']);
                $offset = $seg['end'];
            }
            unset($seg);

            $fullText = implode('', array_column($segments, 'text'));
            if (strpos($fullText, '{') === false) continue;
            if (!preg_match_all('/\{(?![0-9A-F]{8}-)[^}]+\}/', $fullText, $matches, PREG_OFFSET_CAPTURE)) continue;

            for ($m = 0; $m < count($matches[0]); $m++) {
                $charStart = mb_strlen(substr($fullText, 0, $matches[0][$m][1]));
                $charEnd = $charStart + mb_strlen($matches[0][$m][0]);

                foreach ($segments as &$seg) {
                    if ($seg['end'] <= $charStart || $seg['start'] >= $charEnd) continue;

                    $run = $seg['run'];
                    $rPrList = $xpath->query('w:rPr', $run);
                    if ($rPrList->length > 0) {
                        $rPr = $rPrList->item(0);
                    } else {
                        $rPr = $dom->createElementNS($ns, 'w:rPr');
                        $run->insertBefore($rPr, $run->firstChild);
                    }

                    $colorList = $xpath->query('w:color', $rPr);
                    if ($colorList->length > 0) {
                        $colorEl = $colorList->item(0);
                    } else {
                        $colorEl = $dom->createElementNS($ns, 'w:color');
                        $rPr->appendChild($colorEl);
                    }
                    $colorEl->setAttribute('w:val', 'FFFFFF');
                    $changed = true;
                }
                unset($seg);
            }
        }

        if ($changed) {
            file_put_contents($xmlFile, $dom->saveXML());
        }
    }

    private function convertToPdf(string $docxPath, string $tmpDir, string $loProfile): string
    {
        // т.к. LibreOffice требует HOME при запуске от www-data - /tmp — наш выходд.
        $cmd = sprintf(
            'HOME=/tmp timeout %d %s --headless --norestore --nofirststartwizard %s --convert-to pdf --outdir %s %s 2>&1',
            $this->timeoutSeconds,
            escapeshellarg($this->libreOfficeBin),
            escapeshellarg('-env:UserInstallation=file://' . $loProfile),
            escapeshellarg($tmpDir),
            escapeshellarg($docxPath)
        );

        exec($cmd, $output, $returnCode);

        $pdfPath = $tmpDir . '/' . pathinfo($docxPath, PATHINFO_FILENAME) . '.pdf';
        if ($returnCode !== 0 || !file_exists($pdfPath)) {
            throw new \RuntimeException(
                'Конвертация в PDF через LibreOffice завершилась с ошибкой (код ' . $returnCode . '): '
                . implode("\n", $output)
            );
        }

        return $pdfPath;
    }

    private function generateJpegs(string $pdfPath, string $tmpDir): array
    {
        $jpgBase = $tmpDir . '/preview';
        $cmd = sprintf(
            'pdftoppm -jpeg -r %d %s %s 2>&1',
            $this->jpegDpi,
            escapeshellarg($pdfPath),
            escapeshellarg($jpgBase)
        );

        exec($cmd, $output, $returnCode);

        $jpgFiles = glob($jpgBase . '-*.jpg');

        if (empty($jpgFiles) && file_exists($jpgBase . '.jpg')) {
            $jpgFiles = [$jpgBase . '.jpg'];
        }

        if (empty($jpgFiles)) {
            throw new \RuntimeException('Ошибка генерации JPEG: ' . implode("\n", $output));
        }

        sort($jpgFiles);

        $images = [];
        foreach ($jpgFiles as $jpgPath) {
            $images[] = 'data:image/jpeg;base64,' . base64_encode(file_get_contents($jpgPath));
        }

        return $images;
    }

    public static function checkDependencies(): array
    {
        $missing = [];

        exec('which soffice 2>/dev/null', $out1, $rc1);
        if ($rc1 !== 0) $missing[] = 'soffice (LibreOffice) — apt install libreoffice-core';

        exec('which pdftotext 2>/dev/null', $out2, $rc2);
        if ($rc2 !== 0) $missing[] = 'pdftotext — apt install poppler-utils';

        exec('which pdftoppm 2>/dev/null', $out3, $rc3);
        if ($rc3 !== 0) $missing[] = 'pdftoppm — apt install poppler-utils';

        return [
            'ok' => empty($missing),
            'missing' => $missing,
        ];
    }
}
