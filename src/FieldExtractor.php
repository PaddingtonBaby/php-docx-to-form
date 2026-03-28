<?php

declare(strict_types=1);

namespace DocxCard;

class FieldExtractor
{
    public function extractPositions(string $pdfPath, array $fields): array
    {
        $cmd = sprintf('pdftotext -bbox -f 1 -l 1 %s - 2>&1', escapeshellarg($pdfPath));
        $bboxXml = shell_exec($cmd);

        $sxml = @simplexml_load_string($bboxXml);
        if (!$sxml) return [];

        $sxml->registerXPathNamespace('xhtml', 'http://www.w3.org/1999/xhtml');
        $pages = $sxml->xpath('//xhtml:page');
        if (empty($pages)) return [];

        $pageW = (float)$pages[0]['width'];
        $pageH = (float)$pages[0]['height'];
        if ($pageW <= 0 || $pageH <= 0) return [];

        $fieldByNameType = [];
        $fieldByName = [];
        foreach ($fields as $i => $f) {
            $name = trim((string)($f['name'] ?? ''));
            $type = trim((string)($f['type'] ?? ''));
            if ($name === '') continue;
            $fieldByName[$name] = $i + 1;
            if ($type !== '') {
                $fieldByNameType[$name . '_' . $type] = $i + 1;
            }
        }

        $words = $sxml->xpath('//xhtml:word');
        if (empty($words)) return [];

        $markers = [];
        $seen = [];
        $duplicatePositions = [];

        foreach ($words as $w) {
            $text = (string)$w;
            if (strpos($text, '{') === false) continue;
            if (preg_match('/[0-9A-F]{8}-/i', $text)) continue;
            if (!preg_match('/\{(\w+)/', $text, $m)) continue;

            $rawName = $m[1];
            $fieldIdx = null;

            if (strpos($text, ':') !== false && preg_match('/\{(\w+):/', $text, $mc)) {
                $nameBeforeColon = $mc[1];
                if (isset($fieldByNameType[$nameBeforeColon])) {
                    $fieldIdx = $fieldByNameType[$nameBeforeColon];
                }
                if ($fieldIdx === null && preg_match('/^(.+)_[a-z]+\d*$/', $nameBeforeColon, $np)) {
                    if (isset($fieldByName[$np[1]])) {
                        $fieldIdx = $fieldByName[$np[1]];
                    }
                }
            }

            if ($fieldIdx === null && strpos($rawName, '_') !== false) {
                if (isset($fieldByNameType[$rawName])) {
                    $fieldIdx = $fieldByNameType[$rawName];
                }
                if ($fieldIdx === null) {
                    foreach ($fieldByNameType as $key => $idx) {
                        if (strpos($key, $rawName) === 0) {
                            $fieldIdx = $idx;
                            break;
                        }
                    }
                }
                if ($fieldIdx === null && preg_match('/^(.+?)_[a-z]/', $rawName, $np)) {
                    if (isset($fieldByName[$np[1]])) {
                        $fieldIdx = $fieldByName[$np[1]];
                    }
                }
            }

            if ($fieldIdx === null && strlen($rawName) >= 3) {
                $candidates = [];
                foreach ($fieldByName as $fname => $fIdx) {
                    if (strpos($fname, $rawName) === 0) {
                        $candidates[$fname] = $fIdx;
                    }
                }
                if (count($candidates) === 1) {
                    $fieldIdx = reset($candidates);
                }
            }

            if ($fieldIdx === null) continue;

            $pos = [
                'index' => $fieldIdx,
                'leftPct' => round((float)$w['xMin'] / $pageW * 100, 2),
                'topPct' => round((float)$w['yMin'] / $pageH * 100, 2),
            ];

            if (!isset($seen[$fieldIdx])) {
                $seen[$fieldIdx] = true;
                $markers[] = $pos;
            } else {
                $pos['duplicate'] = true;
                $duplicatePositions[] = $pos;
                $markers[] = $pos;
            }
        }

        // Если не находит поле  - берет соседний дубль. Сомнительно, но может помочь в случае разорванных переменных на несколько слов.
        $allFieldIndices = range(1, count($fields));
        $missing = array_diff($allFieldIndices, array_keys($seen));
        foreach ($missing as $missIdx) {
            foreach ($duplicatePositions as $dp) {
                if (abs($dp['index'] - $missIdx) <= 1) {
                    $markers[] = [
                        'index' => $missIdx,
                        'leftPct' => $dp['leftPct'],
                        'topPct' => $dp['topPct'],
                    ];
                    break;
                }
            }
        }

        return $markers;
    }
}
