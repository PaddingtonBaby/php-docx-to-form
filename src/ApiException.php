<?php

declare(strict_types=1);

namespace DocxCard;

class ApiException extends \RuntimeException
{
    private int $statusCode;

    public function __construct(string $message, int $statusCode = 400)
    {
        parent::__construct($message);
        $this->statusCode = $statusCode;
    }

    public function statusCode(): int
    {
        return $this->statusCode;
    }
}
