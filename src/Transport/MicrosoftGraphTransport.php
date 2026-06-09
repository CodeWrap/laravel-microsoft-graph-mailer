<?php

namespace Codewrap\MicrosoftGraphMailer\Transport;

use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Str;
use Stringable;
use Symfony\Component\Mailer\Exception\TransportException;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\AbstractTransport;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Component\Mime\MessageConverter;

class MicrosoftGraphTransport extends AbstractTransport implements Stringable
{
    protected const MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024;

    public function __construct(
        protected array $config,
    ) {
        parent::__construct();
    }

    protected function doSend(SentMessage $message): void
    {
        $email = MessageConverter::toEmail($message->getOriginalMessage());
        $user = $this->resolveUser($message);

        $payload = [
            'message' => array_filter([
                'subject' => $email->getSubject(),
                'body' => [
                    'contentType' => $email->getHtmlBody() ? 'HTML' : 'Text',
                    'content' => $email->getHtmlBody() ?? $email->getTextBody() ?? '',
                ],
                'toRecipients' => $this->formatRecipients($email->getTo()),
                'ccRecipients' => $this->formatRecipients($email->getCc()) ?: null,
                'bccRecipients' => $this->formatRecipients($email->getBcc()) ?: null,
                'replyTo' => $this->formatRecipients($email->getReplyTo()) ?: null,
                'internetMessageHeaders' => $this->formatCustomHeaders($email) ?: null,
                'attachments' => $this->formatAttachments($email) ?: null,
            ]),
            'saveToSentItems' => $this->config['save_to_sent_items'] ?? true,
        ];

        $url = sprintf(
            '%s/users/%s/sendMail',
            $this->getApiBaseUrl(),
            urlencode($user),
        );

        $response = $this->sendWithTokenRetry($url, $payload);

        $messageId = Str::uuid()->toString() . '@microsoft-graph';
        $message->getOriginalMessage()->getHeaders()->addHeader('X-Message-ID', $messageId);
        $message->getOriginalMessage()->getHeaders()->addHeader('X-Graph-Message-ID', $messageId);
    }

    protected function sendWithTokenRetry(string $url, array $payload): \Illuminate\Http\Client\Response
    {
        $token = $this->getAccessToken();

        try {
            $response = Http::withToken($token)
                ->acceptJson()
                ->timeout(30)
                ->post($url, $payload);
        } catch (\Exception $e) {
            throw new TransportException(
                sprintf('Request to Microsoft Graph API failed. Reason: %s.', $e->getMessage()),
                is_int($e->getCode()) ? $e->getCode() : 0,
                $e,
            );
        }

        if ($response->status() === 401) {
            $this->clearCachedToken();
            $token = $this->getAccessToken();

            try {
                $response = Http::withToken($token)
                    ->acceptJson()
                    ->post($url, $payload);
            } catch (\Exception $e) {
                throw new TransportException(
                    sprintf('Request to Microsoft Graph API failed. Reason: %s.', $e->getMessage()),
                    is_int($e->getCode()) ? $e->getCode() : 0,
                    $e,
                );
            }
        }

        if ($response->failed()) {
            throw new TransportException(
                sprintf(
                    'Microsoft Graph API returned %d. Error: %s (%s). Request ID: %s.',
                    $response->status(),
                    $response->json('error.message', 'Unknown error'),
                    $response->json('error.code', ''),
                    $response->header('request-id', 'n/a'),
                ),
                $response->status(),
            );
        }

        return $response;
    }

    protected function clearCachedToken(): void
    {
        $tenant = $this->config['tenant'] ?? 'common';
        $clientId = $this->config['client'] ?? '';

        Cache::forget(sprintf('msgraph-mailer:%s:%s', $tenant, $clientId));
        Cache::forget(sprintf('msgraph-mailer-ttl:%s:%s', $tenant, $clientId));
    }

    protected function resolveUser(SentMessage $message): string
    {
        if (! empty($this->config['user'])) {
            return $this->config['user'];
        }

        return $message->getEnvelope()->getSender()->getAddress();
    }

    protected function getAccessToken(): string
    {
        $tenant = $this->config['tenant'] ?? 'common';
        $clientId = $this->config['client'] ?? '';

        $cacheKey = sprintf('msgraph-mailer:%s:%s', $tenant, $clientId);

        $ttlKey = sprintf('msgraph-mailer-ttl:%s:%s', $tenant, $clientId);

        return Cache::remember($cacheKey, $this->getTokenCacheTtl($ttlKey), function () use ($tenant, $ttlKey) {
            try {
                $response = Http::asForm()
                    ->timeout(10)
                    ->post(
                        sprintf('%s/%s/oauth2/v2.0/token', $this->getAuthUrl(), $tenant),
                        [
                            'grant_type' => 'client_credentials',
                            'client_id' => $this->config['client'],
                            'client_secret' => $this->config['secret'],
                            'scope' => 'https://graph.microsoft.com/.default',
                        ],
                    );
            } catch (\Exception $e) {
                throw new TransportException(
                    sprintf('Failed to obtain Microsoft Graph access token. Reason: %s.', $e->getMessage()),
                    is_int($e->getCode()) ? $e->getCode() : 0,
                    $e,
                );
            }

            if ($response->failed()) {
                throw new TransportException(
                    sprintf(
                        'Failed to obtain Microsoft Graph access token. Status: %d.',
                        $response->status(),
                    ),
                    $response->status(),
                );
            }

            $data = $response->json();
            $expiresIn = $data['expires_in'] ?? 3600;
            $skew = config('microsoft-graph-mailer.token_cache_skew', 300);
            $ttl = max($expiresIn - $skew, 60);

            Cache::put($ttlKey, $ttl, $ttl);

            return $data['access_token'];
        });
    }

    protected function getTokenCacheTtl(string $ttlKey): int
    {
        return (int) Cache::get($ttlKey, 3300);
    }

    /**
     * @param  Address[]  $addresses
     * @return array<int, array{emailAddress: array{address: string, name?: string}}>
     */
    protected function formatRecipients(array $addresses): array
    {
        return array_map(fn (Address $address) => [
            'emailAddress' => array_filter([
                'address' => $address->getAddress(),
                'name' => $address->getName() ?: null,
            ]),
        ], $addresses);
    }

    protected function formatCustomHeaders(Email $email): array
    {
        $headers = [];

        foreach ($email->getHeaders()->all() as $header) {
            $name = $header->getName();

            if (! str_starts_with(strtolower($name), 'x-')) {
                continue;
            }

            $headers[] = [
                'name' => $name,
                'value' => $header->getBodyAsString(),
            ];
        }

        return $headers;
    }

    protected function formatAttachments(Email $email): array
    {
        $attachments = [];

        foreach ($email->getAttachments() as $attachment) {
            $headers = $attachment->getPreparedHeaders();
            $filename = $headers->getHeaderParameter('Content-Disposition', 'filename') ?? 'attachment';
            $body = $attachment->getBody();

            if (strlen($body) > self::MAX_ATTACHMENT_BYTES) {
                throw new TransportException(
                    sprintf(
                        'Attachment "%s" exceeds the 3 MB limit for Microsoft Graph sendMail.',
                        $filename,
                    ),
                );
            }

            $item = [
                '@odata.type' => '#microsoft.graph.fileAttachment',
                'name' => $filename,
                'contentType' => $headers->get('Content-Type')->getBody(),
                'contentBytes' => base64_encode($body),
            ];

            $disposition = $headers->getHeaderBody('Content-Disposition');
            if ($disposition === 'inline') {
                $contentId = $attachment->getContentId();
                if ($contentId) {
                    $item['contentId'] = $contentId;
                    $item['isInline'] = true;
                }
            }

            $attachments[] = $item;
        }

        return $attachments;
    }

    protected function getApiBaseUrl(): string
    {
        return config('microsoft-graph-mailer.api_base_url', 'https://graph.microsoft.com/v1.0');
    }

    protected function getAuthUrl(): string
    {
        return config('microsoft-graph-mailer.auth_url', 'https://login.microsoftonline.com');
    }

    public function __toString(): string
    {
        return 'microsoft-graph';
    }
}
