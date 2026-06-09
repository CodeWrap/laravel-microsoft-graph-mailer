<?php

use Codewrap\MicrosoftGraphMailer\Transport\MicrosoftGraphTransport;
use Illuminate\Support\Facades\Cache;
use Illuminate\Support\Facades\Http;
use Symfony\Component\Mailer\Exception\TransportException;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;

function fakeTokenAndSend(): void
{
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response(null, 202),
    ]);
}

function buildEmail(): Email
{
    return (new Email())
        ->from('sender@example.com')
        ->to('recipient@example.com')
        ->subject('Test Subject')
        ->html('<p>Hello</p>');
}

function transport(array $config = []): MicrosoftGraphTransport
{
    return new MicrosoftGraphTransport(array_merge([
        'transport' => 'microsoft-graph',
        'tenant' => 'test-tenant',
        'client' => 'test-client',
        'secret' => 'test-secret',
        'save_to_sent_items' => true,
    ], $config));
}

function sendEmail(?Email $email = null, array $config = []): void
{
    transport($config)->send($email ?? buildEmail());
}

test('transport resolves from mail manager', function () {
    $transport = app('mail.manager')
        ->mailer('microsoft-graph')
        ->getSymfonyTransport();

    expect($transport)->toBeInstanceOf(MicrosoftGraphTransport::class);
    expect((string) $transport)->toBe('microsoft-graph');
});

test('sends email with correct payload', function () {
    fakeTokenAndSend();

    sendEmail();

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $body = $request->data();

        expect($body['message']['subject'])->toBe('Test Subject');
        expect($body['message']['body']['contentType'])->toBe('HTML');
        expect($body['message']['body']['content'])->toContain('Hello');
        expect($body['message']['toRecipients'][0]['emailAddress']['address'])->toBe('recipient@example.com');
        expect($body['saveToSentItems'])->toBeTrue();

        return true;
    });
});

test('sends to correct graph api url with sender', function () {
    fakeTokenAndSend();

    sendEmail();

    Http::assertSent(function ($request) {
        return str_contains($request->url(), '/users/sender%40example.com/sendMail');
    });
});

test('maps cc and bcc recipients', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->cc('cc@example.com');
    $email->bcc('bcc@example.com');

    sendEmail($email);

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $body = $request->data();

        expect($body['message']['ccRecipients'][0]['emailAddress']['address'])->toBe('cc@example.com');
        expect($body['message']['bccRecipients'][0]['emailAddress']['address'])->toBe('bcc@example.com');

        return true;
    });
});

test('maps reply-to', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->replyTo('reply@example.com');

    sendEmail($email);

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        expect($request->data()['message']['replyTo'][0]['emailAddress']['address'])->toBe('reply@example.com');

        return true;
    });
});

test('sends plain text when no html body', function () {
    fakeTokenAndSend();

    $email = (new Email())
        ->from('sender@example.com')
        ->to('recipient@example.com')
        ->subject('Plain')
        ->text('Just text');

    sendEmail($email);

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $body = $request->data();

        expect($body['message']['body']['contentType'])->toBe('Text');
        expect($body['message']['body']['content'])->toBe('Just text');

        return true;
    });
});

test('passes only x- custom headers', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->getHeaders()->addTextHeader('X-Custom-Tag', 'my-tag');
    $email->getHeaders()->addTextHeader('X-Request-ID', 'abc-123');

    sendEmail($email);

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $headers = $request->data()['message']['internetMessageHeaders'] ?? [];
        $names = array_column($headers, 'name');

        expect($names)->toContain('X-Custom-Tag');
        expect($names)->toContain('X-Request-ID');

        foreach ($names as $name) {
            expect(strtolower($name))->toStartWith('x-');
        }

        return true;
    });
});

test('excludes standard headers from custom headers', function () {
    fakeTokenAndSend();

    sendEmail();

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $headers = $request->data()['message']['internetMessageHeaders'] ?? [];
        $names = array_map(fn ($h) => strtolower($h['name']), $headers);

        expect($names)->not->toContain('from');
        expect($names)->not->toContain('to');
        expect($names)->not->toContain('subject');
        expect($names)->not->toContain('content-type');
        expect($names)->not->toContain('mime-version');

        return true;
    });
});

test('handles file attachment', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->attach('small content', 'test.txt', 'text/plain');

    sendEmail($email);

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'sendMail')) {
            return false;
        }

        $attachments = $request->data()['message']['attachments'] ?? [];

        expect($attachments)->toHaveCount(1);
        expect($attachments[0]['@odata.type'])->toBe('#microsoft.graph.fileAttachment');
        expect($attachments[0]['name'])->toBe('test.txt');
        expect($attachments[0]['contentType'])->toBe('text/plain');
        expect(base64_decode($attachments[0]['contentBytes']))->toBe('small content');

        return true;
    });
});

test('rejects attachment over 3mb', function () {
    fakeTokenAndSend();

    $email = buildEmail();
    $email->attach(str_repeat('x', 3 * 1024 * 1024 + 1), 'big.bin', 'application/octet-stream');

    expect(fn () => sendEmail($email))
        ->toThrow(TransportException::class, '3 MB limit');
});

test('token request uses correct credentials', function () {
    fakeTokenAndSend();

    sendEmail();

    Http::assertSent(function ($request) {
        if (! str_contains($request->url(), 'oauth2/v2.0/token')) {
            return false;
        }

        expect($request->url())->toContain('test-tenant');
        expect($request->data()['client_id'])->toBe('test-client');
        expect($request->data()['client_secret'])->toBe('test-secret');
        expect($request->data()['scope'])->toBe('https://graph.microsoft.com/.default');
        expect($request->data()['grant_type'])->toBe('client_credentials');

        return true;
    });
});

test('caches access token', function () {
    fakeTokenAndSend();

    sendEmail();
    sendEmail();

    $tokenRequests = Http::recorded(function ($request) {
        return str_contains($request->url(), 'oauth2/v2.0/token');
    });

    expect($tokenRequests)->toHaveCount(1);
});

test('retries with fresh token on 401', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fresh-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::sequence()
            ->push(null, 401)
            ->push(null, 202),
    ]);

    Cache::put('msgraph-mailer:test-tenant:test-client', 'stale-token', 3600);

    sendEmail();

    $sendRequests = Http::recorded(function ($request) {
        return str_contains($request->url(), 'sendMail');
    });

    expect($sendRequests)->toHaveCount(2);
});

test('throws on 403 forbidden', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => [
                'code' => 'ErrorAccessDenied',
                'message' => 'Access is denied.',
            ],
        ], 403),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'Access is denied');
});

test('throws on 500 server error', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => [
                'code' => 'InternalServerError',
                'message' => 'An internal error occurred.',
            ],
        ], 500),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'internal error');
});

test('error message does not contain secrets', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response([
            'access_token' => 'fake-token',
            'expires_in' => 3600,
        ]),
        'graph.microsoft.com/*' => Http::response([
            'error' => [
                'code' => 'Forbidden',
                'message' => 'Denied',
            ],
        ], 403),
    ]);

    try {
        sendEmail();
    } catch (TransportException $e) {
        expect($e->getMessage())->not->toContain('test-secret');
        expect($e->getMessage())->not->toContain('fake-token');

        return;
    }

    test()->fail('Expected TransportException was not thrown.');
});

test('throws on token fetch failure', function () {
    Http::fake([
        'login.microsoftonline.com/*' => Http::response(['error' => 'invalid_client'], 401),
    ]);

    expect(fn () => sendEmail())
        ->toThrow(TransportException::class, 'access token');
});

test('uses configured user for graph url', function () {
    fakeTokenAndSend();

    sendEmail(null, ['user' => 'mailer@contras.net']);

    Http::assertSent(function ($request) {
        return str_contains($request->url(), '/users/mailer%40contras.net/sendMail');
    });
});
